<#
.SYNOPSIS
  Headless harness for vba-sync Excel-export iteration.

.DESCRIPTION
  1. Starts a Win32 dialog watcher (from vba-dataset/tools pattern) that
     captures + dismisses any modal (MsgBox, VBA runtime error, system
     error) owned by our Excel PID.
  2. Launches a hidden Excel instance with macros disabled in opened
     workbooks (AutomationSecurityForceDisable).
  3. Rebuilds VBA Sync.xlam by re-importing current .bas/.cls/.frm
     sources into its VBProject and saving.
  4. Opens the target workbook read-only and calls
     modExcelExport.DoExportExcelStructure directly (bypasses
     DoExportProject which is wrapped in MsgBoxes).
  5. Prints phase + error from .export_error.log plus artefact summary.
  6. Cleans up Excel and the watcher.

.NOTES
  Exit codes:
    0 = export succeeded (MANIFEST.json present)
    1 = export failed (.export_error.log present)
    2 = export incomplete (no manifest, no log)
    3 = dialog captured (something modal happened)
    4 = harness error
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $Target,
    [string] $AddinPath  = 'C:\Users\ArnaudLavignolle\AppData\Roaming\Microsoft\AddIns\VBA Sync.xlam',
    [string] $SourceRoot = 'C:\Users\ArnaudLavignolle\AppData\Roaming\Microsoft\AddIns\VBA Sync'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'dialog-watcher.ps1')

if (-not (Test-Path $Target))    { throw "Target workbook not found: $Target" }
if (-not (Test-Path $AddinPath)) { throw "Addin not found: $AddinPath" }
if (-not (Test-Path $SourceRoot)) { throw "Source root not found: $SourceRoot" }

# Wipe stale export so we can detect whether THIS run produced output.
$rootPath  = (Split-Path $Target -Parent) + '\'
$exportDir = Join-Path $rootPath 'Excel'
if (Test-Path $exportDir) {
    Write-Host "Wiping stale export dir: $exportDir"
    Remove-Item -Recurse -Force $exportDir
}

$xl = $null
$wb = $null
$addin = $null
$watcher = $null
$exitCode = 4
$stage = 'init'
$startTime = Get-Date

try {
    $stage = 'create-app'
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    # NB: do NOT set AutomationSecurityForceDisable globally - it kills our .xlam macros too.
    # Instead, AskToUpdateLinks=False + EnableEvents=False is enough to keep the target's
    # Workbook_Open handler from firing.
    $xl.AskToUpdateLinks = $false
    $xl.EnableEvents = $false

    $excelHwnd = [IntPtr]$xl.Hwnd
    $excelPid  = Get-WindowProcessId -Hwnd $excelHwnd
    Write-Host "Excel PID: $excelPid"

    $stage = 'start-watcher'
    $watcher = Start-DialogWatcher -ProcessId $excelPid

    # ---- Rebuild .xlam from current source ----
    $stage = 'open-addin'
    Write-Host "Opening addin: $AddinPath"
    $addin = $xl.Workbooks.Open($AddinPath)

    $stage = 'probe-vbproject'
    try { $null = $addin.VBProject.VBComponents.Count } catch {
        throw "TRUST_CENTER: Enable 'Trust access to the VBA project object model' in Excel Trust Center."
    }

    $stage = 'strip-vba'
    $proj = $addin.VBProject
    $toRemove = @()
    foreach ($c in $proj.VBComponents) {
        # 1=StdModule, 2=ClassModule, 3=MSForm, 100=Document
        if ($c.Type -ne 100) { $toRemove += $c.Name }
    }
    foreach ($name in $toRemove) {
        Write-Host "  remove $name"
        $proj.VBComponents.Remove($proj.VBComponents.Item($name))
    }

    $stage = 'import-vba'
    foreach ($dir in @('Modules','ClassModules','Forms')) {
        $d = Join-Path $SourceRoot $dir
        if (-not (Test-Path $d)) { continue }
        Get-ChildItem $d -File | ForEach-Object {
            if ($_.Extension -in @('.bas','.cls','.frm')) {
                Write-Host "  import $($_.Name)"
                $proj.VBComponents.Import($_.FullName) | Out-Null
            }
        }
    }

    $stage = 'compile'
    try {
        # CommandBars control 578 = "Compile VBAProject". Surfaces compile errors as a dialog
        # which the watcher will capture.
        $compileBtn = $xl.VBE.CommandBars.FindControl([type]::Missing, 578)
        if ($null -ne $compileBtn) { $compileBtn.Execute() }
    } catch {
        throw "COMPILE_ERROR: $($_.Exception.Message)"
    }

    $stage = 'save-addin'
    Write-Host "Saving .xlam..."
    $addin.Save()

    # ---- Open target workbook ----
    $stage = 'open-target'
    Write-Host "Opening target: $Target"
    $wb = $xl.Workbooks.Open($Target, $null, $true)   # ReadOnly=True

    # ---- Run export ----
    $stage = 'run-export'
    $exportedDict = New-Object -ComObject Scripting.Dictionary
    Write-Host "Calling modExcelExport.DoExportExcelStructure..."
    try {
        $xl.Run("'VBA Sync.xlam'!modExcelExport.DoExportExcelStructure", $wb, $rootPath, $exportedDict)
        Write-Host "Run returned."
    } catch {
        Start-Sleep -Milliseconds 500
        $captured = $watcher.State.Captured
        if ($captured) { throw "DIALOG: $captured" }
        throw "RUN_ERROR: $($_.Exception.Message)"
    }

    # ---- Inspect results ----
    $stage = 'inspect'
    $errLogPath = Join-Path $exportDir '.export_error.log'
    $manifestPath = Join-Path $exportDir 'MANIFEST.json'

    if (Test-Path $errLogPath) {
        Write-Host ""
        Write-Host "FAILED. .export_error.log:" -ForegroundColor Red
        Get-Content $errLogPath | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        $exitCode = 1
    } elseif (Test-Path $manifestPath) {
        Write-Host ""
        Write-Host "SUCCESS." -ForegroundColor Green
        $exitCode = 0
    } else {
        Write-Host ""
        Write-Host "INCOMPLETE: no manifest, no error log." -ForegroundColor Yellow
        $exitCode = 2
    }

    if (Test-Path $exportDir) {
        $files = Get-ChildItem $exportDir -Recurse -File
        Write-Host "Artefacts: $($files.Count) files"
        $files | Select-Object -First 30 | ForEach-Object {
            $rel = $_.FullName.Substring($exportDir.Length + 1)
            Write-Host "  $rel ($([math]::Round($_.Length/1KB,1)) KB)"
        }
        if ($files.Count -gt 30) { Write-Host "  ... ($($files.Count - 30) more)" }
    }
}
catch {
    $msg = $_.Exception.Message
    if ($msg -like 'TRUST_CENTER:*') {
        Write-Host $msg -ForegroundColor Red
    } elseif ($msg -like 'COMPILE_ERROR:*') {
        Write-Host "Compile error: $($msg -replace '^COMPILE_ERROR:\s*','')" -ForegroundColor Red
        $exitCode = 3
    } elseif ($msg -like 'DIALOG:*') {
        Write-Host "Modal dialog captured during run:" -ForegroundColor Red
        Write-Host "  $($msg -replace '^DIALOG:\s*','')" -ForegroundColor Red
        $exitCode = 3
    } elseif ($msg -like 'RUN_ERROR:*') {
        Write-Host "COM Run error (no dialog): $($msg -replace '^RUN_ERROR:\s*','')" -ForegroundColor Red
        Write-Host "  stage: $stage" -ForegroundColor Yellow
        $exitCode = 4
    } else {
        Write-Host "Harness error at stage '$stage': $msg" -ForegroundColor Red
        $exitCode = 4
    }
}
finally {
    if ($watcher) {
        $finalCaptured = Stop-DialogWatcher $watcher
        if ($finalCaptured -and $exitCode -eq 0) {
            Write-Host "Note: dialog appeared during run but didn't break export: $finalCaptured" -ForegroundColor Yellow
        }
    }
    if ($wb)    { try { $wb.Close($false)    } catch {} }
    if ($addin) { try { $addin.Close($false) } catch {} }
    if ($xl)    {
        try { $xl.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
    # Final cleanup: kill any orphan Excel we own.
    if ($excelPid) {
        Get-Process -Id $excelPid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    $elapsed = (Get-Date) - $startTime
    Write-Host ""
    Write-Host "Elapsed: $([math]::Round($elapsed.TotalSeconds,1))s. Exit: $exitCode"
}

exit $exitCode
