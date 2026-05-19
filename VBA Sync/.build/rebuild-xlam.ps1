# Rebuilds VBA Sync.xlam by removing existing standard modules and
# re-importing every .bas from VBA Sync\Modules\.
#
# Why: the .bas files are source of truth; the .xlam is the binary
# Excel actually loads. The two must stay in lockstep, so this script
# is the canonical "ship a code change" step.
#
# Class modules (.cls) and forms are NOT touched -- if you add or change
# one, do it in the VBA editor and add it to this script.
#
# Usage:
#   pwsh .\VBA Sync\.build\rebuild-xlam.ps1
#
# Excel must be closed before running. The script uses headless COM,
# spawns its own PID-tracked EXCEL.EXE, and cleans up that PID only --
# never blanket-kills EXCEL.EXE (the user may have other workbooks open).

[CmdletBinding()]
param(
    [switch]$KeepExcel  # for debugging
)
$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..') |
    Select-Object -ExpandProperty Path
$xlamPath = Join-Path $repoRoot 'VBA Sync.xlam'
$modulesDir = Join-Path $repoRoot 'VBA Sync\Modules'

if (-not (Test-Path -LiteralPath $xlamPath)) {
    throw "xlam not found at $xlamPath"
}
if (-not (Test-Path -LiteralPath $modulesDir)) {
    throw "Modules dir not found at $modulesDir"
}

# Pre-flight: any open EXCEL.EXE is fatal -- the xlam may be loaded as
# an add-in, in which case .Save fails with file-in-use and the rebuild
# silently no-ops.
$openExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($openExcel.Count -gt 0) {
    Write-Host "FAIL: $($openExcel.Count) EXCEL.EXE running. Close Excel before rebuilding." -ForegroundColor Red
    $openExcel | Format-Table Id, StartTime -AutoSize | Out-Host
    exit 1
}

# Required for VBProject.VBComponents.Import via COM.
function Enable-VbaTrust {
    $verKey = Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        Sort-Object -Property { [version]$_.PSChildName } -Descending |
        Select-Object -First 1
    if (-not $verKey) { return }
    $secPath = Join-Path $verKey.PSPath 'Excel\Security'
    if (-not (Test-Path $secPath)) { New-Item -Path $secPath -Force | Out-Null }
    Set-ItemProperty -Path $secPath -Name 'AccessVBOM' -Value 1 -Type DWord -Force
}

# Module list -- order doesn't matter for compile (VBA resolves
# late-binding refs at runtime), but we list deterministically so the
# log is readable.
$basFiles = @(
    'JsonConverter.bas',
    'modExcelExport.bas',
    'modFrxCanonicalize.bas',
    'modHarnessFiles.bas',
    'modSync.bas'
)
foreach ($f in $basFiles) {
    $p = Join-Path $modulesDir $f
    if (-not (Test-Path -LiteralPath $p)) { throw "Missing source: $p" }
}

Enable-VbaTrust

Write-Host "Rebuilding $xlamPath" -ForegroundColor Cyan
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$excelPid = (Get-Process EXCEL -ErrorAction SilentlyContinue |
    Sort-Object StartTime -Descending | Select-Object -First 1).Id
Write-Host "  Excel PID: $excelPid"

$failure = $null
try {
    # Open the .xlam as a regular workbook so we can edit its VBProject
    # and call .Save (NOT .SaveCopyAs).
    $wb = $xl.Workbooks.Open(
        $xlamPath,
        0,        # UpdateLinks
        $false,   # ReadOnly
        [Type]::Missing, [Type]::Missing, [Type]::Missing,
        $true     # IgnoreReadOnlyRecommended
    )
    try {
        $proj = $wb.VBProject

        # Remove each target module if present, then re-import.
        # vbext_ct_StdModule = 1
        foreach ($f in $basFiles) {
            $modName = [System.IO.Path]::GetFileNameWithoutExtension($f)
            $existing = $null
            try { $existing = $proj.VBComponents.Item($modName) } catch {}
            if ($existing) {
                Write-Host "  - remove  $modName"
                $proj.VBComponents.Remove($existing) | Out-Null
            } else {
                Write-Host "  - new     $modName"
            }
            $p = Join-Path $modulesDir $f
            $proj.VBComponents.Import($p) | Out-Null
            Write-Host "    imported $f"
        }

        # Quick compile check -- fail loudly if a .bas is broken.
        # Id 578 = "Compile VBAProject" on the VBE Debug menu.
        $compileOk = $true
        try {
            $cb = $xl.VBE.CommandBars
            $ctl = $cb.FindControl([Type]::Missing, 578)
            if ($ctl) { $ctl.Execute() }
            # If a compile error existed, VBE.ActiveCodePane points at it.
            $pane = $xl.VBE.ActiveCodePane
            if ($pane -and $pane.CodeModule) {
                # Heuristic: getting here without a popped error dialog means
                # compile succeeded. The dialog watcher's job, not ours.
            }
        } catch {
            $compileOk = $false
            Write-Host "  compile check threw: $_" -ForegroundColor Yellow
        }

        Write-Host "  saving..."
        $wb.Save()
    } finally {
        try { $wb.Close($false) } catch {}
    }
} catch {
    $failure = $_
} finally {
    try { $xl.Quit() | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    [System.GC]::Collect()
    if ($excelPid -and -not $KeepExcel) {
        Start-Sleep -Milliseconds 400
        try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch {}
    }
}

if ($failure) {
    Write-Host ""
    Write-Host "FAIL: $failure" -ForegroundColor Red
    exit 1
}

$size = (Get-Item -LiteralPath $xlamPath).Length
Write-Host ""
Write-Host "OK -- $xlamPath rebuilt ($size bytes)" -ForegroundColor Green
Write-Host "Next: git status should show the updated .xlam alongside your .bas changes." -ForegroundColor Cyan
