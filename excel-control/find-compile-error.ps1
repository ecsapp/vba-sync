# Pinpoints a VBA compile error by:
# 1. Opening the .xlam
# 2. Stripping + re-importing current .bas
# 3. Calling VBE 'Compile VBAProject'
# 4. After the modal dialog is dismissed, reading VBE.ActiveCodePane to learn
#    which module + line VBA tagged as broken.

[CmdletBinding()]
param(
    [string] $AddinPath  = 'C:\Users\ArnaudLavignolle\AppData\Roaming\Microsoft\AddIns\VBA Sync.xlam',
    [string] $SourceRoot = 'C:\Users\ArnaudLavignolle\AppData\Roaming\Microsoft\AddIns\VBA Sync'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'dialog-watcher.ps1')

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.EnableEvents = $false

$excelHwnd = [IntPtr]$xl.Hwnd
$excelPid  = Get-WindowProcessId -Hwnd $excelHwnd
$watcher = Start-DialogWatcher -ProcessId $excelPid

try {
    $addin = $xl.Workbooks.Open($AddinPath)
    $proj = $addin.VBProject

    $toRemove = @()
    foreach ($c in $proj.VBComponents) { if ($c.Type -ne 100) { $toRemove += $c.Name } }
    foreach ($name in $toRemove) { $proj.VBComponents.Remove($proj.VBComponents.Item($name)) }
    foreach ($dir in @('Modules','ClassModules','Forms')) {
        $d = Join-Path $SourceRoot $dir
        if (-not (Test-Path $d)) { continue }
        Get-ChildItem $d -File | ForEach-Object {
            if ($_.Extension -in @('.bas','.cls','.frm')) {
                $proj.VBComponents.Import($_.FullName) | Out-Null
            }
        }
    }

    # Compile
    Write-Host "Forcing compile..."
    try {
        $compileBtn = $xl.VBE.CommandBars.FindControl([type]::Missing, 578)
        if ($null -ne $compileBtn) { $compileBtn.Execute() }
    } catch {
        Write-Host "Compile call threw: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Start-Sleep -Milliseconds 1000
    $captured = $watcher.State.Captured
    if ($captured) {
        Write-Host ""
        Write-Host "Dialog content: $captured" -ForegroundColor Red
    } else {
        Write-Host "No dialog captured." -ForegroundColor Green
    }

    # VBE.ActiveCodePane should be positioned on the offending line.
    try {
        $vbe = $xl.VBE
        $pane = $vbe.ActiveCodePane
        if ($pane) {
            $moduleName = $pane.CodeModule.Name
            # GetSelection returns (startLine, startCol, endLine, endCol) via out-params
            # in COM. In PowerShell we have to read .TopLine and .CountOfVisibleLines.
            $topLine = $pane.TopLine
            Write-Host ""
            Write-Host "ActiveCodePane:" -ForegroundColor Cyan
            Write-Host "  Module: $moduleName"
            Write-Host "  TopLine: $topLine"

            # Use GetSelection via the late-bound COM call. Result is in the form of
            # 4 ByRef Int params; PowerShell handles via [ref].
            $sLine = 0; $sCol = 0; $eLine = 0; $eCol = 0
            [void]$pane.GetSelection([ref]$sLine, [ref]$sCol, [ref]$eLine, [ref]$eCol)
            Write-Host "  Selection: line $sLine col $sCol -> line $eLine col $eCol"

            # Dump 6 lines centered on the selection
            $start = [math]::Max(1, $sLine - 3)
            $end   = $sLine + 3
            Write-Host ""
            Write-Host "Source context:" -ForegroundColor Cyan
            for ($i = $start; $i -le $end; $i++) {
                $line = $pane.CodeModule.Lines($i, 1)
                $marker = if ($i -eq $sLine) { '>>' } else { '  ' }
                Write-Host "  $marker $($i): $line"
            }
        } else {
            Write-Host "No active code pane (compile may have succeeded)." -ForegroundColor Green
        }
    } catch {
        Write-Host "Couldn't read VBE state: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} finally {
    Stop-DialogWatcher $watcher | Out-Null
    try { $addin.Close($false) } catch {}
    try { $xl.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Start-Sleep -Milliseconds 500
    Get-Process -Id $excelPid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
