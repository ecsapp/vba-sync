# Phase 6.8 test — list_macros / list_sheets / get_workbook_info.

[CmdletBinding()]
param([string]$SessionId = "test-int-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$workbook  = Join-Path $here 'fixtures\msgbox.xlsm'  # has modTest.PopMsgBox
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

Write-Host "Test: introspection (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-Process pwsh -ArgumentList @(
    '-NoProfile','-File', $startSess, '-Workbook', $workbook,
    '-SessionId', $SessionId, '-SessionsRoot', $sessions
) -PassThru -WindowStyle Hidden

function Read-Events($p) { if (-not (Test-Path $p)) { return @() }; Get-Content -LiteralPath $p | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json } }
function Wait-ForEvent($p, $T, $S=15, $Id=$null) {
    $end = (Get-Date).AddSeconds($S)
    while ((Get-Date) -lt $end) {
        foreach ($e in Read-Events $p) {
            if ($e.t -eq $T -and (-not $Id -or $e.id -eq $Id)) { return $e }
        }
        Start-Sleep -Milliseconds 200
    }
    return $null
}

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no started" }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"list_macros"}' -Encoding UTF8
    $m = Wait-ForEvent $eventsFile 'macros_listed' 15 -Id 'c1'
    if (-not $m) { throw "no macros_listed" }
    $popMsg = $m.macros | Where-Object { $_.name -eq 'PopMsgBox' }
    if (-not $popMsg) { throw "PopMsgBox not in macros list" }
    Write-Host "  list_macros: found $(($m.macros).Count) macro(s), PopMsgBox in modTest" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"list_sheets"}' -Encoding UTF8
    $s = Wait-ForEvent $eventsFile 'sheets_listed' 15 -Id 'c2'
    if (-not $s) { throw "no sheets_listed" }
    if (($s.sheets).Count -lt 1) { throw "no sheets reported" }
    Write-Host "  list_sheets: $(($s.sheets).Count) sheet(s), first=$(($s.sheets)[0].name)" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"get_workbook_info"}' -Encoding UTF8
    $w = Wait-ForEvent $eventsFile 'workbook_info' 15 -Id 'c3'
    if (-not $w) { throw "no workbook_info" }
    if ($w.info.sheet_count -lt 1) { throw "info.sheet_count invalid: $($w.info.sheet_count)" }
    if (-not $w.info.has_vba) { throw "info.has_vba expected true for msgbox.xlsm" }
    Write-Host "  get_workbook_info: name=$($w.info.name) sheets=$($w.info.sheet_count) has_vba=$($w.info.has_vba) size=$($w.info.size_bytes)" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-introspection" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-introspection: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    if ($proc -and -not $proc.HasExited) { try { $proc.Kill() } catch {} }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
