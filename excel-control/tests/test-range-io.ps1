# Phase 6.5 test — read_range + write_range.

[CmdletBinding()]
param([string]$SessionId = "test-rio-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$workbook  = Join-Path $here 'fixtures\empty.xlsm'
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: range I/O (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-Process pwsh -ArgumentList @(
    '-NoProfile','-File', $startSess, '-Workbook', $workbook,
    '-SessionId', $SessionId, '-SessionsRoot', $sessions
) -PassThru -WindowStyle Hidden

function Read-Events($p) {
    if (-not (Test-Path $p)) { return @() }
    Get-Content -LiteralPath $p | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json }
}
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

    # Write a 3x2 grid
    $writeCmd = @{
        id='c1'; cmd='write_range'; sheet='Sheet1'; range='A1'
        values = @( @('Name','Age'), @('Alice',30), @('Bob',25) )
    } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $writeCmd -Encoding UTF8
    $w = Wait-ForEvent $eventsFile 'range_written' 10 -Id 'c1'
    if (-not $w) { throw "no range_written" }
    Write-Host "  wrote $($w.rows)x$($w.cols) into $($w.range)" -ForegroundColor Green

    # Read it back
    $readCmd = @{ id='c2'; cmd='read_range'; sheet='Sheet1'; range='A1:B3' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $readCmd -Encoding UTF8
    $r = Wait-ForEvent $eventsFile 'range_read' 10 -Id 'c2'
    if (-not $r) { throw "no range_read" }
    if ($r.rows[0][0] -ne 'Name') { throw "A1 expected 'Name', got '$($r.rows[0][0])'" }
    if ($r.rows[1][1] -ne 30)     { throw "B2 expected 30, got '$($r.rows[1][1])'" }
    if ($r.rows[2][0] -ne 'Bob')  { throw "A3 expected 'Bob', got '$($r.rows[2][0])'" }
    Write-Host "  read back: $($r.rows[0][0])/$($r.rows[0][1]), $($r.rows[1][0])=$($r.rows[1][1]), $($r.rows[2][0])=$($r.rows[2][1])" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-range-io" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-range-io: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
