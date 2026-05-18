# Phase 7 test — run_tests discovers Test_* Subs, runs them, reports
# pass/fail per test.

[CmdletBinding()]
param([string]$SessionId = "test-rt-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$workbook  = Join-Path $here 'fixtures\empty.xlsm'
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

$srcDir = Join-Path $sessionDir '_src'
New-Item -ItemType Directory -Force -Path (Join-Path $srcDir 'Modules') | Out-Null

# Build a module with 1 pass + 1 fail + 1 error
$bas = @'
Attribute VB_Name = "modTests"
Option Explicit

Public Sub Test_Sales_Pass()
    ' simple pass — no Assert lib used
End Sub

Public Sub Test_Sales_Fail()
    Err.Raise vbObjectError + 1, "Test_Sales_Fail", "expected=5 got=3"
End Sub

Public Sub Test_Sales_Error()
    Dim a(1) As String
    a(99) = "x"
End Sub
'@
Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modTests.bas') -Value $bas -Encoding ASCII

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

Write-Host "Test: run_tests (SessionId=$SessionId)" -ForegroundColor Cyan

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

    # Sync the test module
    $syncCmd = @{ id='c1'; cmd='sync_vba'; source_dir=$srcDir } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $syncCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'sync_completed' 20 -Id 'c1')) { throw "no sync_completed" }

    # Run tests
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"run_tests"}' -Encoding UTF8
    $summary = Wait-ForEvent $eventsFile 'tests_completed' 60 -Id 'c2'
    if (-not $summary) { throw "no tests_completed" }

    $results = @(Read-Events $eventsFile | Where-Object { $_.t -eq 'test_result' })
    Write-Host "  ran $($summary.total) tests: passed=$($summary.passed) failed=$($summary.failed)" -ForegroundColor Green
    foreach ($r in $results) {
        Write-Host ("    {0,-25} status={1} dur={2}ms" -f $r.name, $r.status, $r.duration_ms)
    }
    if ($summary.total -ne 3) { throw "expected 3 tests, got $($summary.total)" }
    if ($summary.passed -ne 1) { throw "expected 1 pass, got $($summary.passed)" }
    if ($summary.failed -ne 2) { throw "expected 2 failures (raised err + runtime err), got $($summary.failed)" }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-run-tests" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-run-tests: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    if ($proc -and -not $proc.HasExited) { try { $proc.Kill() } catch {} }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
