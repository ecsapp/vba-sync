# Phase 7 test — run_tests discovers Test_* Subs, runs them, reports
# pass/fail per test. Plus a P1.6 case: pre-flight compile_check
# refuses to start the test loop if any module has a compile error,
# emitting `tests_aborted` instead of freezing the command pump.

[CmdletBinding()]
param([string]$SessionId = "test-rt-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$fixture   = Join-Path $here 'fixtures\empty.xlsm'
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

# Stage a workbook with a 1-pass + 1-fail + 1-error module in the
# vba-sync sibling-folder convention so `import` can find it.
$modGood = @'
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
$stagedWb = New-StagedWorkbook -Fixture $fixture -Modules @{ modTests = $modGood }

Write-Host "Test: run_tests (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $stagedWb `
    -SessionId $SessionId -SessionsRoot $sessions

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

    # Import the test module via the vba-sync addin (sheet-safe).
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"import"}' -Encoding UTF8
    $imp = Wait-ForEvent $eventsFile 'import_completed' 60 -Id 'c1'
    if (-not $imp) {
        $f = Wait-ForEvent $eventsFile 'import_failed' 2 -Id 'c1'
        throw ('no import_completed' + $(if ($f) { " (import_failed: $($f.error))" } else { '' }))
    }

    # ----- happy path: discover + run -----
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

    # ----- P1.6: pre-flight compile_check rejects a broken module
    # with tests_aborted instead of freezing the pump.
    $modBroken = @'
Attribute VB_Name = "modBroken"
Option Explicit

Public Sub Test_BrokenSibling()
    Dim x As Long
    For x = 1 To 3
        If x = 1 Then Exit For
    ' deliberate: no Next — VBA "For without Next" compile error
End Sub
'@
    $brokenPath = Join-Path ([System.IO.Path]::GetDirectoryName($stagedWb)) `
        ([System.IO.Path]::GetFileNameWithoutExtension($stagedWb))
    Set-Content -LiteralPath (Join-Path $brokenPath 'Modules\modBroken.bas') `
        -Value $modBroken -Encoding ASCII

    # Re-import so the broken module is in the project alongside modTests.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"import"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'import_completed' 60 -Id 'c3')) {
        $f = Wait-ForEvent $eventsFile 'import_failed' 2 -Id 'c3'
        throw ('no import_completed (broken module) ' + $(if ($f) { " (import_failed: $($f.error))" } else { '' }))
    }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"run_tests"}' -Encoding UTF8
    $aborted = Wait-ForEvent $eventsFile 'tests_aborted' 30 -Id 'c4'
    if (-not $aborted) {
        # If we got tests_completed instead, the guard didn't fire — that's a regression.
        $tc = Wait-ForEvent $eventsFile 'tests_completed' 2 -Id 'c4'
        throw ("expected tests_aborted on broken-module pre-flight; " +
               $(if ($tc) { 'got tests_completed (P1.6 guard regressed)' } else { 'got nothing' }))
    }
    if ($aborted.reason -ne 'compile_error') {
        throw "tests_aborted.reason expected 'compile_error', got '$($aborted.reason)'"
    }
    Write-Host "  P1.6 ok — tests_aborted reason=$($aborted.reason) module=$($aborted.module) line=$($aborted.line)" -ForegroundColor Green

    # The pump must still be alive — send close and expect closed.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c5","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed — pump stuck after tests_aborted (P1.6 regressed)" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-run-tests (happy path + P1.6 pre-flight)" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-run-tests: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
