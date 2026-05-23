# pre_crash_log regression — exercises the xcHarness_Log injection +
# auto-vacuum primitive. Four sub-tests in one session:
#   (1) instrument:{log:true} injects xcHarness_Log + emits
#       instrument_active{injected:true}. xcHarness_Log calls from the
#       agents own macro produce lines in macro.log.
#   (2) macro_failed -> pre_crash_log carries the recent xcHarness_Log
#       lines from THIS run.
#   (3) idempotent re-inject: a second instrument-flagged run_macro
#       sees the module already present (injected:false).
#   (4) successful run_macro with instrument leaves no pre_crash_log
#       event and clears the active record (next failure without
#       instrument must NOT emit pre_crash_log).

[CmdletBinding()]
param(
    [string]$SessionId = "test-precrash-$(Get-Random -Maximum 99999)"
)

$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$workbook   = Join-Path $here 'fixtures\empty.xlsm'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: pre_crash_log (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $workbook `
    -SessionId $SessionId -SessionsRoot $sessions

function Read-Events([string]$path) {
    if (-not (Test-Path $path)) { return @() }
    Get-Content -LiteralPath $path | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json }
}
function Wait-ForEvent($EventsPath, [string]$Type, [int]$TimeoutSec = 15, [string]$IdMatch = $null) {
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        foreach ($e in Read-Events $EventsPath) {
            if ($e.t -eq $Type) {
                if ($IdMatch -and $e.id -ne $IdMatch) { continue }
                return $e
            }
        }
        Start-Sleep -Milliseconds 200
    }
    return $null
}
function Count-Events($EventsPath, [string]$Type, [string]$IdMatch = $null) {
    $n = 0
    foreach ($e in Read-Events $EventsPath) {
        if ($e.t -eq $Type) {
            if ($IdMatch -and $e.id -ne $IdMatch) { continue }
            $n++
        }
    }
    return $n
}

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no 'started' event" }
    Write-Host "  started" -ForegroundColor Green

    # Seed two macros. LogAndOk: 3 xcHarness_Log calls (via Application.Run
    # so it compiles even before the module is injected — the agent flow
    # may compile_check before run_macro). LogAndFail: same shape but
    # raises after the third log line.
    $srcDir = Join-Path $sessionDir '_src'
    New-Item -ItemType Directory -Force -Path (Join-Path $srcDir 'Modules') | Out-Null
    $bas = @'
Attribute VB_Name = "modPreCrash"
Option Explicit

Public Sub LogAndOk()
    Application.Run "xcHarness_Log", "ok step 1"
    Application.Run "xcHarness_Log", "ok step 2"
    Application.Run "xcHarness_Log", "ok step 3"
End Sub

Public Sub LogAndFail()
    Application.Run "xcHarness_Log", "fail step A"
    Application.Run "xcHarness_Log", "fail step B"
    Application.Run "xcHarness_Log", "fail step C — about to err"
    Err.Raise 9999, "LogAndFail", "deliberate fail after 3 log lines"
End Sub
'@
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modPreCrash.bas') -Value $bas -Encoding ASCII

    Add-Content -LiteralPath $commandsFile -Value (@{ id='c0'; cmd='sync_vba'; source_dir=$srcDir } | ConvertTo-Json -Compress) -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'sync_completed' 15 -IdMatch 'c0')) { throw "no sync_completed" }
    Write-Host "  sync_completed" -ForegroundColor Green

    # Pre-arm: deliberate-fail produces a Microsoft Visual Basic dialog
    # that the watcher auto-Ends; on a stopped headless session this
    # raises a macro_failed via the COM dispatch path.

    # ---------- (1) inject + successful instrumented run ----------
    $cmd1 = @{ id='c1'; cmd='run_macro'; name='LogAndOk'
               instrument=@{ log=$true; tail_lines=200 } } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $cmd1 -Encoding UTF8

    $iAct1 = Wait-ForEvent $eventsFile 'instrument_active' 15 -IdMatch 'c1'
    if (-not $iAct1) { throw "no instrument_active for c1" }
    if (-not $iAct1.injected) { throw "first instrument should report injected:true (c1)" }
    if (-not $iAct1.log_file) { throw "instrument_active missing log_file" }
    Write-Host "  (1) instrument_active id=c1 injected=$($iAct1.injected) log=$($iAct1.log_file)" -ForegroundColor Green

    $r1 = Wait-ForEvent $eventsFile 'macro_completed' 30 -IdMatch 'c1'
    if (-not $r1) { throw "no macro_completed for c1" }
    # Successful run — must NOT emit pre_crash_log.
    if ((Count-Events $eventsFile 'pre_crash_log' 'c1') -ne 0) { throw "pre_crash_log fired on successful run (c1)" }
    # Log file must exist with 3 lines from LogAndOk.
    $logPath = [string]$iAct1.log_file
    if (-not (Test-Path -LiteralPath $logPath)) { throw "log file not created at $logPath" }
    $logLines = @(Get-Content -LiteralPath $logPath)
    if (@($logLines | Where-Object { $_ -match 'ok step 1' }).Count -ne 1) { throw "'ok step 1' line missing from log" }
    if (@($logLines | Where-Object { $_ -match 'ok step 3' }).Count -ne 1) { throw "'ok step 3' line missing from log" }
    Write-Host "  (1) success path: 3 xcHarness_Log lines in macro.log; no pre_crash_log emitted" -ForegroundColor Green

    # ---------- (2) idempotent re-inject + failing run + pre_crash_log ----------
    $cmd2 = @{ id='c2'; cmd='run_macro'; name='LogAndFail'
               instrument=@{ log=$true; tail_lines=200 } } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $cmd2 -Encoding UTF8

    $iAct2 = Wait-ForEvent $eventsFile 'instrument_active' 15 -IdMatch 'c2'
    if (-not $iAct2) { throw "no instrument_active for c2" }
    if ($iAct2.injected) { throw "second instrument should report injected:false (c2)" }
    Write-Host "  (2a) idempotent re-inject: instrument_active id=c2 injected=False" -ForegroundColor Green

    $r2 = Wait-ForEvent $eventsFile 'macro_failed' 30 -IdMatch 'c2'
    if (-not $r2) { throw "no macro_failed for c2" }
    $pcl2 = Wait-ForEvent $eventsFile 'pre_crash_log' 10 -IdMatch 'c2'
    if (-not $pcl2) { throw "no pre_crash_log for c2 failure" }
    if ($pcl2.trigger -ne 'macro_failed') { throw "pre_crash_log.trigger expected 'macro_failed', got '$($pcl2.trigger)'" }
    if ($pcl2.macro -ne 'LogAndFail') { throw "pre_crash_log.macro expected 'LogAndFail', got '$($pcl2.macro)'" }
    $pclLines = @($pcl2.log_lines)
    # The tail should include the 3 'fail step' lines we just emitted.
    if (@($pclLines | Where-Object { $_ -match 'fail step A' }).Count -ne 1) { throw "'fail step A' missing from pre_crash_log" }
    if (@($pclLines | Where-Object { $_ -match 'fail step C' }).Count -ne 1) { throw "'fail step C' missing from pre_crash_log" }
    Write-Host "  (2b) pre_crash_log id=c2 trigger=macro_failed lines_returned=$($pcl2.lines_returned) (incl. 'fail step A..C')" -ForegroundColor Green

    # Critical: pre_crash_log fires EXACTLY ONCE for this failure. (No
    # excel_crashed here — but the once-per-active-instrument invariant
    # still asserts on count.)
    $pclCount2 = Count-Events $eventsFile 'pre_crash_log' 'c2'
    if ($pclCount2 -ne 1) { throw "pre_crash_log fired $pclCount2 times for c2 (expected exactly 1)" }
    Write-Host "  (2c) pre_crash_log fired exactly once for c2" -ForegroundColor Green

    # ---------- (3) un-instrumented failure must NOT emit pre_crash_log ----------
    # The active record was cleared after c2's macro_failed; a fresh
    # un-instrumented run_macro that also fails should not emit
    # pre_crash_log (would mean LogVacuumActive leaked).
    $cmd3 = @{ id='c3'; cmd='run_macro'; name='LogAndFail' } | ConvertTo-Json -Compress  # no instrument
    Add-Content -LiteralPath $commandsFile -Value $cmd3 -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'macro_failed' 30 -IdMatch 'c3')) { throw "no macro_failed for c3" }
    Start-Sleep -Milliseconds 500   # give any stray vacuum a chance to emit before counting
    if ((Count-Events $eventsFile 'pre_crash_log' 'c3') -ne 0) {
        throw "pre_crash_log fired for c3 (un-instrumented) — LogVacuumActive leaked from c2"
    }
    Write-Host "  (3) un-instrumented c3 failure: no pre_crash_log (active record cleared cleanly)" -ForegroundColor Green

    # Close
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c9","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no 'closed' event" }
    Write-Host "  closed" -ForegroundColor Green

    $proc.WaitForExit(8000) | Out-Null
    if (-not $proc.HasExited) { throw "host did not exit cleanly" }

    Write-Host "PASS test-pre-crash-log" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-pre-crash-log: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
