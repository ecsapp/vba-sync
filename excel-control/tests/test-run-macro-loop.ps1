# run_macro_loop regression — three sub-tests in one session:
#   (1) happy path        — 3 iterations of a no-op macro, completed.
#   (2) stop_on            — macro fails on iteration 2, stop_on macro_failed
#                            exits cleanly with stop_at_iteration:2.
#   (3) {{i}} interpolation — between_iterations writes "row {{i}}" into A1
#                             each cycle; after the loop, A1 holds "row N-1"
#                             (between fires N-1 times for N iterations).

[CmdletBinding()]
param(
    [string]$SessionId = "test-loop-$(Get-Random -Maximum 99999)"
)

$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$fixture    = Join-Path $here 'fixtures\empty.xlsm'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

# Stage the test macros into the workbook's sibling folder so `import`
# (vba-sync addin-driven, sheet-safe) picks them up.
$bas = @'
Attribute VB_Name = "modLoopTest"
Option Explicit

Private mCallCount As Long

Public Sub LoopNoop()
End Sub

Public Sub LoopFailOnSecond()
    mCallCount = mCallCount + 1
    If mCallCount >= 2 Then
        Err.Raise 9999, "LoopFailOnSecond", "deliberate fail on call " & mCallCount
    End If
End Sub
'@
$stagedWb = New-StagedWorkbook -Fixture $fixture -Modules @{ modLoopTest = $bas }

Write-Host "Test: run_macro_loop (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $stagedWb `
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

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no 'started' event" }
    Write-Host "  started" -ForegroundColor Green

    # Import the test macros (LoopNoop + LoopFailOnSecond) from the
    # workbook's sibling folder via vba-sync's addin (sheet-safe).
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c0","cmd":"import"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'import_completed' 60 -IdMatch 'c0')) {
        $f = Wait-ForEvent $eventsFile 'import_failed' 2 -IdMatch 'c0'
        throw ('no import_completed' + $(if ($f) { " (import_failed: $($f.error))" } else { '' }))
    }
    Write-Host "  import_completed" -ForegroundColor Green

    # ---------- (1) happy path ----------
    $cmd1 = @{ id='c1'; cmd='run_macro_loop'; name='LoopNoop'; iterations=3 } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $cmd1 -Encoding UTF8
    $r1 = Wait-ForEvent $eventsFile 'macro_loop_completed' 30 -IdMatch 'c1'
    if (-not $r1) { throw "no macro_loop_completed for c1" }
    if ($r1.stop_reason -ne 'completed') { throw "happy-path stop_reason expected 'completed', got '$($r1.stop_reason)'" }
    if ([int]$r1.iterations_requested -ne 3) { throw "iterations_requested expected 3, got $($r1.iterations_requested)" }
    if ([int]$r1.iterations_completed -ne 3) { throw "iterations_completed expected 3, got $($r1.iterations_completed)" }
    if (@($r1.per_iteration).Count -ne 3)   { throw "per_iteration length expected 3, got $(@($r1.per_iteration).Count)" }
    if ($r1.stop_event) { throw "stop_event should be null on happy path" }
    # Progress events: at least 3 with id=c1
    $progress1 = @(Read-Events $eventsFile | Where-Object { $_.t -eq 'macro_loop_progress' -and $_.id -eq 'c1' })
    if ($progress1.Count -ne 3) { throw "macro_loop_progress count expected 3 for c1, got $($progress1.Count)" }
    Write-Host "  (1) happy path: 3/3 completed, $($progress1.Count) progress events" -ForegroundColor Green

    # ---------- (2) stop_on ----------
    # mCallCount is module-level so it's already 0 here (LoopNoop didn't touch it).
    # LoopFailOnSecond passes on call 1, fails on call 2.
    $cmd2 = @{ id='c2'; cmd='run_macro_loop'; name='LoopFailOnSecond'; iterations=5
               stop_on=@('macro_failed','runtime_error') } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $cmd2 -Encoding UTF8
    $r2 = Wait-ForEvent $eventsFile 'macro_loop_completed' 30 -IdMatch 'c2'
    if (-not $r2) { throw "no macro_loop_completed for c2" }
    if ($r2.stop_reason -ne 'stop_on') { throw "stop-on stop_reason expected 'stop_on', got '$($r2.stop_reason)'" }
    if ([int]$r2.iterations_completed -ne 2) { throw "iterations_completed expected 2, got $($r2.iterations_completed)" }
    if ([int]$r2.stop_at_iteration -ne 2) { throw "stop_at_iteration expected 2, got $($r2.stop_at_iteration)" }
    if (-not $r2.stop_event) { throw "stop_event missing" }
    if ($r2.stop_event.t -notin 'macro_failed','runtime_error') {
        throw "stop_event.t expected macro_failed or runtime_error, got '$($r2.stop_event.t)'"
    }
    Write-Host "  (2) stop_on: stopped at iteration $($r2.stop_at_iteration) on '$($r2.stop_event.t)'" -ForegroundColor Green

    # ---------- (3) {{i}} interpolation ----------
    # between_iterations writes "row {{i}}" into Sheet1!A1 between each
    # macro call. With iterations=4, between fires after i=1,2,3 (not
    # after the last). So A1 ends up as "row 3".
    $cmd3 = @{ id='c3'; cmd='run_macro_loop'; name='LoopNoop'; iterations=4
               between_iterations=@(
                 @{ cmd='write_range'; sheet='Sheet1'; range='A1'; values=@(@('row {{i}}')) }
               ) } | ConvertTo-Json -Compress -Depth 10
    Add-Content -LiteralPath $commandsFile -Value $cmd3 -Encoding UTF8
    $r3 = Wait-ForEvent $eventsFile 'macro_loop_completed' 30 -IdMatch 'c3'
    if (-not $r3) { throw "no macro_loop_completed for c3" }
    if ($r3.stop_reason -ne 'completed') { throw "interpolation stop_reason expected 'completed', got '$($r3.stop_reason)'" }
    if ([int]$r3.iterations_completed -ne 4) { throw "iterations_completed expected 4, got $($r3.iterations_completed)" }

    # Read A1 back to confirm {{i}} interpolated correctly on the last
    # between-iteration write (i=3).
    $readCmd = @{ id='c4'; cmd='read_range'; sheet='Sheet1'; range='A1' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $readCmd -Encoding UTF8
    $rr = Wait-ForEvent $eventsFile 'range_read' 10 -IdMatch 'c4'
    if (-not $rr) { throw "no range_read for c4" }
    $a1 = [string]$rr.rows[0][0]
    if ($a1 -ne 'row 3') { throw "A1 expected 'row 3' (last between-write at i=3), got '$a1'" }
    Write-Host "  (3) {{i}} interpolation: A1=$a1 (between fired at i=1,2,3 not after the last)" -ForegroundColor Green

    # Close
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c5","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no 'closed' event" }
    Write-Host "  closed" -ForegroundColor Green

    $proc.WaitForExit(8000) | Out-Null
    if (-not $proc.HasExited) { throw "host did not exit cleanly" }

    Write-Host "PASS test-run-macro-loop" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-run-macro-loop: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
