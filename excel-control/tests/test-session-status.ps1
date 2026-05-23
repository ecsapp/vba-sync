# session_status regression — start session, run one no-op command,
# then call session_status and assert the report carries the expected
# fields and meaningful values.

[CmdletBinding()]
param(
    [string]$SessionId = "test-status-$(Get-Random -Maximum 99999)"
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

Write-Host "Test: session-status (SessionId=$SessionId)" -ForegroundColor Cyan

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

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    $started = Wait-ForEvent $eventsFile 'started' 30
    if (-not $started) { throw "no 'started' event" }
    $excelPid = [int]$started.excel_pid
    Write-Host "  started, excel_pid=$excelPid" -ForegroundColor Green

    # Run one no-op so last_command + last_event have something to point to.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"calculate"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'calculated' 10 -IdMatch 'c1')) { throw "no calculated for c1" }

    # Now ask for status.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"session_status"}' -Encoding UTF8
    $r = Wait-ForEvent $eventsFile 'session_status_report' 10 -IdMatch 'c2'
    if (-not $r) { throw "no session_status_report for c2" }

    # Required fields + invariants on a healthy session.
    if ($r.report_version -ne 1) { throw "report_version expected 1, got '$($r.report_version)'" }
    if ($r.status -ne 'ready') { throw "status expected 'ready', got '$($r.status)'" }
    if (-not $r.host_alive)    { throw "host_alive false" }
    if (-not $r.excel_alive)   { throw "excel_alive false on a healthy session" }
    if ([int]$r.excel_pid -ne $excelPid) { throw "excel_pid mismatch (got $($r.excel_pid), expected $excelPid)" }
    if ([int]$r.host_pid -ne $proc.Id)   { throw "host_pid mismatch (got $($r.host_pid), expected $($proc.Id))" }
    if (-not $r.session_id)    { throw "session_id missing" }
    if (-not $r.started_at)    { throw "started_at missing" }
    if (-not $r.last_excel_check_ts) { throw "last_excel_check_ts missing (Test-ExcelAlive should have set it)" }

    if (-not $r.last_command -or $r.last_command.id -ne 'c2' -or $r.last_command.cmd -ne 'session_status') {
        throw "last_command should point at the session_status command itself: got id='$($r.last_command.id)' cmd='$($r.last_command.cmd)'"
    }
    if (-not $r.last_event) { throw "last_event missing" }
    # events_since_last_command counts events between the PREVIOUS command_ack
    # (c1 calculate) and this one (c2 session_status). Expect at least 1
    # (the calculated event); could be more if the watcher emitted anything.
    if ([int]$r.events_since_last_command -lt 1) {
        throw "events_since_last_command expected >=1 (calculated should be counted), got $($r.events_since_last_command)"
    }
    # idle_seconds: time since LastCommandTime was last rolled forward —
    # which happens on EVERY command dispatch including this session_status,
    # so should be near zero.
    if ([int]$r.idle_seconds -gt 5) {
        throw "idle_seconds unexpectedly high on an active session: $($r.idle_seconds)"
    }
    if ([int]$r.armed_responses_active -ne 0) {
        throw "armed_responses_active expected 0 (no rules armed), got $($r.armed_responses_active)"
    }
    if ($null -eq $r.stranger_excel_pids) { throw "stranger_excel_pids field missing" }
    if ($r.crashed_event) { throw "crashed_event should NOT be present when excel_alive=true" }

    Write-Host "  session_status_report: status=$($r.status) excel_alive=$($r.excel_alive) idle=$($r.idle_seconds)s events_since=$($r.events_since_last_command) armed=$($r.armed_responses_active)" -ForegroundColor Green

    # Arm one rule and verify armed_responses_active bumps.
    $armCmd = @{ id='a1'; cmd='arm_response'; match=@{ title='Microsoft Excel' }; button='OK' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $armCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'response_armed' 5 -IdMatch 'a1')) { throw "no response_armed for a1" }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"session_status"}' -Encoding UTF8
    $r2 = Wait-ForEvent $eventsFile 'session_status_report' 10 -IdMatch 'c3'
    if (-not $r2) { throw "no session_status_report for c3" }
    if ([int]$r2.armed_responses_active -ne 1) {
        throw "armed_responses_active expected 1 after arm_response, got $($r2.armed_responses_active)"
    }
    Write-Host "  armed_responses_active=1 after arm_response" -ForegroundColor Green

    # Close
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no 'closed' event" }
    Write-Host "  closed" -ForegroundColor Green

    $proc.WaitForExit(8000) | Out-Null
    if (-not $proc.HasExited) { throw "host did not exit cleanly" }

    Write-Host "PASS test-session-status" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-session-status: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
