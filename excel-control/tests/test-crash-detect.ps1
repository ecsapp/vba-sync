# Crash detection regression — kills the session's Excel, asserts the
# harness notices and transitions cleanly.
#
# Covers:
#   - excel_crashed fires within a few seconds of Stop-Process on excel_pid
#   - state.json.status flips to "crashed"
#   - a subsequent non-close command emits command_rejected reason:excel_crashed
#   - close still works (closed event fires) — i.e. the host doesn't get stuck

[CmdletBinding()]
param(
    [string]$SessionId = "test-crash-$(Get-Random -Maximum 99999)"
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

Write-Host "Test: crash-detect (SessionId=$SessionId)" -ForegroundColor Cyan

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
    $stateFile    = Join-Path $sessionDir 'state.json'

    # 1. Wait for started, then grab the spawned Excel PID
    $started = Wait-ForEvent $eventsFile 'started' 30
    if (-not $started) { throw "no 'started' event" }
    $state = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json
    $excelPid = [int]$state.excel_pid
    if ($excelPid -le 0) { throw "no excel_pid in state.json" }
    Write-Host "  started, excel_pid=$excelPid" -ForegroundColor Green

    # 2. Kill the session's Excel — simulates the bug4b/bug4c crash mode
    Stop-Process -Id $excelPid -Force -ErrorAction Stop
    Write-Host "  killed excel_pid=$excelPid" -ForegroundColor Green

    # 3. excel_crashed should fire within a few seconds
    $crashed = Wait-ForEvent $eventsFile 'excel_crashed' 10
    if (-not $crashed) { throw "no 'excel_crashed' event after killing Excel" }
    if ([int]$crashed.pid -ne $excelPid) { throw "excel_crashed.pid=$($crashed.pid) does not match $excelPid" }
    if (-not $crashed.reason) { throw "excel_crashed.reason is empty" }
    Write-Host "  excel_crashed pid=$($crashed.pid) reason='$($crashed.reason)'" -ForegroundColor Green

    # 4. state.json.status flipped to crashed
    $deadline = (Get-Date).AddSeconds(5)
    $statusOk = $false
    while ((Get-Date) -lt $deadline) {
        $st = $null
        try { $st = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json } catch {}
        if ($st -and $st.status -eq 'crashed') { $statusOk = $true; break }
        Start-Sleep -Milliseconds 200
    }
    if (-not $statusOk) { throw "state.json.status did not flip to 'crashed'" }
    Write-Host "  state.json.status=crashed" -ForegroundColor Green

    # 5. A non-close command must be rejected with reason:excel_crashed
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"list_sheets"}' -Encoding UTF8
    $rej = Wait-ForEvent $eventsFile 'command_rejected' 10 -IdMatch 'c1'
    if (-not $rej) { throw "no 'command_rejected' event for c1" }
    if ($rej.reason -ne 'excel_crashed') { throw "command_rejected.reason='$($rej.reason)', expected 'excel_crashed'" }
    if ($rej.cmd -ne 'list_sheets') { throw "command_rejected.cmd='$($rej.cmd)', expected 'list_sheets'" }
    Write-Host "  command_rejected id=c1 cmd=list_sheets reason=excel_crashed" -ForegroundColor Green

    # 6. close still works
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no 'closed' event after close" }
    Write-Host "  closed" -ForegroundColor Green

    $proc.WaitForExit(8000) | Out-Null
    if (-not $proc.HasExited) { throw "host did not exit cleanly" }

    Write-Host "PASS test-crash-detect" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-crash-detect: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
