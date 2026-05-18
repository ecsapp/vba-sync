# Phase 2 test — session host lifecycle and command_ack ordering.
#
# Spawns start-session.ps1 against empty.xlsm, appends a close command,
# waits for clean shutdown, asserts the expected event sequence.

[CmdletBinding()]
param(
    [string]$SessionId = "test-proto-$(Get-Random -Maximum 99999)"
)

$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$workbook   = Join-Path $here 'fixtures\empty.xlsm'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId

if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

# Kill any orphan Excel before we start so PIDs don't clash
. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: session protocol (SessionId=$SessionId)" -ForegroundColor Cyan

# Spawn the session host in the background. Use Start-Process so the
# script gets its own PID and a clean COM apartment.
$proc = Start-Process pwsh -ArgumentList @(
    '-NoProfile','-File', $startSess,
    '-Workbook', $workbook,
    '-SessionId', $SessionId,
    '-SessionsRoot', $sessions
) -PassThru -WindowStyle Hidden

try {
    # Wait for 'started' event
    $eventsFile = Join-Path $sessionDir 'events.jsonl'
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        if ((Test-Path $eventsFile) -and (Get-Content -LiteralPath $eventsFile -Raw) -match '"t":"started"') { break }
        Start-Sleep -Milliseconds 200
    }
    if (-not (Test-Path $eventsFile) -or -not ((Get-Content -LiteralPath $eventsFile -Raw) -match '"t":"started"')) {
        throw "Session did not emit 'started' within 30s"
    }
    Write-Host "  started event observed" -ForegroundColor Green

    # Verify state.json
    $state = Get-Content (Join-Path $sessionDir 'state.json') -Raw | ConvertFrom-Json
    if ($state.status -ne 'ready') { throw "state.status expected 'ready', got '$($state.status)'" }
    if ($state.pid -ne $proc.Id)   { throw "state.pid expected $($proc.Id), got '$($state.pid)'" }
    Write-Host "  state.json valid (pid=$($state.pid), status=ready)" -ForegroundColor Green

    # Append a close command
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"close"}' -Encoding UTF8

    # Wait for closed event
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        if ((Get-Content -LiteralPath $eventsFile -Raw) -match '"t":"closed"') { break }
        Start-Sleep -Milliseconds 200
    }
    if (-not ((Get-Content -LiteralPath $eventsFile -Raw) -match '"t":"closed"')) {
        throw "Session did not emit 'closed' within 15s of close command"
    }
    Write-Host "  closed event observed" -ForegroundColor Green

    # Verify event ordering
    $events = Get-Content -LiteralPath $eventsFile | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json }
    $types = $events | ForEach-Object { $_.t }
    $expected = @('started','command_ack','closing','closed')
    for ($i = 0; $i -lt $expected.Count; $i++) {
        if ($types[$i] -ne $expected[$i]) {
            throw "Event $i expected '$($expected[$i])', got '$($types[$i])' (full sequence: $($types -join ','))"
        }
    }
    Write-Host "  event ordering: $($types -join ' -> ')" -ForegroundColor Green

    # Wait for the host process to actually exit. Excel COM cleanup
    # can take 10+ seconds when prior tests in a run-all sweep left
    # COM handles around — give it room.
    $proc.WaitForExit(20000) | Out-Null
    if (-not $proc.HasExited) { throw "Session host process did not exit within 20s of 'closed' event" }
    Write-Host "  host process exited cleanly" -ForegroundColor Green

    Write-Host "PASS test-session-protocol" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-session-protocol: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events captured so far:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
