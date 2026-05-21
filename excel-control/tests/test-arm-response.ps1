# Phase 10 test — pre-armed responses (item 4b) against msgbox.xlsm.
#
# Arms an arm_response rule BEFORE running PopMsgBox: the watcher must
# auto-click the matching dialog with no respond_dialog round-trip and
# emit dialog_auto_responded. Also asserts every event carries the
# item-4a `ts` timestamp (checked on a main-loop event and a watcher
# event — the two emitters).

[CmdletBinding()]
param([string]$SessionId = "test-arm-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$workbook   = Join-Path $here 'fixtures\msgbox.xlsm'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: arm_response (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $workbook `
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
# ConvertFrom-Json coerces an ISO date string to [datetime], so verify the
# exact ms-precision UTC ts format on the raw event line (item 4a).
function Assert-EventTs([string]$path, [string]$type) {
    $line = Get-Content -LiteralPath $path | Where-Object { $_ -match "`"t`":`"$type`"" } | Select-Object -First 1
    if (-not $line) { throw "no '$type' event to check ts" }
    if ($line -notmatch '"ts":"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z"') {
        throw "'$type' event has no well-formed ts"
    }
}

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    $started = Wait-ForEvent $eventsFile 'started' 30
    if (-not $started) { throw "no started" }
    # item 4a — the start-session.ps1 emitter stamps every event
    Assert-EventTs $eventsFile 'started'
    Write-Host "  started (ts present)" -ForegroundColor Green

    # 1. Arm a response BEFORE the macro runs
    $armCmd = @{ id='a1'; cmd='arm_response'; match=@{ title='msgbox fixture' }; button='OK' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $armCmd -Encoding UTF8
    $armed = Wait-ForEvent $eventsFile 'response_armed' 10 -Id 'a1'
    if (-not $armed) {
        $f = Wait-ForEvent $eventsFile 'arm_failed' 2 -Id 'a1'
        throw ("no response_armed" + $(if ($f) { " (arm_failed: $($f.reason))" } else { '' }))
    }
    Write-Host "  response_armed kind=$($armed.kind) repeat=$($armed.repeat)" -ForegroundColor Green

    # 2. Run the macro — NO respond_dialog will be sent
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"run_macro","name":"PopMsgBox"}' -Encoding UTF8

    # 3. dialog_appeared is still emitted (audit trail) — and timestamped
    #    by the watcher emitter
    $dlg = Wait-ForEvent $eventsFile 'dialog_appeared' 20
    if (-not $dlg) { throw "no dialog_appeared" }
    # item 4a — the watcher emitter stamps too
    Assert-EventTs $eventsFile 'dialog_appeared'
    Write-Host "  dialog_appeared id=$($dlg.id) (audit trail kept, ts ok)" -ForegroundColor Green

    # 4. The armed rule must auto-respond — no agent round-trip
    $auto = Wait-ForEvent $eventsFile 'dialog_auto_responded' 10
    if (-not $auto) { throw "no dialog_auto_responded — armed rule did not fire" }
    if ($auto.rule_id -ne 'a1') { throw "dialog_auto_responded rule_id=$($auto.rule_id), expected a1" }
    if ($auto.button -ne 'OK')  { throw "dialog_auto_responded button=$($auto.button), expected OK" }
    if (-not $auto.ok)          { throw "dialog_auto_responded ok=$($auto.ok) — click failed" }
    Write-Host "  dialog_auto_responded rule_id=$($auto.rule_id) button=$($auto.button) — no round-trip" -ForegroundColor Green

    # 5. Macro completes
    if (-not (Wait-ForEvent $eventsFile 'macro_completed' 10 -Id 'c1')) { throw "no macro_completed" }
    Write-Host "  macro_completed" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-arm-response" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-arm-response: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
