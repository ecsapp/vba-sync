# Phase 3 test — run_macro + dialog round trip against msgbox.xlsm.
#
# Spawns start-session against msgbox.xlsm, fires PopMsgBox, expects
# dialog_appeared, responds OK, expects dialog_dismissed + macro_completed.

[CmdletBinding()]
param(
    [string]$SessionId = "test-msgbox-$(Get-Random -Maximum 99999)"
)

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

Write-Host "Test: msgbox flow (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-Process pwsh -ArgumentList @(
    '-NoProfile','-File', $startSess,
    '-Workbook', $workbook,
    '-SessionId', $SessionId,
    '-SessionsRoot', $sessions
) -PassThru -WindowStyle Hidden

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

    # 1. Wait for started
    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no 'started' event" }
    Write-Host "  started" -ForegroundColor Green

    # 2. Fire run_macro PopMsgBox
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"run_macro","name":"PopMsgBox"}' -Encoding UTF8

    # 3. Wait for dialog_appeared
    $dlg = Wait-ForEvent $eventsFile 'dialog_appeared' 20
    if (-not $dlg) { throw "no 'dialog_appeared' event" }
    Write-Host "  dialog_appeared id=$($dlg.id) title='$($dlg.title)' buttons=$(($dlg.buttons -join ','))" -ForegroundColor Green

    if ($dlg.title -notmatch 'msgbox fixture') { throw "dialog title mismatch: '$($dlg.title)'" }
    if (-not ($dlg.buttons -contains 'OK')) { throw "OK button missing: $($dlg.buttons -join ',')" }

    # 4. Send respond_dialog OK
    $respondCmd = @{ id='c2'; cmd='respond_dialog'; dialog_id=$dlg.id; button='OK' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $respondCmd -Encoding UTF8

    # 5. Wait for dialog_dismissed
    $dismissed = Wait-ForEvent $eventsFile 'dialog_dismissed' 10 -IdMatch 'c2'
    if (-not $dismissed) { throw "no 'dialog_dismissed' event for c2" }
    Write-Host "  dialog_dismissed id=$($dismissed.dialog_id) button=$($dismissed.button)" -ForegroundColor Green

    # 6. Wait for macro_completed for c1
    $done = Wait-ForEvent $eventsFile 'macro_completed' 10 -IdMatch 'c1'
    if (-not $done) { throw "no 'macro_completed' event for c1" }
    Write-Host "  macro_completed id=c1 duration=$($done.duration_ms)ms" -ForegroundColor Green

    # 7. Close
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no 'closed' event" }
    Write-Host "  closed" -ForegroundColor Green

    $proc.WaitForExit(8000) | Out-Null
    if (-not $proc.HasExited) { throw "host did not exit cleanly" }

    Write-Host "PASS test-msgbox-flow" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-msgbox-flow: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
