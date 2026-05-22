# Test — VBA project auto-unlock leaves no stray Properties dialog.
#
# password-locked.xlsm has a password-locked VBA project (password
# `test`). Starting a session with -VbaPassword drives
# unlock-vba-project.ps1, which Executes the VBE "VBAProject
# Properties..." menu item to reach the password prompt — that also
# opens the Properties dialog, which the unlock must close afterwards.
#
# This test asserts: the unlock succeeds (project becomes reachable) AND
# no stray Properties dialog is left open (the watcher catches any
# #32770 still up). A stray Properties dialog steals foreground — the
# bug this regression test locks down.

[CmdletBinding()]
param([string]$SessionId = "test-unlock-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$workbook   = Join-Path $here 'fixtures\password-locked.xlsm'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: vba unlock (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $workbook `
    -SessionId $SessionId -SessionsRoot $sessions -VbaPassword 'test'

function Read-Events($p) { if (-not (Test-Path $p)) { return @() }; Get-Content -LiteralPath $p | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json } }
function Wait-ForEvent($p, $T, $S=45, $Id=$null) {
    $end = (Get-Date).AddSeconds($S)
    while ((Get-Date) -lt $end) {
        foreach ($e in Read-Events $p) { if ($e.t -eq $T -and (-not $Id -or $e.id -eq $Id)) { return $e } }
        Start-Sleep -Milliseconds 200
    }
    return $null
}

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    # The unlock runs during startup — give it room (the dialog wait plus
    # the Properties-dialog close poll).
    if (-not (Wait-ForEvent $eventsFile 'started' 45)) { throw 'no started' }
    Write-Host '  started' -ForegroundColor Green

    # Unlock must have succeeded — no vba_unlock_failed.
    $uf = @(Read-Events $eventsFile | Where-Object { $_.t -eq 'vba_unlock_failed' })
    if ($uf.Count -gt 0) { throw "vba_unlock_failed: $($uf[0].error)" }
    Write-Host '  no vba_unlock_failed' -ForegroundColor Green

    # The project is unlocked iff VBComponents are reachable — list_macros
    # returns nothing on a still-locked project.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"list_macros"}' -Encoding UTF8
    $lm = Wait-ForEvent $eventsFile 'macros_listed' 20 -Id 'c1'
    if (-not $lm) { throw 'no macros_listed' }
    if (@($lm.macros | Where-Object { $_.name -eq 'PopMsgBox' }).Count -eq 0) {
        throw 'PopMsgBox not listed — VBComponents not reachable, project still locked'
    }
    Write-Host '  project unlocked — list_macros sees PopMsgBox' -ForegroundColor Green

    # Give the watcher a few polls, then assert NO stray Properties dialog
    # was caught — the watcher emits dialog_appeared for any #32770 left up.
    Start-Sleep -Seconds 3
    $stray = @(Read-Events $eventsFile | Where-Object {
        $_.t -eq 'dialog_appeared' -and
        ("$($_.title)" -match 'Properties' -or "$($_.title)" -match 'VBAProject')
    })
    if ($stray.Count -gt 0) { throw "stray dialog left open after unlock: '$($stray[0].title)'" }
    Write-Host '  no stray Properties dialog after unlock' -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 20)) { throw 'no closed' }
    $proc.WaitForExit(10000) | Out-Null

    Write-Host 'PASS test-vba-unlock' -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-vba-unlock: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
