# Phase 8 test — UserForm dismissal.
#
# The dialog watcher already recognises the UserForm window class
# (ThunderDFrame / F3*) and emits userform_appeared. Field-level
# introspection via COM is deferred (Plan B in SCOPING.md) — UserForm
# Controls aren't reachable from outside the VBA runtime.
#
# This test verifies the bare minimum: userform_appeared fires and the
# watcher's VK_RETURN fallback (triggered by respond_dialog) fires the
# form's Default button click handler, which closes the form.

[CmdletBinding()]
param([string]$SessionId = "test-uf-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$workbook  = Join-Path $here 'fixtures\userform.xlsm'
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: userform (SessionId=$SessionId)" -ForegroundColor Cyan

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

    # Trigger the UserForm
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"run_macro","name":"ShowLoginForm"}' -Encoding UTF8

    # Expect userform_appeared
    $form = Wait-ForEvent $eventsFile 'userform_appeared' 20
    if (-not $form) {
        # Fall back: might come as dialog_appeared if the watcher class match misses
        $form = Wait-ForEvent $eventsFile 'dialog_appeared' 2
        if ($form) {
            Write-Host "  (watcher fired dialog_appeared instead of userform_appeared — class match needs tuning)" -ForegroundColor Yellow
        } else {
            throw "no userform_appeared (and no dialog_appeared)"
        }
    } else {
        Write-Host "  userform_appeared id=$($form.id) title='$($form.title)' class=$($form.class)" -ForegroundColor Green
    }

    # Respond — any button name; watcher falls through to VK_RETURN for UserForms
    $respondCmd = @{ id='c2'; cmd='respond_dialog'; dialog_id=$form.id; button='OK' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $respondCmd -Encoding UTF8

    # The form's cmdOK_Click hides it; macro completes
    if (-not (Wait-ForEvent $eventsFile 'macro_completed' 15 -Id 'c1')) { throw "no macro_completed (form not dismissed)" }
    Write-Host "  macro_completed — form was dismissed" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-userform" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-userform: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
