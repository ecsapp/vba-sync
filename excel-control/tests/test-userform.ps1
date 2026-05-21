# Phase 8 test — UserForm controls via MSAA.
#
# MSForms UserForm controls are WINDOWLESS — a ThunderDFrame UserForm is a
# real HWND but its buttons / option buttons / text boxes are not child
# windows, so EnumChildWindows finds nothing. The watcher drives them
# through MSAA / IAccessible (oleacc.dll):
#   - userform_appeared carries a controls[] list (name, role, value,
#     checked, enabled)
#   - set_form_control selects an option button / toggles a checkbox /
#     sets a text box, and emits form_control_set
#   - respond_dialog resolves a button by caption and clicks it via
#     accDoDefaultAction
#
# Fixture frmLogin (userform.xlsm) carries two option buttons
# ("Units" / "Units / Time"), a checkbox, two text boxes, and OK / Cancel.
# This test: userform_appeared enumerates the controls; set_form_control
# selects the "Units" option (form_control_set, checked:true);
# respond_dialog clicks OK (dialog_dismissed); the macro completes.

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

Write-Host "Test: userform / MSAA (SessionId=$SessionId)" -ForegroundColor Cyan

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

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'
    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no started" }

    # Trigger the UserForm
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"run_macro","name":"ShowLoginForm"}' -Encoding UTF8

    # userform_appeared must carry the MSAA controls[] list
    $form = Wait-ForEvent $eventsFile 'userform_appeared' 20
    if (-not $form) { throw "no userform_appeared" }
    if (-not $form.controls) { throw "userform_appeared carried no controls[] — MSAA enumeration failed" }
    $names = @($form.controls | ForEach-Object { "$($_.name)".Trim() })
    Write-Host "  userform_appeared id=$($form.id) controls: $($names -join ', ')" -ForegroundColor Green
    foreach ($want in @('Units','OK')) {
        if ($names -notcontains $want) { throw "controls[] missing '$want' (got: $($names -join ', '))" }
    }
    Write-Host "  controls[] enumerated the windowless option button + OK button" -ForegroundColor Green

    # set_form_control — select the "Units" option button via MSAA
    $setCmd = @{ id='c2'; cmd='set_form_control'; dialog_id=$form.id; control='Units'; value=$true } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $setCmd -Encoding UTF8
    $set = Wait-ForEvent $eventsFile 'form_control_set' 15 -Id 'c2'
    if (-not $set) {
        $fail = Wait-ForEvent $eventsFile 'form_control_failed' 2 -Id 'c2'
        throw ("no form_control_set" + $(if ($fail) { " (form_control_failed: $($fail.reason))" } else { '' }))
    }
    if (-not $set.checked) { throw "form_control_set but checked=$($set.checked) — the option button was not selected" }
    Write-Host "  set_form_control 'Units' -> form_control_set checked=$($set.checked) role=$($set.role)" -ForegroundColor Green

    # respond_dialog — click OK by caption via MSAA accDoDefaultAction
    $respondCmd = @{ id='c3'; cmd='respond_dialog'; dialog_id=$form.id; button='OK' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $respondCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'dialog_dismissed' 15 -Id 'c3')) { throw "no dialog_dismissed — OK not clicked via MSAA" }
    Write-Host "  respond_dialog 'OK' -> dialog_dismissed" -ForegroundColor Green

    # cmdOK_Click hides the form; the macro returns
    if (-not (Wait-ForEvent $eventsFile 'macro_completed' 15 -Id 'c1')) { throw "no macro_completed — form not dismissed" }
    Write-Host "  macro_completed" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"close"}' -Encoding UTF8
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
