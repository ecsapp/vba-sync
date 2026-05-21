# Phase 9 test — vba-sync Export bundles excel-control as a skill.
#
# Eat-your-own-dogfood: drive vba-sync's ExportProject THROUGH the
# excel-control session protocol. The session's dialog watcher catches
# vba-sync's "Export completed" MsgBox; this test (as the agent) responds
# with OK. Demonstrates the bidirectional model end-to-end.

[CmdletBinding()]
param([string]$SessionId = "test-bundle-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

$AddinPath = 'C:\Users\ArnaudLavignolle\AppData\Roaming\Microsoft\AddIns\VBA Sync.xlam'
if (-not (Test-Path $AddinPath)) { throw "VBA Sync.xlam not found at $AddinPath" }

# Stage a fresh copy of the fixture in a temp workdir so we can verify
# the export bundles the harness skill next to it
$fixtureSrc = Join-Path $here 'fixtures\msgbox.xlsm'
$workDir = Join-Path $env:TEMP "vba-sync-bundle-test-$(Get-Random -Maximum 99999)"
New-Item -ItemType Directory -Force -Path $workDir | Out-Null
$workWb = Join-Path $workDir 'msgbox.xlsm'
Copy-Item -LiteralPath $fixtureSrc -Destination $workWb -Force

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: bundle (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $workWb `
    -SessionId $SessionId -SessionsRoot $sessions

function Read-Events($p) { if (-not (Test-Path $p)) { return @() }; Get-Content -LiteralPath $p | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json } }
function Wait-ForEvent($p, $T, $S=30, $Id=$null) {
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
    Write-Host "  session started" -ForegroundColor Green

    # Load the xlam so its macros are runnable (open_workbook activates
    # the addin; we'll re-activate fixture below).
    $openCmd = @{ id='c1'; cmd='open_workbook'; path=$AddinPath } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $openCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'workbook_opened' 15 -Id 'c1')) { throw "no workbook_opened for xlam" }
    Write-Host "  xlam loaded" -ForegroundColor Green

    # Sync a tiny wrapper macro into the fixture: it calls modSync.ExportProject
    # passing Nothing for the `control As Object` ribbon arg. Lets us trigger
    # vba-sync's export from outside the ribbon via Application.Run on a
    # zero-arg macro (PowerShell's COM dispatch can't synthesize a VBA Object
    # arg for the direct ExportProject call).
    $srcDir = Join-Path $sessionDir '_src'
    New-Item -ItemType Directory -Force -Path (Join-Path $srcDir 'Modules') | Out-Null
    $bas = @'
Attribute VB_Name = "modTrigger"
Option Explicit

Public Sub TriggerVbaSyncExport()
    Application.Run "'VBA Sync.xlam'!modSync.ExportProject", Nothing
End Sub
'@
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modTrigger.bas') -Value $bas -Encoding ASCII
    $syncCmd = @{ id='c2'; cmd='sync_vba'; source_dir=$srcDir } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $syncCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'sync_completed' 15 -Id 'c2')) { throw "no sync_completed" }

    # Run the wrapper. activate_workbook to make the fixture ActiveWorkbook
    # for vba-sync's TargetWB() lookup.
    $runCmd = @{
        id   = 'c3'
        cmd  = 'run_macro'
        name = 'TriggerVbaSyncExport'
        activate_workbook = 'msgbox.xlsm'
    } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $runCmd -Encoding UTF8

    # vba-sync ends with MsgBox "VBA Sync export completed successfully!"
    # The session's dialog watcher catches it. We (the agent) respond OK.
    $dlg = Wait-ForEvent $eventsFile 'dialog_appeared' 60
    if (-not $dlg) { throw "no dialog_appeared from ExportProject" }
    if ($dlg.title -notmatch 'VBA Sync') {
        throw "unexpected dialog title: '$($dlg.title)'"
    }
    Write-Host "  dialog_appeared id=$($dlg.id) title='$($dlg.title)'" -ForegroundColor Green

    $respondCmd = @{ id='c4'; cmd='respond_dialog'; dialog_id=$dlg.id; button='OK' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $respondCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'dialog_dismissed' 15 -Id 'c4')) { throw "no dialog_dismissed" }
    Write-Host "  dialog_dismissed via respond_dialog" -ForegroundColor Green

    if (-not (Wait-ForEvent $eventsFile 'macro_completed' 30 -Id 'c3')) { throw "no macro_completed for ExportProject" }
    Write-Host "  ExportProject completed" -ForegroundColor Green

    # Close session
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c5","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    # Verify the harness bundled as a self-contained skill under
    # .claude/skills/excel-control/ next to the exported xlsm.
    $exportFolder = Join-Path $workDir 'msgbox'
    $skillFolder  = Join-Path $workDir '.claude\skills\excel-control'
    if (-not (Test-Path $exportFolder)) { throw "per-workbook export folder not created at $exportFolder" }
    if (-not (Test-Path $skillFolder))  { throw "skill not bundled at $skillFolder" }

    $expected = @(
        'SKILL.md',
        'tools\start-session.ps1',
        'tools\session-dialog-watcher.ps1',
        'tools\capture.ps1',
        'tools\unlock-vba-project.ps1',
        'tools\INTERFACE.md',
        'tools\clsAssert.cls'
    )
    foreach ($rel in $expected) {
        $p = Join-Path $skillFolder $rel
        if (-not (Test-Path -LiteralPath $p)) { throw "missing in bundle: .claude/skills/excel-control/$rel" }
        Write-Host "  bundled: .claude/skills/excel-control/$rel  ($((Get-Item -LiteralPath $p).Length) bytes)" -ForegroundColor Green
    }

    # The harness must no longer land under a repo-root tools/.
    if (Test-Path (Join-Path $workDir 'tools')) { throw "stale repo-root tools/ should not be bundled" }
    Write-Host "  no stale repo-root tools/" -ForegroundColor Green
    Remove-Item -Recurse -Force $workDir

    Write-Host "PASS test-bundle" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-bundle: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
    if (Test-Path $workDir) { try { Remove-Item -Recurse -Force $workDir } catch {} }
}
