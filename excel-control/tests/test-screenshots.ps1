# Phase 4 test — screenshots auto-captured on dialog_appeared, plus
# on-demand `screenshot` command for window/sheet/dialog targets.

[CmdletBinding()]
param(
    [string]$SessionId = "test-shot-$(Get-Random -Maximum 99999)"
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

Write-Host "Test: screenshots (SessionId=$SessionId)" -ForegroundColor Cyan

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

function Assert-Png([string]$Path, [string]$Label) {
    if (-not (Test-Path $Path)) { throw "$Label PNG missing: $Path" }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 100) { throw "$Label PNG too small: $($bytes.Length) bytes" }
    # PNG magic bytes: 89 50 4E 47 0D 0A 1A 0A
    if ($bytes[0] -ne 0x89 -or $bytes[1] -ne 0x50 -or $bytes[2] -ne 0x4E -or $bytes[3] -ne 0x47) {
        throw "$Label not a valid PNG (bad magic bytes)"
    }
    return $bytes.Length
}

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no 'started'" }

    # Trigger dialog and verify screenshot is captured with the event
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"run_macro","name":"PopMsgBox"}' -Encoding UTF8
    $dlg = Wait-ForEvent $eventsFile 'dialog_appeared' 20
    if (-not $dlg) { throw "no dialog_appeared" }
    if (-not $dlg.screenshot) { throw "dialog_appeared missing screenshot field" }
    $sz = Assert-Png $dlg.screenshot 'dialog'
    Write-Host "  dialog screenshot: $($dlg.screenshot) ($sz bytes)" -ForegroundColor Green

    # Dismiss
    Add-Content -LiteralPath $commandsFile -Value (@{id='c2'; cmd='respond_dialog'; dialog_id=$dlg.id; button='OK'} | ConvertTo-Json -Compress) -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'macro_completed' 10 -IdMatch 'c1')) { throw "macro didn't complete" }

    # On-demand window screenshot
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"screenshot","target":"window"}' -Encoding UTF8
    $shot = Wait-ForEvent $eventsFile 'screenshot_captured' 10 -IdMatch 'c3'
    if (-not $shot) { throw "no screenshot_captured for c3" }
    $sz = Assert-Png $shot.path 'window'
    Write-Host "  window screenshot: $($shot.path) ($($shot.width)x$($shot.height), $sz bytes)" -ForegroundColor Green

    # Worksheet target
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"screenshot","target":"worksheet:Sheet1"}' -Encoding UTF8
    $shot2 = Wait-ForEvent $eventsFile 'screenshot_captured' 10 -IdMatch 'c4'
    if (-not $shot2) { throw "no screenshot_captured for c4 (worksheet)" }
    $sz = Assert-Png $shot2.path 'worksheet'
    Write-Host "  worksheet screenshot: $($shot2.path) ($($shot2.width)x$($shot2.height), $sz bytes)" -ForegroundColor Green

    # Close
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c5","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-screenshots" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-screenshots: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
