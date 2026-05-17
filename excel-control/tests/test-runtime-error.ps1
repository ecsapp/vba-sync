# Phase 5 test — runtime_error capture against runtime-error.xlsm.

[CmdletBinding()]
param(
    [string]$SessionId = "test-rterr-$(Get-Random -Maximum 99999)"
)

$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$workbook   = Join-Path $here 'fixtures\runtime-error.xlsm'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

Write-Host "Test: runtime_error (SessionId=$SessionId)" -ForegroundColor Cyan

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

    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no started" }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"run_macro","name":"TriggerError"}' -Encoding UTF8

    # runtime_error uses the dialog id, not the command id. Agent
    # correlates by ordering (this event arrives between command_ack
    # and macro_failed for the same command).
    $err = Wait-ForEvent $eventsFile 'runtime_error' 20
    if (-not $err) { throw "no runtime_error event" }
    Write-Host "  runtime_error number=$($err.number) description='$($err.description)' duration=$($err.duration_ms)ms" -ForegroundColor Green
    if ($err.number -ne 9) { throw "expected Err 9 (Subscript out of range), got $($err.number)" }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-runtime-error" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-runtime-error: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    if ($proc -and -not $proc.HasExited) { try { $proc.Kill() } catch {} }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
