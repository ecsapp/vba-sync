# Phase 6 test — compile_check + sync_vba.
#
# Round trip: sync_vba a broken module → compile_check fails →
# sync_vba a fixed module → compile_check passes.

[CmdletBinding()]
param(
    [string]$SessionId = "test-cs-$(Get-Random -Maximum 99999)"
)

$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$workbook   = Join-Path $here 'fixtures\msgbox.xlsm'  # has modTest pre-loaded
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

# Build a scratch source dir with one good module + one broken
$srcDir = Join-Path $sessionDir '_src'
New-Item -ItemType Directory -Force -Path (Join-Path $srcDir 'Modules') | Out-Null

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

Write-Host "Test: compile + sync (SessionId=$SessionId)" -ForegroundColor Cyan

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

    # Sync 1: broken module
    $broken = @'
Attribute VB_Name = "modSales"
Option Explicit

Public Sub Broken()
    Dim x As Long
    y = 5    ' undeclared
End Sub
'@
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modSales.bas') -Value $broken -Encoding ASCII

    $syncCmd1 = @{ id='c1'; cmd='sync_vba'; source_dir=$srcDir } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $syncCmd1 -Encoding UTF8
    $sync1 = Wait-ForEvent $eventsFile 'sync_completed' 20 -IdMatch 'c1'
    if (-not $sync1) { throw "no sync_completed for c1" }
    Write-Host "  sync c1: imported=$($sync1.imported -join ',') removed=$($sync1.removed -join ',')" -ForegroundColor Green

    # Compile 1: expect fail
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"compile_check"}' -Encoding UTF8
    $r1 = Wait-ForEvent $eventsFile 'compile_result' 20 -IdMatch 'c2'
    if (-not $r1) { throw "no compile_result for c2" }
    if ($r1.ok) { throw "expected compile FAIL, got ok=true" }
    Write-Host "  compile c2: ok=false module=$($r1.module) line=$($r1.line)" -ForegroundColor Green

    # Sync 2: fixed module
    $fixed = @'
Attribute VB_Name = "modSales"
Option Explicit

Public Sub Fixed()
    Dim x As Long
    x = 5
End Sub
'@
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modSales.bas') -Value $fixed -Encoding ASCII

    $syncCmd2 = @{ id='c3'; cmd='sync_vba'; source_dir=$srcDir } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $syncCmd2 -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'sync_completed' 20 -IdMatch 'c3')) { throw "no sync c3" }
    Write-Host "  sync c3 ok" -ForegroundColor Green

    # Compile 2: expect pass
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"compile_check"}' -Encoding UTF8
    $r2 = Wait-ForEvent $eventsFile 'compile_result' 20 -IdMatch 'c4'
    if (-not $r2) { throw "no compile_result for c4" }
    if (-not $r2.ok) { throw "expected compile OK, got ok=false (module=$($r2.module) line=$($r2.line))" }
    Write-Host "  compile c4: ok=true" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c5","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-compile-sync" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-compile-sync: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    if ($proc -and -not $proc.HasExited) { try { $proc.Kill() } catch {} }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
