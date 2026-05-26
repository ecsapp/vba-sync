# Phase 6 test — compile_check + import.
#
# Round trip: import a broken module → compile_check fails →
# import a fixed module → compile_check passes.

[CmdletBinding()]
param(
    [string]$SessionId = "test-cs-$(Get-Random -Maximum 99999)"
)

$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$fixture    = Join-Path $here 'fixtures\msgbox.xlsm'  # has modTest pre-loaded
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

$broken = @'
Attribute VB_Name = "modSales"
Option Explicit

Public Sub Broken()
    Dim x As Long
    y = 5    ' undeclared
End Sub
'@
$stagedWb = New-StagedWorkbook -Fixture $fixture -Modules @{ modSales = $broken }
$stemDir  = Join-Path ([System.IO.Path]::GetDirectoryName($stagedWb)) `
    ([System.IO.Path]::GetFileNameWithoutExtension($stagedWb))

Write-Host "Test: compile + import (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $stagedWb `
    -SessionId $SessionId -SessionsRoot $sessions

function Read-Events([string]$path) {
    if (-not (Test-Path $path)) { return @() }
    Get-Content -LiteralPath $path | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json }
}
function Wait-ForEvent($EventsPath, [string]$Type, [int]$TimeoutSec = 60, [string]$IdMatch = $null) {
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

    # Import 1: broken module
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"import"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'import_completed' 60 -IdMatch 'c1')) {
        $f = Wait-ForEvent $eventsFile 'import_failed' 2 -IdMatch 'c1'
        throw ('no import_completed for c1' + $(if ($f) { " (import_failed: $($f.error))" } else { '' }))
    }
    Write-Host "  import c1 ok (broken module)" -ForegroundColor Green

    # Compile 1: expect fail
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"compile_check"}' -Encoding UTF8
    $r1 = Wait-ForEvent $eventsFile 'compile_result' 20 -IdMatch 'c2'
    if (-not $r1) { throw "no compile_result for c2" }
    if ($r1.ok) { throw "expected compile FAIL, got ok=true" }
    Write-Host "  compile c2: ok=false module=$($r1.module) line=$($r1.line)" -ForegroundColor Green

    # Swap to fixed module on disk, re-import.
    $fixed = @'
Attribute VB_Name = "modSales"
Option Explicit

Public Sub Fixed()
    Dim x As Long
    x = 5
End Sub
'@
    Set-Content -LiteralPath (Join-Path $stemDir 'Modules\modSales.bas') -Value $fixed -Encoding ASCII

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"import"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'import_completed' 60 -IdMatch 'c3')) {
        $f = Wait-ForEvent $eventsFile 'import_failed' 2 -IdMatch 'c3'
        throw ('no import_completed for c3' + $(if ($f) { " (import_failed: $($f.error))" } else { '' }))
    }
    Write-Host "  import c3 ok (fixed module)" -ForegroundColor Green

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
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
