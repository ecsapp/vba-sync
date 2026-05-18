# Phases 7.5, 8.5, 8.7 combined test — save/calculate/refresh + multi-workbook + create.

[CmdletBinding()]
param([string]$SessionId = "test-state-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$workbook  = Join-Path $here 'fixtures\empty.xlsm'
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

Write-Host "Test: state control (SessionId=$SessionId)" -ForegroundColor Cyan

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

    # Calculate (no-op against empty sheet but should succeed)
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"calculate"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'calculated' 10 -Id 'c1')) { throw "no calculated" }
    Write-Host "  calculate: ok" -ForegroundColor Green

    # Refresh all
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"refresh_all"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'refreshed_all' 15 -Id 'c2')) { throw "no refreshed_all" }
    Write-Host "  refresh_all: ok" -ForegroundColor Green

    # save_as into a temp path
    $tmpSave = Join-Path $sessionDir 'saved.xlsm'
    $saveCmd = @{ id='c3'; cmd='save_as'; path=$tmpSave; file_format=52 } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $saveCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'workbook_saved_as' 10 -Id 'c3')) { throw "no workbook_saved_as" }
    if (-not (Test-Path $tmpSave)) { throw "save_as did not create file" }
    Write-Host "  save_as: $tmpSave" -ForegroundColor Green

    # create_workbook with save_as
    $newWbPath = Join-Path $sessionDir 'created.xlsm'
    $createCmd = @{ id='c4'; cmd='create_workbook'; save_as=$newWbPath } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $createCmd -Encoding UTF8
    $created = Wait-ForEvent $eventsFile 'workbook_created' 15 -Id 'c4'
    if (-not $created) { throw "no workbook_created" }
    if (-not (Test-Path $newWbPath)) { throw "create_workbook didn't save_as: $newWbPath" }
    Write-Host "  create_workbook: $($created.name) -> $newWbPath" -ForegroundColor Green

    # close_workbook for the newly created one (by name)
    $closeCmd = @{ id='c5'; cmd='close_workbook'; name=$created.name; save=$false } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $closeCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'workbook_closed_cmd' 10 -Id 'c5')) { throw "no workbook_closed_cmd" }
    Write-Host "  close_workbook: $($created.name)" -ForegroundColor Green

    # open_workbook (re-open the saved one)
    $openCmd = @{ id='c6'; cmd='open_workbook'; path=$tmpSave } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $openCmd -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'workbook_opened' 10 -Id 'c6')) { throw "no workbook_opened" }
    Write-Host "  open_workbook: $tmpSave" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c7","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-state-control" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-state-control: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    if ($proc -and -not $proc.HasExited) { try { $proc.Kill() } catch {} }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
