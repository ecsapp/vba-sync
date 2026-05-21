# Item 5 test — import / export drive vba-sync's addin (sheet-safe).
#
# The fixture sheetcode.xlsm carries worksheet code-behind (a function in
# a sheet document module). The test exports it (vba-sync writes the
# source tree), imports it back, then compile_checks. The naive sync_vba
# imports Objects/ sheet code-behind as orphan standard modules and
# breaks compilation; the addin's importer replaces the sheet document
# module line-by-line. compile_check ok + the sheet function still under
# a single module proves `import` is sheet-safe.

[CmdletBinding()]
param([string]$SessionId = "test-impexp-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

# Stage the fixture in a temp workdir so vba-sync's <name>/ source folder
# lands somewhere disposable.
$fixtureSrc = Join-Path $here 'fixtures\sheetcode.xlsm'
if (-not (Test-Path $fixtureSrc)) { throw "fixture missing: $fixtureSrc (run build-fixtures.ps1)" }
$workDir = Join-Path $env:TEMP "xc-impexp-$(Get-Random -Maximum 99999)"
New-Item -ItemType Directory -Force -Path $workDir | Out-Null
$workWb = Join-Path $workDir 'sheetcode.xlsm'
Copy-Item -LiteralPath $fixtureSrc -Destination $workWb -Force

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: import/export sheet-safety (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $workWb `
    -SessionId $SessionId -SessionsRoot $sessions

function Read-Events($p) { if (-not (Test-Path $p)) { return @() }; Get-Content -LiteralPath $p | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json } }
function Wait-ForEvent($p, $T, $S=60, $Id=$null) {
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
    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw 'no started' }
    Write-Host '  started' -ForegroundColor Green

    # export — vba-sync writes the source tree (the command pre-arms the
    # addin's success MsgBox itself).
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"export"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'export_completed' 60 -Id 'c1')) {
        $f = Wait-ForEvent $eventsFile 'export_failed' 2 -Id 'c1'
        throw ('no export_completed' + $(if ($f) { " (export_failed: $($f.error))" } else { '' }))
    }
    Write-Host '  export_completed' -ForegroundColor Green

    # the sheet code-behind must have exported under Objects/
    $objDir = Join-Path $workDir 'sheetcode\Objects'
    if (-not (Test-Path $objDir)) { throw "export produced no Objects/ folder at $objDir" }
    $sheetCls = @(Get-ChildItem -LiteralPath $objDir -Filter '*.cls' -ErrorAction SilentlyContinue)
    if ($sheetCls.Count -eq 0) { throw 'export wrote no document .cls under Objects/' }
    Write-Host "  exported sheet code-behind: $($sheetCls.Name -join ', ')" -ForegroundColor Green

    # import — back into the workbook
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"import"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'import_completed' 60 -Id 'c2')) {
        $f = Wait-ForEvent $eventsFile 'import_failed' 2 -Id 'c2'
        throw ('no import_completed' + $(if ($f) { " (import_failed: $($f.error))" } else { '' }))
    }
    Write-Host '  import_completed' -ForegroundColor Green

    # compile_check — a mishandled sheet module breaks compilation
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"compile_check"}' -Encoding UTF8
    $cc = Wait-ForEvent $eventsFile 'compile_result' 30 -Id 'c3'
    if (-not $cc) { throw 'no compile_result' }
    if (-not $cc.ok) { throw "compile FAILED after import — sheet module mishandled (module=$($cc.module) line=$($cc.line))" }
    Write-Host '  compile_check ok — project compiles after import' -ForegroundColor Green

    # the sheet function must survive as a single module entry, not be
    # duplicated into an orphan standard module
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"list_macros"}' -Encoding UTF8
    $lm = Wait-ForEvent $eventsFile 'macros_listed' 30 -Id 'c4'
    if (-not $lm) { throw 'no macros_listed' }
    $sheetFn = @($lm.macros | Where-Object { $_.name -eq 'SheetGreeting' })
    if ($sheetFn.Count -eq 0) { throw 'SheetGreeting lost after import — sheet code-behind dropped' }
    if ($sheetFn.Count -gt 1) { throw "SheetGreeting duplicated across $($sheetFn.Count) modules — orphan-module bug" }
    Write-Host "  SheetGreeting intact under module '$($sheetFn[0].module)' (single, sheet-safe)" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c5","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 20)) { throw 'no closed' }
    $proc.WaitForExit(10000) | Out-Null

    Write-Host 'PASS test-import-export' -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-import-export: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
    if (Test-Path $workDir) { try { Remove-Item -Recurse -Force $workDir } catch {} }
}
