# End-to-end test for the rebuilt VBA Sync.xlam, driven via the
# excel-control session harness (start-session.ps1 + watcher).
#
# Why the harness: dialogs that VBA pops (like the post-export
# "Export completed successfully" MsgBox) get reported as
# `dialog_appeared` events and dismissed via `respond_dialog`. No
# ad-hoc Win32 polling, no stuck COM threads.
#
# Flow:
#   start session on userform.xlsm copy
#   --> dialog_appeared / userform_appeared / etc. events flow back
#   open_workbook VBA Sync.xlam
#   run_macro 'VBA Sync.xlam'!modSync.ExportProject (activate userform.xlsm)
#   on dialog_appeared (title="VBA Sync"): respond_dialog button=OK
#   wait for macro_completed
#   close
#   verify outputs

[CmdletBinding()]
param([switch]$Visible)
$ErrorActionPreference = 'Stop'

$repoRoot   = Resolve-Path (Join-Path $PSScriptRoot '..\..') |
    Select-Object -ExpandProperty Path
$xlamPath   = Join-Path $repoRoot 'VBA Sync.xlam'
$fixtureWb  = Join-Path $repoRoot 'excel-control\canonicalize\test\userform.xlsm'
$sessionPs1 = Join-Path $repoRoot 'excel-control\start-session.ps1'

if (-not (Test-Path -LiteralPath $xlamPath))   { throw "xlam not found: $xlamPath" }
if (-not (Test-Path -LiteralPath $fixtureWb))  { throw "fixture not found: $fixtureWb" }
if (-not (Test-Path -LiteralPath $sessionPs1)) { throw "session host not found: $sessionPs1" }

# Sandbox: copy fixture into a clean dir so the export folder lands beside it.
$sandbox = Join-Path $env:TEMP "vbasync-harness-test-$(Get-Random)"
New-Item -ItemType Directory -Path $sandbox -Force | Out-Null
$wbCopy = Join-Path $sandbox 'userform.xlsm'
Copy-Item -LiteralPath $fixtureWb -Destination $wbCopy -Force
Write-Host "Sandbox: $sandbox" -ForegroundColor Cyan

# Session lives under a unique dir so concurrent test runs don't collide.
$sessionId    = "test-$(Get-Random)"
$sessionsRoot = Join-Path $sandbox 'sessions'
$sessionDir   = Join-Path $sessionsRoot $sessionId
$commandsFile = Join-Path $sessionDir 'commands.jsonl'
$eventsFile   = Join-Path $sessionDir 'events.jsonl'

# Trust access to VBA project model (required for VBProject.* COM access).
function Enable-VbaTrust {
    $verKey = Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        Sort-Object -Property { [version]$_.PSChildName } -Descending |
        Select-Object -First 1
    if (-not $verKey) { return }
    $secPath = Join-Path $verKey.PSPath 'Excel\Security'
    if (-not (Test-Path $secPath)) { New-Item -Path $secPath -Force | Out-Null }
    Set-ItemProperty -Path $secPath -Name 'AccessVBOM' -Value 1 -Type DWord -Force
}
Enable-VbaTrust

# ---------- harness helpers ----------

$script:nextCmdId = 1
function Send-Command([hashtable]$Cmd) {
    if (-not $Cmd.id) { $Cmd.id = "c$($script:nextCmdId)"; $script:nextCmdId++ }
    $line = ConvertTo-Json -Compress -Depth 10 -InputObject $Cmd
    Add-Content -LiteralPath $commandsFile -Value $line -Encoding UTF8
    Write-Host "  > $line" -ForegroundColor DarkGray
    return $Cmd.id
}

# Read all events emitted so far. Cheap enough for a one-shot test.
function Get-AllEvents {
    if (-not (Test-Path -LiteralPath $eventsFile)) { return @() }
    $raw = Get-Content -LiteralPath $eventsFile -Encoding UTF8
    $out = @()
    foreach ($l in $raw) {
        if ([string]::IsNullOrWhiteSpace($l)) { continue }
        try { $out += (ConvertFrom-Json $l) } catch {}
    }
    return $out
}

# Wait for the first event matching the predicate, up to TimeoutSec.
# Returns the event (or $null on timeout).
function Wait-Event {
    param(
        [scriptblock]$Pred,
        [int]$TimeoutSec = 30,
        [string]$Label = '<event>'
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        foreach ($e in Get-AllEvents) {
            if (& $Pred $e) { return $e }
        }
        Start-Sleep -Milliseconds 200
    }
    Write-Host "  TIMEOUT waiting for $Label after ${TimeoutSec}s" -ForegroundColor Red
    return $null
}

# ---------- launch session ----------

Write-Host "[1] Spawning session host (background)..." -ForegroundColor Cyan
$pwshExe = (Get-Command pwsh).Source
$visArg = if ($Visible) { '-Visible' } else { '' }
$sessionArgs = @(
    '-NoProfile', '-File', $sessionPs1,
    '-Workbook', $wbCopy,
    '-SessionId', $sessionId,
    '-SessionsRoot', $sessionsRoot
)
if ($Visible) { $sessionArgs += '-Visible' }
$sessionProc = Start-Process -FilePath $pwshExe -ArgumentList $sessionArgs `
    -PassThru -WindowStyle Hidden `
    -RedirectStandardOutput (Join-Path $sandbox 'session.stdout.log') `
    -RedirectStandardError  (Join-Path $sandbox 'session.stderr.log')
Write-Host "  session PID: $($sessionProc.Id)"

$ok = $false
$excelPid = $null
try {
    # Wait for the 'started' event so we know Excel is up.
    $started = Wait-Event -Pred { $args[0].t -eq 'started' } -TimeoutSec 30 -Label 'started'
    if (-not $started) { throw "session never emitted 'started'" }
    Write-Host "[2] Session ready (Excel pid will be in state.json)" -ForegroundColor Cyan

    # Grab excel_pid from state.json for PID-targeted cleanup.
    $state = Get-Content -LiteralPath (Join-Path $sessionDir 'state.json') -Raw | ConvertFrom-Json
    $excelPid = [int]$state.excel_pid
    Write-Host "  excel_pid: $excelPid"

    # ---------- open xlam alongside fixture ----------
    Write-Host "[3] open_workbook VBA Sync.xlam..." -ForegroundColor Cyan
    $openId = Send-Command @{ cmd = 'open_workbook'; path = $xlamPath }
    $opened = Wait-Event -Pred { $args[0].t -eq 'workbook_opened' -and $args[0].id -eq $openId } `
        -TimeoutSec 30 -Label "workbook_opened/$openId"
    if (-not $opened) { throw "xlam did not open" }

    # ---------- run ExportProject ----------
    Write-Host "[4] run_macro ExportProject..." -ForegroundColor Cyan
    $runId = Send-Command @{
        cmd = 'run_macro'
        name = "'VBA Sync.xlam'!modSync.ExportProject"
        args = @($null)
        activate_workbook = 'userform.xlsm'
    }
    # The macro will pause when it pops the success MsgBox. We need to
    # dismiss it before macro_completed will fire. Watch for both in
    # parallel.
    $msgboxDismissed = $false
    $macroEvent = $null
    $deadline = (Get-Date).AddSeconds(90)
    while ((Get-Date) -lt $deadline) {
        $events = Get-AllEvents
        if (-not $msgboxDismissed) {
            # Look for dialog_appeared whose title says "VBA Sync".
            $dlg = $events | Where-Object {
                $_.t -eq 'dialog_appeared' -and $_.title -like '*VBA Sync*'
            } | Select-Object -First 1
            if ($dlg) {
                Write-Host "  dialog: title='$($dlg.title)' buttons=[$($dlg.buttons -join ', ')]" -ForegroundColor Yellow
                Send-Command @{ cmd = 'respond_dialog'; dialog_id = $dlg.id; button = 'OK' } | Out-Null
                $msgboxDismissed = $true
            }
        }
        $done = $events | Where-Object { ($_.t -eq 'macro_completed' -or $_.t -eq 'macro_failed') -and $_.id -eq $runId } | Select-Object -First 1
        if ($done) { $macroEvent = $done; break }
        Start-Sleep -Milliseconds 200
    }
    if (-not $macroEvent) { throw "macro_completed/macro_failed never arrived" }
    if ($macroEvent.t -eq 'macro_failed') {
        Write-Host "  MACRO FAILED: $($macroEvent.error)" -ForegroundColor Red
        throw "ExportProject failed: $($macroEvent.error)"
    }
    Write-Host "  macro completed in $($macroEvent.duration_ms)ms" -ForegroundColor Green

    # ---------- close session ----------
    Write-Host "[5] close session..." -ForegroundColor Cyan
    Send-Command @{ cmd = 'close' } | Out-Null
    [void](Wait-Event -Pred { $args[0].t -eq 'closing' } -TimeoutSec 15 -Label 'closing')
    [void]$sessionProc.WaitForExit(15000)

    $ok = $true
} finally {
    # PID-targeted cleanup. Never blanket-kill EXCEL.EXE.
    if ($sessionProc -and -not $sessionProc.HasExited) {
        try { Stop-Process -Id $sessionProc.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
    if ($excelPid) {
        Start-Sleep -Milliseconds 400
        try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch {}
    }
}

if (-not $ok) {
    Write-Host ""
    Write-Host "Session stdout:" -ForegroundColor Yellow
    if (Test-Path (Join-Path $sandbox 'session.stdout.log')) {
        Get-Content -LiteralPath (Join-Path $sandbox 'session.stdout.log') | ForEach-Object { Write-Host "  $_" }
    }
    Write-Host "Session stderr:" -ForegroundColor Yellow
    if (Test-Path (Join-Path $sandbox 'session.stderr.log')) {
        Get-Content -LiteralPath (Join-Path $sandbox 'session.stderr.log') | ForEach-Object { Write-Host "  $_" }
    }
    Write-Host "Events tail:" -ForegroundColor Yellow
    if (Test-Path $eventsFile) {
        Get-Content -LiteralPath $eventsFile -Tail 20 | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}

# ---------- assertions ----------
Write-Host "[6] Verifying outputs..." -ForegroundColor Cyan
$exportDir = Join-Path $sandbox 'userform'
$pass = $true
function Assert-Path([string]$P, [string]$Label) {
    if (Test-Path -LiteralPath $P) {
        Write-Host "    OK   $Label" -ForegroundColor Green
    } else {
        Write-Host "    FAIL $Label  (missing: $P)" -ForegroundColor Red
        $script:pass = $false
    }
}
function Assert-NoPath([string]$P, [string]$Label) {
    if (-not (Test-Path -LiteralPath $P)) {
        Write-Host "    OK   $Label" -ForegroundColor Green
    } else {
        Write-Host "    FAIL $Label  (present: $P)" -ForegroundColor Red
        $script:pass = $false
    }
}

Assert-Path (Join-Path $sandbox 'tools\start-session.ps1') 'tools/start-session.ps1'
Assert-Path (Join-Path $sandbox 'tools\session-dialog-watcher.ps1') 'tools/session-dialog-watcher.ps1'
Assert-Path (Join-Path $sandbox 'tools\capture.ps1') 'tools/capture.ps1'
Assert-Path (Join-Path $sandbox 'tools\INTERFACE.md') 'tools/INTERFACE.md'
Assert-Path (Join-Path $sandbox 'tools\clsAssert.cls') 'tools/clsAssert.cls'
Assert-Path (Join-Path $sandbox 'tools\.claude\settings.json') 'tools/.claude/settings.json'
Assert-Path (Join-Path $sandbox '.claude\skills\excel-control\SKILL.md') '.claude/skills/excel-control/SKILL.md'

$frx = Join-Path $exportDir 'Forms\frmLogin.frx'
Assert-Path $frx 'Forms/frmLogin.frx'
Assert-NoPath (Join-Path $exportDir 'Forms\frmLogin.frx.fingerprint') 'no fingerprint sidecar'

Write-Host ""
if ($pass) {
    Write-Host "PASS -- harness-driven integration test green." -ForegroundColor Green
    Write-Host "Sandbox: $sandbox" -ForegroundColor DarkGray
    exit 0
} else {
    Write-Host "FAIL -- see above. Sandbox preserved: $sandbox" -ForegroundColor Red
    exit 1
}
