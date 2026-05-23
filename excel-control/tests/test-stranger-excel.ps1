# Stranger-Excel regression — verifies the watcher's stranger-sweep
# surfaces an EXCEL.EXE that appeared after session startup, and that
# close_recovery cleanly terminates it via the validated PID path.
#
# Spawns a vanilla second Excel via New-Object Excel.Application (not a
# real recovery panel — title has no " - Repaired ", so the event reports
# looks_like_recovery: false). The stranger detection is the gate; the
# recovery classifier is only a payload qualifier.
#
# Cleanup-careful: only ever Stop-Process the specific stranger PID we
# spawned. Never blanket-kill EXCEL.EXE.

[CmdletBinding()]
param(
    [string]$SessionId = "test-stranger-$(Get-Random -Maximum 99999)"
)

$ErrorActionPreference = 'Stop'

$here       = $PSScriptRoot
$repo       = Resolve-Path (Join-Path $here '..\..')
$startSess  = Join-Path $repo 'excel-control\start-session.ps1'
$workbook   = Join-Path $here 'fixtures\empty.xlsm'
$sessions   = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

Write-Host "Test: stranger-excel (SessionId=$SessionId)" -ForegroundColor Cyan

$proc       = $null
$strangerXl = $null
$strangerPid = 0

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

function Wait-ForEventMatching($EventsPath, [string]$Type, [scriptblock]$Predicate, [int]$TimeoutSec = 15) {
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        foreach ($e in Read-Events $EventsPath) {
            if ($e.t -eq $Type -and (& $Predicate $e)) { return $e }
        }
        Start-Sleep -Milliseconds 200
    }
    return $null
}

try {
    $proc = Start-SessionHost -StartSession $startSess -Workbook $workbook `
        -SessionId $SessionId -SessionsRoot $sessions

    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    # 1. Wait for started
    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no 'started' event" }
    Write-Host "  started" -ForegroundColor Green

    # 2. Spawn a stranger Excel — vanilla, no recovery panel
    $strangerXl = New-Object -ComObject Excel.Application
    $strangerXl.Visible = $false
    $strangerXl.DisplayAlerts = $false
    # Find this Excel's PID via its Hwnd (or as the new EXCEL.EXE since spawn)
    $hwnd = 0
    try { $hwnd = [int64]$strangerXl.Hwnd } catch {}
    if ($hwnd -ne 0) {
        # GetWindowThreadProcessId — declared in start-session's Win32 type.
        # Re-declare a minimal version here.
        if (-not ([System.Management.Automation.PSTypeName]'TestStrangerWin32').Type) {
            Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class TestStrangerWin32 {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
        }
        $pidOut = [uint32]0
        [void][TestStrangerWin32]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$pidOut)
        $strangerPid = [int]$pidOut
    }
    if ($strangerPid -le 0) { throw "could not determine stranger Excel PID" }
    Write-Host "  spawned stranger excel pid=$strangerPid" -ForegroundColor Green

    # 3. recovery_instance_detected should fire for this PID
    $evt = Wait-ForEventMatching $eventsFile 'recovery_instance_detected' {
        param($e); [int]$e.pid -eq $strangerPid
    } 12
    if (-not $evt) { throw "no 'recovery_instance_detected' for pid=$strangerPid" }
    if ([bool]$evt.looks_like_recovery) {
        # A vanilla new Excel has no "Repaired" in its title — should be false.
        Write-Host "  WARN: looks_like_recovery=true on a vanilla stranger; title='$($evt.main_window_title)'" -ForegroundColor Yellow
    }
    Write-Host "  recovery_instance_detected pid=$($evt.pid) title='$($evt.main_window_title)' looks_like_recovery=$($evt.looks_like_recovery)" -ForegroundColor Green

    # 4. close_recovery
    $closeCmd = @{ id='c1'; cmd='close_recovery'; pid=$strangerPid } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $closeCmd -Encoding UTF8
    $rc = Wait-ForEvent $eventsFile 'recovery_closed' 8 -IdMatch 'c1'
    if (-not $rc) { throw "no 'recovery_closed' event for c1" }
    if (-not $rc.ok) { throw "recovery_closed.ok=false" }
    if ([int]$rc.pid -ne $strangerPid) { throw "recovery_closed.pid mismatch ($($rc.pid) vs $strangerPid)" }
    Write-Host "  recovery_closed pid=$($rc.pid) ok=true" -ForegroundColor Green

    # 5. Stranger PID must be gone
    Start-Sleep -Milliseconds 500
    $stillThere = Get-Process -Id $strangerPid -ErrorAction SilentlyContinue
    if ($stillThere) { throw "stranger pid=$strangerPid still alive after recovery_closed" }
    $strangerPid = 0  # don't double-kill in finally
    Write-Host "  stranger pid gone" -ForegroundColor Green

    # 6. close
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no 'closed' event" }
    Write-Host "  closed" -ForegroundColor Green

    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-stranger-excel" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-stranger-excel: $($_.Exception.Message)" -ForegroundColor Red
    if ($eventsFile -and (Test-Path $eventsFile)) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    # Release the stranger COM ref first so its Excel can exit on its own
    if ($strangerXl) {
        try { $strangerXl.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($strangerXl) | Out-Null } catch {}
        $strangerXl = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    # Belt-and-braces: kill ONLY our recorded stranger PID if it survived.
    if ($strangerPid -gt 0) {
        try { Stop-Process -Id $strangerPid -Force -ErrorAction SilentlyContinue } catch {}
    }
    if ($proc) { Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir }
}
