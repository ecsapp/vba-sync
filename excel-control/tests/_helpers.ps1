# Shared helpers for excel-control test scripts.
#
# CORE PRINCIPLE: never blanket-kill EXCEL.EXE. The user may have their
# own Excel session open with unsaved work; killing it is destructive
# and silent. Instead, target the specific PIDs the test spawned.
#
# Each session records its `pid` (the PowerShell session host) and
# `excel_pid` (the COM-spawned EXCEL.EXE) in `sessions/<id>/state.json`.
# Tests read those PIDs and kill only them.

function Stop-SessionProcesses {
    <#
        .SYNOPSIS
        Cleanup helper for test teardown — kills ONLY the session host
        PowerShell process and the EXCEL.EXE it spawned, NEVER any other
        Excel processes (the user may have their own session open).
    #>
    param(
        [System.Diagnostics.Process]$HostProc,
        [string]$SessionDir
    )

    $pidsToKill = New-Object System.Collections.ArrayList

    # Host process (the start-session.ps1 pwsh)
    if ($HostProc -and -not $HostProc.HasExited) {
        [void]$pidsToKill.Add($HostProc.Id)
    }

    # EXCEL.EXE PID from state.json
    if ($SessionDir) {
        $stateFile = Join-Path $SessionDir 'state.json'
        if (Test-Path $stateFile) {
            try {
                $state = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json
                if ($state.excel_pid -and [int]$state.excel_pid -gt 0) {
                    [void]$pidsToKill.Add([int]$state.excel_pid)
                }
            } catch {}
        }
    }

    foreach ($targetPid in $pidsToKill) {
        try {
            Get-Process -Id $targetPid -ErrorAction Stop |
                Stop-Process -Force -ErrorAction SilentlyContinue
        } catch {
            # Process already gone — that's the happy path
        }
    }
}

function Clear-OrphanSessionExcels {
    <#
        .SYNOPSIS
        Pre-test cleanup — finds state.json files from prior sessions in
        the given sessions root and kills any EXCEL.EXE PIDs that are
        still alive (orphan from crashed/aborted tests). Never touches
        Excel PIDs that weren't recorded by a session.
    #>
    param([string]$SessionsRoot)
    if (-not (Test-Path $SessionsRoot)) { return }
    Get-ChildItem -LiteralPath $SessionsRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $stateFile = Join-Path $_.FullName 'state.json'
        if (-not (Test-Path $stateFile)) { return }
        try {
            $state = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json
            # A cleanly-closed session has no orphan, and its recorded PIDs
            # may since have been recycled by unrelated processes. Prune the
            # dir so the pile cannot grow into a stale-PID kill hazard.
            if ($state.status -eq 'closed') {
                Remove-Item -Recurse -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue
                return
            }
            # Kill a recorded PID only if the live process is plausibly this
            # session's: a recycled PID starts well after the session did.
            $startedAt = $null
            try { $startedAt = [datetime]$state.started_at } catch {}
            if (-not $startedAt) { return }
            $cutoff = $startedAt.AddMinutes(2)
            foreach ($pair in @(
                @{ Id = $state.excel_pid; Name = 'EXCEL' },
                @{ Id = $state.pid;       Name = 'pwsh'  }
            )) {
                if (-not $pair.Id -or [int]$pair.Id -le 0) { continue }
                $proc = Get-Process -Id ([int]$pair.Id) -ErrorAction SilentlyContinue
                if (-not $proc -or $proc.ProcessName -ne $pair.Name) { continue }
                try { if ($proc.StartTime -gt $cutoff) { continue } } catch { continue }
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}

function Start-SessionHost {
    <#
        .SYNOPSIS
        Spawn a start-session.ps1 host process for a test.

        Builds one pre-quoted argument string rather than passing
        -ArgumentList an array: an array does NOT quote elements
        containing spaces, so a workbook path with a space (e.g.
        "VBA Sync.xlam") gets split and breaks the host's parameter
        binding. The host then dies before creating the session dir and
        the test only sees a missing 'started' event.
    #>
    param(
        [Parameter(Mandatory)][string]$StartSession,
        [Parameter(Mandatory)][string]$Workbook,
        [Parameter(Mandatory)][string]$SessionId,
        [Parameter(Mandatory)][string]$SessionsRoot,
        [string]$VbaPassword
    )
    $argStr = '-NoProfile -File "{0}" -Workbook "{1}" -SessionId "{2}" -SessionsRoot "{3}"' -f `
        $StartSession, $Workbook, $SessionId, $SessionsRoot
    if ($VbaPassword) { $argStr += ' -VbaPassword "{0}"' -f $VbaPassword }
    return Start-Process pwsh -ArgumentList $argStr -PassThru -WindowStyle Hidden
}
