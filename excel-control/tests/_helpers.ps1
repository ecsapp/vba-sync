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
            if ($state.excel_pid -and [int]$state.excel_pid -gt 0) {
                $proc = Get-Process -Id ([int]$state.excel_pid) -ErrorAction SilentlyContinue
                if ($proc -and $proc.ProcessName -eq 'EXCEL') {
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                }
            }
            if ($state.pid -and [int]$state.pid -gt 0) {
                $hostProc = Get-Process -Id ([int]$state.pid) -ErrorAction SilentlyContinue
                if ($hostProc -and $hostProc.ProcessName -eq 'pwsh') {
                    Stop-Process -Id $hostProc.Id -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {}
    }
}
