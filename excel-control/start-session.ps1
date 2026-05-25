# excel-control session host. Long-running PowerShell process that holds
# an Excel COM instance open and exchanges events with an external agent
# over append-only JSONL files.
#
# Layout:
#   sessions/<SessionId>/
#     state.json        pid, workbook, status, last_command_offset, started_at
#     commands.jsonl    agent appends; harness consumes in order
#     events.jsonl      harness appends; agent watches
#     captures/         PNG screenshots referenced by events (Phase 4+)
#
# Phase 2 supports:
#   run_macro {id, cmd:"run_macro", name, args?}
#   close     {id, cmd:"close"}
#
# Unknown commands emit `unknown_command`. Modal dialogs raised by a macro
# will currently BLOCK the COM thread — Phase 3 wires the dialog watcher.
#
# Usage:
#   pwsh .\start-session.ps1 -Workbook X.xlsm -SessionId s1

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string]$Workbook,
    [Parameter(Mandatory=$true)] [string]$SessionId,
    # -SessionsRoot: where per-session runtime files live. When omitted the
    # default is computed below from the workbook path — a gitignored dir
    # beside the workbook, deliberately NOT under .claude/ (every read/write
    # under .claude/ triggers a permission prompt that cannot be allowlisted).
    [string]$SessionsRoot,
    [int]$PollMs = 250,
    # -Visible: show Excel on the desktop so you can watch the agent
    # work. Worksheet screenshots become meaningful (headless mode
    # renders mostly-empty main window). Side-effects: any UI element
    # the agent triggers steals foreground focus from other windows.
    [switch]$Visible,
    # -VbaPassword: password for a locked VBA project. Without it the
    # harness cannot sync/compile/test a password-protected VBA project.
    # Precedence: this param > $env:XC_VBA_PASSWORD > a gitignored
    # .vba-password file beside the harness scripts. Never commit it.
    [string]$VbaPassword
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'session-dialog-watcher.ps1')
. (Join-Path $PSScriptRoot 'capture.ps1')
. (Join-Path $PSScriptRoot 'unlock-vba-project.ps1')

# Win32 declaration to grab Excel's PID from its HWND
if (-not ([System.Management.Automation.PSTypeName]'XcSession.Win32').Type) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
namespace XcSession {
  public static class Win32 {
    [DllImport("user32.dll", SetLastError = true)] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  }
}
"@
}

$resolvedWb = (Resolve-Path -LiteralPath $Workbook).Path
if (-not $SessionsRoot) {
    # Beside the workbook (the repo root in the vba-sync layout), never
    # under .claude/. Workbook-relative is layout-independent — it does not
    # care whether this script runs from the deployed skill tree or the dev
    # tree.
    $SessionsRoot = Join-Path (Split-Path -Parent $resolvedWb) '.excel-control\sessions'
}

$sessionDir   = Join-Path $SessionsRoot $SessionId
$capturesDir  = Join-Path $sessionDir 'captures'
$stateFile    = Join-Path $sessionDir 'state.json'
$commandsFile = Join-Path $sessionDir 'commands.jsonl'
$eventsFile   = Join-Path $sessionDir 'events.jsonl'

foreach ($d in @($sessionDir, $capturesDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
}
foreach ($f in @($commandsFile, $eventsFile)) {
    if (-not (Test-Path $f)) { Set-Content -LiteralPath $f -Value '' -Encoding UTF8 -NoNewline }
}

$script:LastOffset      = 0
$script:Stop            = $false
$script:StartedAt       = (Get-Date).ToUniversalTime().ToString('o')
# Crash-detection state. ExcelDead flips once when either the per-loop
# Get-Process check or a pre-dispatch COM heartbeat sees the spawned Excel
# gone; the dispatch loop then accepts only `close` / `close_recovery` and
# emits `command_rejected` for anything else. LastCommandTime drives the
# zombie-host idle timeout in `crashed` (no agent reads => self-tear-down).
$script:ExcelDead       = $false
$script:LastCommandTime = [DateTime]::UtcNow
$script:CrashedIdleTimeoutMs = 600000  # 10 min
# session_status bookkeeping. Each Write-EventLine snaps the last event;
# the dispatch loop snaps the last command + the events-since counter
# at command_ack time. Test-ExcelAlive snaps the excel_crashed payload
# so session_status can surface it inline when excel_alive:false.
$script:LastCommandSnapshot    = $null
$script:LastEventSnapshot      = $null
$script:EventsSinceLastCommand = 0
$script:LastCrashedEvent       = $null
# instrument tracking — populated by an instrument-flagged run_macro so
# Test-ExcelAlive can auto-vacuum the log on excel_crashed. emitted gates
# the once-per-failure invariant (a heap-corruption sequence fires both
# macro_failed AND excel_crashed; both should not emit pre_crash_log).
$script:LogVacuumActive        = $null

# Resume offset if state.json says so (allows restart-after-crash to skip
# already-consumed commands)
if (Test-Path $stateFile) {
    try {
        $prev = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json
        if ($prev.last_command_offset) { $script:LastOffset = [int64]$prev.last_command_offset }
    } catch {}
}

function Write-EventLine([hashtable]$EventObj) {
    if (-not $EventObj.ContainsKey('ts')) {
        $EventObj['ts'] = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    }
    # session_status bookkeeping (sub-microsecond — script-scope assigns).
    # command_ack resets the since-counter; everything else bumps it.
    $script:LastEventSnapshot = @{ t = $EventObj.t; at = $EventObj.ts }
    if ($EventObj.t -eq 'command_ack') { $script:EventsSinceLastCommand = 0 }
    else                                { $script:EventsSinceLastCommand++ }
    $line = ConvertTo-Json -Compress -Depth 10 -InputObject $EventObj
    # Retry on transient IO contention — the dialog watcher writes the
    # same events.jsonl from another runspace. Without this a collision
    # throws and (ErrorActionPreference=Stop) crashes the whole session.
    for ($i = 0; $i -lt 5; $i++) {
        try { Add-Content -LiteralPath $eventsFile -Value $line -Encoding UTF8; return }
        catch { Start-Sleep -Milliseconds 30 }
    }
    Add-Content -LiteralPath $eventsFile -Value $line -Encoding UTF8
}

function Save-State([string]$Status) {
    $s = [ordered]@{
        pid                  = $PID
        excel_pid            = $script:ExcelPid
        workbook             = $Workbook
        session_id           = $SessionId
        status               = $Status
        visible              = [bool]$Visible
        started_at           = $script:StartedAt
        last_command_offset  = $script:LastOffset
        last_excel_check_ts  = $script:LastExcelCheckTs
    }
    # Write temp then rename — a reader must never catch state.json
    # empty, which a direct Set-Content leaves it while writing.
    $tmp = "$stateFile.tmp"
    Set-Content -LiteralPath $tmp -Value (ConvertTo-Json -Depth 5 -InputObject $s) -Encoding UTF8
    for ($i = 0; $i -lt 25; $i++) {
        try { [System.IO.File]::Move($tmp, $stateFile, $true); return }
        catch { Start-Sleep -Milliseconds 20 }
    }
}

# Reads new lines from commands.jsonl starting at $LastOffset. Returns
# array of (parsed object, raw line) pairs; updates $LastOffset to EOF.
function Read-NewCommands {
    $results = @()
    if (-not (Test-Path $commandsFile)) { return $results }
    $fs = [System.IO.File]::Open($commandsFile, 'Open', 'Read', 'ReadWrite')
    try {
        if ($script:LastOffset -gt $fs.Length) { $script:LastOffset = 0 }
        $fs.Position = $script:LastOffset
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 1024, $true)
        while (-not $sr.EndOfStream) {
            $line = $sr.ReadLine()
            if ($null -eq $line) { break }
            $trim = $line.Trim()
            if ($trim) {
                try {
                    $obj = $trim | ConvertFrom-Json
                    $results += [pscustomobject]@{ obj = $obj; raw = $trim }
                } catch {
                    # H5/M1: a malformed line (e.g. an un-escaped Windows
                    # path '\U...') still usually carries a parseable "id" —
                    # scrape it so the sender can correlate the failure
                    # instead of waiting forever for a result.
                    $scrapedId = $null
                    $idMatch = [regex]::Match($trim, '"id"\s*:\s*"([^"]*)"')
                    if ($idMatch.Success) { $scrapedId = $idMatch.Groups[1].Value }
                    Write-EventLine @{ t = 'command_error'; id = $scrapedId; error = "Invalid JSON: $($_.Exception.Message)"; raw = $trim }
                }
            }
        }
        $script:LastOffset = $fs.Position
        $sr.Dispose()
    } finally {
        $fs.Dispose()
    }
    return $results
}

# Invokes Excel.Application.Run with variable args. Uses InvokeMember so
# the macro name and args can be passed as a single array.
function Invoke-Macro($Xl, [string]$MacroName, [object[]]$MacroArgsArr) {
    $all = ,$MacroName + ($MacroArgsArr | ForEach-Object { $_ })
    return $Xl.GetType().InvokeMember(
        'Run',
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null, $Xl, $all)
}


# Returns the workbook's VBProject ready for VBComponents access, or throws
# an actionable error. Two distinct failure modes both surface over COM as a
# null/empty value with an unhelpful message:
#   1. $Wb.VBProject is $null  -> "Trust access to the VBA project object
#      model" is disabled in the Excel Trust Center.
#   2. $Wb.VBProject is non-null but .Protection = vbext_pp_locked (1) and
#      .VBComponents is $null -> the VBA project is password-locked. The
#      session auto-unlocks at startup when a password is supplied; this
#      path is the fallback when it was not (or unlock failed).
function Get-AccessibleVBProject($Wb) {
    $proj = $null
    try { $proj = $Wb.VBProject } catch {}
    if ($null -eq $proj) {
        throw "VBA project object model is not accessible. Enable File > " +
              "Options > Trust Center > Trust Center Settings > Macro " +
              "Settings > 'Trust access to the VBA project object model'."
    }
    $locked = $false
    try { $locked = ([int]$proj.Protection -ne 0) } catch {}
    if ($locked -or $null -eq $proj.VBComponents) {
        throw "VBA project is password-locked (Protection=vbext_pp_locked) " +
              "- VBComponents is inaccessible over COM. Supply the password " +
              "via -VbaPassword / `$env:XC_VBA_PASSWORD / a .vba-password file " +
              "so the session auto-unlocks the project at startup."
    }
    return $proj
}

# item 5 — drive vba-sync's own Import/Export (sheet-safe, unlike the
# naive sync_vba). Runs the addin's arg-less HarnessImport/HarnessExport
# and pre-arms the addin's success MsgBox so the run needs no round-trip.
function Invoke-VbaSyncOp($Xl, $Wb, $cmd) {
    $op = $cmd.cmd   # 'import' or 'export'
    try {
        $macro  = if ($op -eq 'export') { 'HarnessExport' } else { 'HarnessImport' }
        $okText = "$op completed successfully"

        # Ensure VBA Sync.xlam is loaded. Do not trust the Excel addin
        # registration to locate it — a stale registered path breaks
        # Application.Run name resolution; open it from the install path.
        $xlamWb = $null
        foreach ($w in $Xl.Workbooks) { if ($w.Name -eq 'VBA Sync.xlam') { $xlamWb = $w; break } }
        if (-not $xlamWb) {
            $xlamPath = Join-Path $env:APPDATA 'Microsoft\AddIns\VBA Sync.xlam'
            if (-not (Test-Path -LiteralPath $xlamPath)) {
                throw "VBA Sync.xlam not found at $xlamPath"
            }
            [void]$Xl.Workbooks.Open($xlamPath, [Type]::Missing, $false, [Type]::Missing,
                                     [Type]::Missing, [Type]::Missing, $true)
        }

        # The addin targets ActiveWorkbook (TargetWB) — make it the session wb.
        $Wb.Activate()

        # Pre-arm the addin's completion MsgBox (item 4) so the run is
        # hands-free. Addin error dialogs are deliberately NOT armed — they
        # surface as dialog_appeared and the op then reports failure.
        $armId  = "$($cmd.id)-arm"
        $armCmd = @{ id = $armId; cmd = 'arm_response'; match = @{ text = $okText }; button = 'OK' } |
                  ConvertTo-Json -Compress
        Add-Content -LiteralPath $commandsFile -Value $armCmd -Encoding UTF8

        if ($script:Watcher) { $script:Watcher.State.MacroRunning = $true }
        $t0 = [DateTime]::UtcNow
        try {
            $Xl.Run("'VBA Sync.xlam'!modSync.$macro")
        } finally {
            if ($script:Watcher) { $script:Watcher.State.MacroRunning = $false }
        }
        $durMs = [int]([DateTime]::UtcNow - $t0).TotalMilliseconds

        # Success iff the armed completion dialog fired. Poll — the watcher
        # writes dialog_auto_responded around the moment Xl.Run returns, so
        # a single scan here can race ahead of that write.
        $ok = $false
        $deadline = [DateTime]::UtcNow.AddSeconds(6)
        while (-not $ok -and [DateTime]::UtcNow -lt $deadline) {
            try {
                foreach ($line in (Get-Content -LiteralPath $eventsFile)) {
                    if ($line -match '"t":"dialog_auto_responded"' -and
                        $line -match ([regex]::Escape('"rule_id":"' + $armId + '"'))) { $ok = $true; break }
                }
            } catch {}
            if (-not $ok) { Start-Sleep -Milliseconds 200 }
        }

        if ($ok) {
            Write-EventLine @{ t = "${op}_completed"; id = $cmd.id; duration_ms = $durMs }
        } else {
            Write-EventLine @{ t = "${op}_failed"; id = $cmd.id; duration_ms = $durMs
                error = "vba-sync $op did not report success — an addin error dialog likely surfaced (see dialog_appeared)" }
        }
    } catch {
        Write-EventLine @{ t = "${op}_failed"; id = $cmd.id; error = $_.Exception.Message }
    }
}

# Crash detection — both gates of items A and C in the design.
# Returns $true when Excel is gone (either path), flipping $script:ExcelDead
# and emitting `excel_crashed` exactly once. Cheap: a Get-Process by PID and
# (optionally) a single `$xl.Hwnd` get; the COM read is what catches a race
# where the process is up but the COM ref has gone disconnected.
function Test-ExcelAlive($Xl, [switch]$ComHeartbeat) {
    if ($script:ExcelDead) { return $false }
    $reason = $null
    if ($script:ExcelPid) {
        $p = $null
        try { $p = Get-Process -Id $script:ExcelPid -ErrorAction SilentlyContinue } catch {}
        if (-not $p) { $reason = 'process exited' }
    }
    if (-not $reason -and $ComHeartbeat -and $Xl) {
        try { $null = $Xl.Hwnd }
        catch { $reason = "com disconnected: $($_.Exception.Message)" }
    }
    if ($reason) {
        $script:ExcelDead = $true
        $crashedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        # Snap for session_status (surfaced inline when excel_alive:false
        # so the agent does not have to grep events.jsonl backward).
        $script:LastCrashedEvent = @{
            pid = $script:ExcelPid
            reason = $reason
            detected_at = $crashedAt
        }
        Write-EventLine @{
            t = 'excel_crashed'
            pid = $script:ExcelPid
            reason = $reason
            detected_at = $crashedAt
        }
        # Pre-crash log auto-vacuum (once per active instrument).
        Invoke-PreCrashLogEmit -TriggerEvent 'excel_crashed'
        Save-State 'crashed'
        return $false
    }
    $script:LastExcelCheckTs = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    return $true
}

# run_macro_loop support — three helpers + the loop driver itself.

# Deep-clone a command object substituting {{i}} and {{n}} with the
# current 1-indexed iteration number on every string value. Uses a JSON
# round-trip so nested arrays/objects come out as a clean independent
# copy (no shared references between iterations).
function Expand-LoopVars($obj, [int]$Iter) {
    if ($null -eq $obj) { return $obj }
    $json = ConvertTo-Json -Compress -Depth 20 -InputObject $obj
    $json = $json -replace '\{\{i\}\}', [string]$Iter
    $json = $json -replace '\{\{n\}\}', [string]$Iter
    return ($json | ConvertFrom-Json)
}

# Scan events.jsonl from a given byte offset forward for any event whose
# t-field matches the supplied type list OR (when -AnyFailedSuffix) any
# t ending in '_failed'. Returns the first matching event as a
# pscustomobject, or $null. Glob-prefilters the raw line so we only
# parse-to-JSON on a hit. Same I/O pattern as Read-NewCommands.
function Find-StopOnEvent {
    param(
        [Parameter(Mandatory)] [string]$EventsFile,
        [Parameter(Mandatory)] [long]$FromOffset,
        [string[]]$StopTypes = @(),
        [switch]$AnyFailedSuffix
    )
    if (-not (Test-Path $EventsFile)) { return $null }
    $fs = [System.IO.File]::Open($EventsFile, 'Open', 'Read', 'ReadWrite')
    try {
        if ($FromOffset -gt $fs.Length) { return $null }
        $fs.Position = $FromOffset
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 1024, $true)
        while (-not $sr.EndOfStream) {
            $line = $sr.ReadLine()
            if (-not $line) { continue }
            $hit = $false
            foreach ($t in $StopTypes) {
                if ($line -like ('*"t":"' + $t + '"*')) { $hit = $true; break }
            }
            # *_failed pattern: agreed with p6 (collab seq 10) — hardcoding a
            # list goes stale; the harness convention is *_failed for every
            # failure event, so one regex stays current as new commands ship.
            if (-not $hit -and $AnyFailedSuffix -and $line -match '"t":"[a-z_]+_failed"') {
                $hit = $true
            }
            if ($hit) {
                try { return ($line | ConvertFrom-Json) } catch { return $null }
            }
        }
        $sr.Dispose()
        return $null
    } finally { $fs.Dispose() }
}

# Run-macro-loop driver. Key invariants (full design notes in the
# cross-project-primitives collab PLAN-run-macro-loop.md):
# - One macro_loop_progress per iteration (streaming); terminal
#   macro_loop_completed with per_iteration[].
# - Loop-scoped armed_responses via state.LoopArmQueue / LoopDisarmQueue
#   (watcher owns $armedRules; queues route mutation through it).
# - {{i}}/{{n}} interpolation per-iteration via Expand-LoopVars.
# - stop_on matched via Find-StopOnEvent's offset-based scan.
# - between_iterations failure detection uses *_failed pattern (any new
#   harness command's failure event surfaces automatically).
# - Mid-loop Test-ExcelAlive — a crashed Excel exits with
#   stop_reason:session_crashed (not stop_on, even if excel_crashed is
#   listed — the crash check runs FIRST per iteration).
function Invoke-MacroLoop($Xl, $Wb, $cmd) {
    $loopId      = $cmd.id
    $iterations  = 0
    try { $iterations = [int]$cmd.iterations } catch {}
    if ($iterations -lt 1) {
        Write-EventLine @{ t = 'macro_loop_failed'; id = $loopId
            error = 'iterations must be a positive integer' }
        return
    }
    $name         = [string]$cmd.name
    if (-not $name) {
        Write-EventLine @{ t = 'macro_loop_failed'; id = $loopId; error = 'name is required' }
        return
    }
    $betweenCmds    = @()
    if ($cmd.between_iterations) { $betweenCmds = @($cmd.between_iterations) }
    $betweenOnError = if ($cmd.between_iterations_on_error) { [string]$cmd.between_iterations_on_error } else { 'stop' }
    $stopOn         = @()
    if ($cmd.stop_on) { $stopOn = @($cmd.stop_on) }
    $armed          = @()
    if ($cmd.armed_responses) { $armed = @($cmd.armed_responses) }
    $argsArr        = @()
    if ($null -ne $cmd.args) { $argsArr = @($cmd.args) }

    # Queue loop-scoped armed_responses through the watcher's LoopArmQueue.
    if ($armed.Count -gt 0 -and $script:Watcher) {
        $armIdx = 0
        foreach ($a in $armed) {
            $armIdx++
            $rep = 1
            if ($a.repeat) { try { $rep = [int]$a.repeat } catch {} }
            if ($rep -lt 1) { $rep = 1 }
            $rule = @{
                Id         = "${loopId}-arm-${armIdx}"
                Kind       = $(if ($a.control) { 'form_control' } else { 'response' })
                Match      = $a.match
                Button     = $a.button
                Control    = $a.control
                Value      = $a.value
                Repeat     = $rep
                loop_owner = $loopId
            }
            [void]$script:Watcher.State.LoopArmQueue.Add($rule)
        }
    }

    $perIter    = New-Object System.Collections.ArrayList
    $stopReason = 'completed'
    $stopEvent  = $null
    $stopAtIter = $null
    $loopStart  = [DateTime]::UtcNow

    for ($i = 1; $i -le $iterations; $i++) {
        $script:LastCommandTime = [DateTime]::UtcNow

        # Mid-loop crash check — bail if the spawned Excel went away.
        if (-not (Test-ExcelAlive -Xl $Xl)) {
            $stopReason = 'session_crashed'
            $stopAtIter = $i
            break
        }

        $preMacroOffset = 0
        try { $preMacroOffset = (Get-Item -LiteralPath $eventsFile).Length } catch {}

        $iterStart  = [DateTime]::UtcNow
        $iterRecord = [ordered]@{ i = $i }
        $expandedArgs = @()
        if ($argsArr.Count -gt 0) {
            $expandedArgs = @(Expand-LoopVars -obj $argsArr -Iter $i)
        }
        $macroRef = if ($name -like '*!*') { $name } else { "'$($Wb.Name)'!$name" }
        try {
            $result = switch ($expandedArgs.Count) {
                0 { $Xl.Run($macroRef) }
                1 { $Xl.Run($macroRef, $expandedArgs[0]) }
                2 { $Xl.Run($macroRef, $expandedArgs[0], $expandedArgs[1]) }
                3 { $Xl.Run($macroRef, $expandedArgs[0], $expandedArgs[1], $expandedArgs[2]) }
                4 { $Xl.Run($macroRef, $expandedArgs[0], $expandedArgs[1], $expandedArgs[2], $expandedArgs[3]) }
                5 { $Xl.Run($macroRef, $expandedArgs[0], $expandedArgs[1], $expandedArgs[2], $expandedArgs[3], $expandedArgs[4]) }
                default { Invoke-Macro -Xl $Xl -MacroName $macroRef -MacroArgsArr $expandedArgs }
            }
            $iterRecord.result      = $result
            $iterRecord.duration_ms = [int]([DateTime]::UtcNow - $iterStart).TotalMilliseconds
        } catch {
            $iterRecord.error       = $_.Exception.Message
            $iterRecord.duration_ms = [int]([DateTime]::UtcNow - $iterStart).TotalMilliseconds
        }

        # stop_on scan: events emitted since the macro started.
        if ($stopOn.Count -gt 0) {
            $match = Find-StopOnEvent -EventsFile $eventsFile -FromOffset $preMacroOffset -StopTypes $stopOn
            if ($match) {
                $iterRecord.stop = $true
                [void]$perIter.Add($iterRecord)
                Write-EventLine @{ t = 'macro_loop_progress'; id = $loopId; i = $i; of = $iterations
                    duration_ms = $iterRecord.duration_ms; result = $iterRecord.result; error = $iterRecord.error
                    stop = $true }
                $stopReason = 'stop_on'
                $stopEvent  = $match
                $stopAtIter = $i
                break
            }
        }

        [void]$perIter.Add($iterRecord)
        Write-EventLine @{ t = 'macro_loop_progress'; id = $loopId; i = $i; of = $iterations
            duration_ms = $iterRecord.duration_ms; result = $iterRecord.result; error = $iterRecord.error }

        # between_iterations — runs between but not after the last iteration.
        if ($i -lt $iterations -and $betweenCmds.Count -gt 0) {
            $betweenFailed = $false
            foreach ($subCmdRaw in $betweenCmds) {
                $expandedSub = Expand-LoopVars -obj $subCmdRaw -Iter $i
                $preBetweenOffset = 0
                try { $preBetweenOffset = (Get-Item -LiteralPath $eventsFile).Length } catch {}
                try { Invoke-SessionCommand -Xl $Xl -Wb $Wb -cmd $expandedSub -Suppress } catch {}
                if ($betweenOnError -eq 'stop') {
                    $betweenMatch = Find-StopOnEvent -EventsFile $eventsFile `
                        -FromOffset $preBetweenOffset -AnyFailedSuffix
                    if ($betweenMatch) {
                        $betweenFailed = $true
                        $stopEvent     = $betweenMatch
                        break
                    }
                }
            }
            if ($betweenFailed) {
                $stopReason = 'between_error'
                $stopAtIter = $i
                break
            }
        }
    }

    # Signal disarm to the watcher; the response_disarmed events will
    # stream on the watcher's next tick(s).
    if ($script:Watcher) {
        try { [void]$script:Watcher.State.LoopDisarmQueue.Add($loopId) } catch {}
    }

    $report = [ordered]@{
        t                         = 'macro_loop_completed'
        id                        = $loopId
        iterations_requested      = $iterations
        iterations_completed      = $perIter.Count
        stop_reason               = $stopReason
        stop_event                = $stopEvent
        stop_at_iteration         = $stopAtIter
        duration_ms               = [int]([DateTime]::UtcNow - $loopStart).TotalMilliseconds
        per_iteration             = $perIter
    }
    Write-EventLine $report
}

# analyze_vba support — walks the live VBE once per call (no disk
# fallback, per the collab seq-4 contract), returns a flat catalogue
# of every Sub/Function definition. Consumed by definition_of and
# callers_of; v2 callees_of / references_to will reuse it.
function Get-VbaProcedures($Proj) {
    $procs = New-Object System.Collections.ArrayList
    foreach ($comp in $Proj.VBComponents) {
        if ($comp.CodeModule.CountOfLines -le 0) { continue }
        $kindStr = switch ([int]$comp.Type) {
            1   { 'standard' } 2 { 'class' } 3 { 'userform' } 100 { 'document' }
            default { "type$([int]$comp.Type)" }
        }
        $cm    = $comp.CodeModule
        $total = [int]$cm.CountOfLines
        $headers = New-Object System.Collections.ArrayList
        for ($ln = 1; $ln -le $total; $ln++) {
            # VBA line-continuation: a logical line can span multiple
            # physical lines, joined by `[space]_` at end-of-line. The
            # procedure-header regex needs the *joined* line to match
            # Sub Name(...) where the closing paren may live several
            # lines down. Without this, multi-line signatures are
            # invisible to definition_of / callers_of (silent false
            # negative — see HARNESS_IMPROVEMENTS P1.5).
            $line = $cm.Lines($ln, 1)
            $declLine = $ln
            $joined = $line
            $peek = $ln
            while ($joined -match '\s_\s*$' -and $peek -lt $total) {
                $peek++
                $next = $cm.Lines($peek, 1)
                # Strip the trailing ` _` from the prior chunk and
                # concatenate with a single space; preserves the regex's
                # \s* boundaries around the signature.
                $joined = ($joined -replace '\s_\s*$', ' ') + $next
            }
            if ($joined -match '^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\((.*?)\)') {
                $vis = "$($Matches[1])".Trim()
                $isPublic = ($vis -eq '') -or ($vis -ieq 'Public')
                [void]$headers.Add(@{
                    Module        = $comp.Name
                    Kind          = $Matches[2].ToLower()
                    Name          = $Matches[3]
                    Line          = $declLine
                    Signature     = $joined.Trim()
                    Args          = $Matches[4].Trim()
                    Public        = $isPublic
                    ComponentKind = $kindStr
                })
                # Skip past the continuation lines we already consumed
                # so we don't re-scan them as separate procedures.
                if ($peek -gt $ln) { $ln = $peek }
            }
        }
        # Assign body_start / body_end (line before next header, or end of module)
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $h = $headers[$i]
            $h.BodyStart = $h.Line
            $h.BodyEnd   = if ($i -lt $headers.Count - 1) { $headers[$i + 1].Line - 1 } else { $total }
            [void]$procs.Add([pscustomobject]$h)
        }
    }
    return $procs
}

function Invoke-VbaDefinitionOf($Proj, [string]$Name, $Cmd) {
    $procs = Get-VbaProcedures $Proj
    # NOT $matches — that's a PowerShell automatic variable set by -match
    # / regex operators; reusing the name would silently clobber the
    # implicit match groups inside this function and is undefined behaviour
    # for the assignment.
    $defs = @($procs | Where-Object { $_.Name -ieq $Name } | ForEach-Object {
        [ordered]@{
            module    = $_.Module
            kind      = $_.Kind
            name      = $_.Name           # canonical casing from the VBE
            line      = $_.Line
            signature = $_.Signature
            public    = $_.Public
            args      = $_.Args
        }
    })
    Write-EventLine @{ t = 'vba_analysis_result'; id = $Cmd.id
        op = 'definition_of'; query = $Name; matches = $defs }
}

function Invoke-VbaReadModule($Proj, [string]$Name, $Cmd) {
    $comp = $null
    foreach ($c in $Proj.VBComponents) {
        if ($c.Name -ieq $Name) { $comp = $c; break }
    }
    if (-not $comp) {
        Write-EventLine @{ t = 'vba_analysis_result'; id = $Cmd.id
            op = 'read_module'; query = $Name; module = $null
            note = "no module named '$Name'" }
        return
    }
    $cm  = $comp.CodeModule
    $src = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { '' }
    $kindStr = switch ([int]$comp.Type) {
        1   { 'standard' } 2 { 'class' } 3 { 'userform' } 100 { 'document' }
        default { "type$([int]$comp.Type)" }
    }
    Write-EventLine @{ t = 'vba_analysis_result'; id = $Cmd.id
        op = 'read_module'; query = $Name
        module = [ordered]@{
            name       = $comp.Name
            kind       = $kindStr
            line_count = [int]$cm.CountOfLines
            source     = $src
        } }
}

function Invoke-VbaCallersOf($Proj, [string]$Name, $Cmd) {
    $procs  = Get-VbaProcedures $Proj
    $needle = [regex]::Escape($Name)
    $callPat = [regex]::new('\b' + $needle + '\b', 'IgnoreCase')

    # Cache CodeModule per module name — catalogue stays pure-data,
    # body-scan looks the component up once per module via name.
    $modCache = @{}
    foreach ($p in $procs) {
        if (-not $modCache.ContainsKey($p.Module)) {
            $cm = $null
            foreach ($c in $Proj.VBComponents) { if ($c.Name -eq $p.Module) { $cm = $c.CodeModule; break } }
            $modCache[$p.Module] = $cm
        }
    }

    $callers = New-Object System.Collections.ArrayList
    foreach ($p in $procs) {
        # A procedure is never reported as a caller of itself. The
        # function return-value assignment (`BuildSql = expr` inside
        # the `BuildSql` function body) would otherwise count as a
        # match, as would the defining `Sub Foo(...)` line itself.
        # Trade-off: this suppresses true recursion detection — rare
        # in business VBA, and `callers_of` is intended to answer "who
        # else uses this", not "does it recurse".
        if ($p.Name -ieq $Name) { continue }
        $cm = $modCache[$p.Module]
        if (-not $cm) { continue }
        $hits = @()
        for ($ln = $p.BodyStart; $ln -le $p.BodyEnd; $ln++) {
            $line = $cm.Lines($ln, 1)
            $trim = $line.TrimStart()
            if ($trim.StartsWith("'")) { continue }      # comment-only line
            if ($callPat.IsMatch($line)) { $hits += $ln }
        }
        if ($hits.Count -gt 0) {
            [void]$callers.Add([ordered]@{
                module     = $p.Module
                kind       = $p.Kind
                name       = $p.Name
                line       = $p.Line
                call_lines = $hits
            })
        }
    }
    Write-EventLine @{ t = 'vba_analysis_result'; id = $Cmd.id
        op = 'callers_of'; query = $Name; callers = $callers }
}

# Pre-crash logger support — xcHarness_Log VBA module is injected
# lazily, only when an agent flags a run_macro with instrument:{...}.
# Workbooks the agent never instruments stay unmutated; xlam regression
# tests stay clean.

# Module is named xcHarnessLog (no underscore) so the contained Sub
# xcHarness_Log (with underscore) is unambiguous to Application.Run.
# Same-name module + sub trips Excel's "Cannot run the macro" error
# because the lookup ambiguates module vs procedure.
$script:XcHarnessLogModuleName  = 'xcHarnessLog'
$script:XcHarnessLogVbaTemplate = @'
Attribute VB_Name = "xcHarnessLog"
Option Explicit

' Injected by excel-control harness (lazy, on first instrumented
' run_macro). Do not edit manually. File-append logger for pre-crash
' state capture. Agent VBA calls: xcHarness_Log "any message".

Private Const xcHarness_LogPath As String = "{{LOG_PATH}}"

Public Sub xcHarness_Log(ByVal s As String)
    On Error Resume Next
    Dim fn As Integer
    fn = FreeFile
    Open xcHarness_LogPath For Append As #fn
    Print #fn, Format(Now, "yyyy-mm-dd hh:nn:ss"); "."; _
               Format(Timer * 1000 Mod 1000, "000"); " "; s
    Close #fn
End Sub
'@

# Read the const value from an existing xcHarness_Log module — needed
# for stale-path detection when a workbook is re-opened in a new session.
function Get-XcLogConstPath($Component) {
    if (-not $Component) { return $null }
    $cm = $Component.CodeModule
    if (-not $cm -or $cm.CountOfLines -le 0) { return $null }
    for ($ln = 1; $ln -le [int]$cm.CountOfLines; $ln++) {
        $line = $cm.Lines($ln, 1)
        if ($line -match 'Private\s+Const\s+xcHarness_LogPath\s+As\s+String\s*=\s*"([^"]*)"') {
            return $Matches[1]
        }
    }
    return $null
}

# Inject the module if not already present. Idempotent. Returns
# @{ path; injected; stale_path? } — stale_path is the OLD baked-in
# path when the module already existed with a different path (a
# re-opened workbook from a prior session). Caller surfaces that as an
# instrument_stale_path warning event.
function Ensure-XcHarnessLogModule($Wb) {
    $intendedPath = Join-Path $script:SessionDir 'macro.log'
    try {
        $proj = Get-AccessibleVBProject $Wb
    } catch {
        throw "xcHarness_Log injection requires VBProject access: $($_.Exception.Message)"
    }
    foreach ($c in $proj.VBComponents) {
        if ($c.Name -eq $script:XcHarnessLogModuleName) {
            $oldPath = Get-XcLogConstPath -Component $c
            $stale = $null
            if ($oldPath -and $oldPath -ne $intendedPath) { $stale = $oldPath }
            return @{ path = $oldPath; injected = $false; stale_path = $stale }
        }
    }
    # Inject via VBComponents.Import on a temp .bas file — same path
    # sync_vba uses, proven to work. The earlier Add(1) +
    # CodeModule.AddFromString sequence left Application.Run unable to
    # resolve other macros in the project ("macro may not be available")
    # — likely a VBE compile-state interaction; Import does not trip it.
    $src = $script:XcHarnessLogVbaTemplate -replace '\{\{LOG_PATH\}\}', ($intendedPath -replace '\\', '\\')
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString('N') + '_xcHarness_Log.bas')
    Set-Content -LiteralPath $tmp -Value $src -Encoding ASCII
    try {
        [void]$proj.VBComponents.Import($tmp)
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
    return @{ path = $intendedPath; injected = $true; stale_path = $null }
}

# Read the last $TailLines lines of the log file. Returns
# @{ lines, total, truncated }. Missing/empty file yields empty array.
function Read-XcLogTail([string]$Path, [int]$TailLines) {
    if (-not (Test-Path -LiteralPath $Path)) { return @{ lines = @(); total = 0; truncated = $false } }
    $all = @()
    try { $all = @(Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction SilentlyContinue) } catch {}
    $total = $all.Count
    if ($total -le $TailLines) { return @{ lines = $all; total = $total; truncated = $false } }
    return @{ lines = $all[($total - $TailLines)..($total - 1)]; total = $total; truncated = $true }
}

# Emit pre_crash_log if instrumentation is active for $TriggerEvent and
# this active-instrument hasn't already emitted (once-per-failure
# invariant — a heap-corruption sequence fires both macro_failed AND
# excel_crashed; the second is a no-op). Clears
# $script:LogVacuumActive.emitted to true; caller is responsible for
# nulling the whole active-instrument record on macro completion.
function Invoke-PreCrashLogEmit([string]$TriggerEvent) {
    $a = $script:LogVacuumActive
    if (-not $a)          { return }
    if ($a.emitted)       { return }
    if ($TriggerEvent -notin @($a.vacuum_on)) { return }
    $tail = Read-XcLogTail -Path $a.path -TailLines $a.tail_lines
    Write-EventLine @{
        t              = 'pre_crash_log'
        id             = $a.id
        macro          = $a.macro
        trigger        = $TriggerEvent
        log_file       = $a.path
        log_lines      = $tail.lines
        lines_returned = $tail.lines.Count
        total_lines    = $tail.total
        tail_truncated = $tail.truncated
    }
    $script:LogVacuumActive.emitted = $true
}

function Invoke-SessionCommand($Xl, $Wb, $cmd, [switch]$Suppress) {
    # -Suppress: caller (run_macro_loop) owns the outer command_ack +
    # Save-State busy/ready boundary. The switch body still emits its
    # natural events (range_written, calculated, …) so the agent sees
    # every sub-step; only the dispatcher-layer ack + state churn is
    # skipped. The switch body itself does not read $Suppress — it is
    # the caller's flag.
    switch ($cmd.cmd) {
        'respond_dialog' {
            # Owned by the dialog watcher (separate runspace). Main loop
            # already skipped these — this branch is here as a safety net
            # in case the dispatch routing changes.
            return
        }
        'set_form_control' {
            # Owned by the dialog watcher (separate runspace) — see above.
            return
        }
        'run_macro' {
            $start = [DateTime]::UtcNow
            # Drain the VBE Immediate window before running so we only emit
            # Debug.Print events for output produced by THIS macro.
            try {
                $cm = $Xl.VBE.Windows.Item('Immediate').CodePane.CodeModule
                if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
            } catch {}

            # Pre-crash instrumentation (item 4 of cross-project-primitives).
            # cmd.instrument:{log:true, tail_lines?, vacuum_on?[]} → inject
            # xcHarness_Log if absent + track active-instrument so the
            # post-fail / excel_crashed paths can auto-vacuum the log file.
            if ($cmd.instrument -and $cmd.instrument.log) {
                $tail = if ($cmd.instrument.tail_lines) { [int]$cmd.instrument.tail_lines } else { 200 }
                $vac  = if ($cmd.instrument.vacuum_on)  { @($cmd.instrument.vacuum_on) }
                         else { @('macro_failed','excel_crashed') }
                try {
                    $inj = Ensure-XcHarnessLogModule -Wb $Wb
                    if ($inj.stale_path) {
                        Write-EventLine @{ t = 'instrument_stale_path'; id = $cmd.id
                            old_path = $inj.stale_path; intended_path = (Join-Path $script:SessionDir 'macro.log')
                            note = "xcHarness_Log module already present from a prior session; the baked xcHarness_LogPath const still points at the old log file. Module left unchanged (immutable contract)." }
                    }
                    $script:LogVacuumActive = @{
                        id         = $cmd.id
                        macro      = $cmd.name
                        path       = $inj.path
                        tail_lines = $tail
                        vacuum_on  = $vac
                        emitted    = $false
                    }
                    if ($inj.injected) {
                        Write-EventLine @{ t = 'instrument_active'; id = $cmd.id
                            log_file = $inj.path; injected = $true; tail_lines = $tail
                            vacuum_on = $vac }
                    } else {
                        Write-EventLine @{ t = 'instrument_active'; id = $cmd.id
                            log_file = $inj.path; injected = $false; tail_lines = $tail
                            vacuum_on = $vac }
                    }
                } catch {
                    Write-EventLine @{ t = 'instrument_failed'; id = $cmd.id; error = $_.Exception.Message }
                    $script:LogVacuumActive = $null
                }
            }

            try {
                $macroArgs = @()
                if ($null -ne $cmd.args) { $macroArgs = @($cmd.args) }
                # Optional: activate a specific workbook before the macro call
                # (some macros use ActiveWorkbook to find their target).
                if ($cmd.activate_workbook) {
                    try { $Xl.Workbooks.Item($cmd.activate_workbook).Activate() } catch {}
                }
                # cmd.name can be either a bare name ("PopMsgBox") which gets
                # qualified with the session's primary workbook, or an
                # already-qualified ref ("'AddinName.xlam'!modX.Foo") which
                # passes through unchanged. Detection: contains '!'.
                $macroRef = if ($cmd.name -like '*!*') { $cmd.name } else { "'$($Wb.Name)'!$($cmd.name)" }
                # Direct PowerShell COM call — InvokeMember overload
                # resolution was failing with "Parameter not optional"
                # when passing args that VBA's Variant parameters
                # accept. Splat by arity (most macros take 0-5 args).
                # Flag the macro in flight (item 2 deferred-modal rescue),
                # and whether a runtime error should drop into break mode
                # for capture instead of auto-End (item 3 debug_on_error).
                if ($script:Watcher) {
                    $script:Watcher.State.MacroRunning = $true
                    $script:Watcher.State.DebugOnError = [bool]$cmd.debug_on_error
                }
                try {
                    $result = switch ($macroArgs.Count) {
                        0 { $Xl.Run($macroRef) }
                        1 { $Xl.Run($macroRef, $macroArgs[0]) }
                        2 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1]) }
                        3 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1], $macroArgs[2]) }
                        4 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1], $macroArgs[2], $macroArgs[3]) }
                        5 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1], $macroArgs[2], $macroArgs[3], $macroArgs[4]) }
                        default { Invoke-Macro -Xl $Xl -MacroName $macroRef -MacroArgsArr $macroArgs }
                    }
                } finally {
                    if ($script:Watcher) {
                        $script:Watcher.State.MacroRunning = $false
                        $script:Watcher.State.DebugOnError = $false
                    }
                }
                # Emit one debug_print event per non-empty line.
                try {
                    $cm = $Xl.VBE.Windows.Item('Immediate').CodePane.CodeModule
                    if ($cm.CountOfLines -gt 0) {
                        $text = $cm.Lines(1, $cm.CountOfLines)
                        foreach ($line in ($text -split "[`r`n]+")) {
                            if ($line.Length -gt 0) {
                                Write-EventLine @{ t = 'debug_print'; during = $cmd.id; text = $line }
                            }
                        }
                    }
                } catch {}
                Write-EventLine @{
                    t = 'macro_completed'
                    id = $cmd.id
                    name = $cmd.name
                    result = $result
                    duration_ms = [int]([DateTime]::UtcNow - $start).TotalMilliseconds
                }
                # Clean macro success — instrument's done its job; clear
                # the active record so the next failure (if any) is not
                # mis-attributed to this run_macro.
                $script:LogVacuumActive = $null
            } catch {
                # If the dialog watcher already emitted a runtime_error
                # for this macro, that's the structured signal. The
                # COMException here is the same error coming back through
                # COM after the watcher auto-clicked End. We still emit
                # macro_failed for completeness (the agent can correlate
                # by id).
                $ex = $_.Exception
                $msg = $ex.Message
                $cur = $ex
                while ($cur.InnerException) { $cur = $cur.InnerException; if ($cur.Message) { $msg = $cur.Message } }
                Write-EventLine @{
                    t = 'macro_failed'
                    id = $cmd.id
                    name = $cmd.name
                    error = $msg
                    error_type = $ex.GetType().FullName
                    hresult = $ex.HResult
                }
                # Auto-vacuum (once per active instrument — if excel_crashed
                # fires next via Test-ExcelAlive the Invoke-PreCrashLogEmit
                # there is a no-op because emitted is now true).
                Invoke-PreCrashLogEmit -TriggerEvent 'macro_failed'
                $script:LogVacuumActive = $null
            }
        }
        'run_macro_loop' {
            Invoke-MacroLoop -Xl $Xl -Wb $Wb -cmd $cmd
        }
        'screenshot' {
            try {
                $target = $cmd.target
                if (-not $target) { $target = 'window' }
                $outPath = Join-Path $script:CapturesDir "shot_$($cmd.id).png"
                switch -Regex ($target) {
                    '^window$' {
                        # Excel main window
                        $hwnd = [int64]$Xl.Hwnd
                        Save-WindowImage -Hwnd $hwnd -Path $outPath | Out-Null
                    }
                    '^worksheet:(.+)$' {
                        $sheetName = $Matches[1]
                        $sheet = $Wb.Sheets.Item($sheetName)
                        $sheet.Activate()
                        Start-Sleep -Milliseconds 200
                        $hwnd = [int64]$Xl.Hwnd
                        Save-WindowImage -Hwnd $hwnd -Path $outPath | Out-Null
                    }
                    '^(dialog|form):(.+)$' {
                        $dlgId = $Matches[2]
                        $info  = $script:Watcher.State.DialogInfo[$dlgId]
                        if (-not $info) { throw "Unknown dialog/form id: $dlgId" }
                        Save-WindowImage -Hwnd $info.Hwnd -Path $outPath | Out-Null
                    }
                    default { throw "Unknown screenshot target: $target" }
                }
                $bmp = New-Object System.Drawing.Bitmap $outPath
                Write-EventLine @{
                    t = 'screenshot_captured'
                    id = $cmd.id
                    target = $target
                    path = $outPath
                    width = $bmp.Width
                    height = $bmp.Height
                }
                $bmp.Dispose()
            } catch {
                Write-EventLine @{ t = 'screenshot_failed'; id = $cmd.id; target = $cmd.target; error = $_.Exception.Message }
            }
        }
        'compile_check' {
            try {
                # CommandBar control id 578 = "Compile VBAProject"
                $btn = $Xl.VBE.CommandBars.FindControl([Type]::Missing, 578)
                if (-not $btn) { throw "Could not find Compile VBAProject control (id 578)" }
                # "Compile VBAProject" is disabled when the project is already
                # fully compiled. Execute() on a disabled control throws
                # "Unexpected HRESULT" — so a disabled control is success:
                # nothing to compile means no compile errors.
                if (-not $btn.Enabled) {
                    Write-EventLine @{ t = 'compile_result'; id = $cmd.id; ok = $true; note = 'already compiled' }
                    return
                }
                $btn.Execute()
                Start-Sleep -Milliseconds 600
                # If compile failed, VBE.ActiveCodePane is positioned on the error line.
                $pane = $null
                try { $pane = $Xl.VBE.ActiveCodePane } catch {}
                if (-not $pane) {
                    Write-EventLine @{ t = 'compile_result'; id = $cmd.id; ok = $true }
                } else {
                    $module = $pane.CodeModule.Name
                    $sLine = 0; $sCol = 0; $eLine = 0; $eCol = 0
                    [void]$pane.GetSelection([ref]$sLine, [ref]$sCol, [ref]$eLine, [ref]$eCol)
                    $start = [Math]::Max(1, $sLine - 2)
                    $end   = [Math]::Min($pane.CodeModule.CountOfLines, $sLine + 2)
                    $ctx = @()
                    for ($i = $start; $i -le $end; $i++) {
                        $ctx += ('{0}: {1}' -f $i, $pane.CodeModule.Lines($i, 1))
                    }
                    Write-EventLine @{
                        t = 'compile_result'
                        id = $cmd.id
                        ok = $false
                        module = $module
                        line = $sLine
                        column = $sCol
                        source_context = $ctx
                    }
                }
            } catch {
                Write-EventLine @{ t = 'compile_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'sync_vba' {
            try {
                $sourceDir = $cmd.source_dir
                if (-not $sourceDir) { throw "sync_vba requires source_dir" }
                $resolvedSrc = (Resolve-Path -LiteralPath $sourceDir).Path

                # Issue 3: a wrong source_dir otherwise looks like an empty
                # success. Require at least one recognised VBA source folder.
                $vbaSubdirs = @('Modules','ClassModules','Forms','Objects')
                $foundSub = @($vbaSubdirs | Where-Object { Test-Path (Join-Path $resolvedSrc $_) })
                if ($foundSub.Count -eq 0) {
                    throw ("No VBA source folders ($($vbaSubdirs -join '/')) " +
                           "found under '$resolvedSrc'. Check source_dir.")
                }

                $imported = @()
                $removed  = @()

                # Strip user components (anything not Document, type != 100).
                # A password-locked project is unlocked at session startup;
                # Get-AccessibleVBProject throws an actionable error if not.
                $proj = Get-AccessibleVBProject $Wb
                $toRemove = @()
                foreach ($c in $proj.VBComponents) {
                    if ($c.Type -ne 100) { $toRemove += $c.Name }
                }
                foreach ($name in $toRemove) {
                    $proj.VBComponents.Remove($proj.VBComponents.Item($name)) | Out-Null
                    $removed += $name
                }
                # Re-import .bas/.cls/.frm from source layout
                foreach ($sub in @('Modules','ClassModules','Forms','Objects')) {
                    $d = Join-Path $resolvedSrc $sub
                    if (-not (Test-Path $d)) { continue }
                    Get-ChildItem -LiteralPath $d -File | ForEach-Object {
                        if ($_.Extension -in '.bas','.cls','.frm') {
                            $imp = $proj.VBComponents.Import($_.FullName)
                            $imported += $imp.Name
                        }
                    }
                }
                Write-EventLine @{
                    t = 'sync_completed'
                    id = $cmd.id
                    imported = $imported
                    removed  = $removed
                }
            } catch {
                Write-EventLine @{ t = 'sync_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'read_range' {
            try {
                $sheetName = $cmd.sheet
                $addr      = $cmd.range
                $sheet = if ($sheetName) { $Wb.Sheets.Item($sheetName) } else { $Wb.ActiveSheet }
                $rng = $sheet.Range($addr)
                $useFormulas = [bool]$cmd.include_formulas
                # Iterate cell-by-cell using Range.Cells.Item(row,col) rather
                # than Range.Value()'s 2D array. PowerShell's COM marshalling
                # of variant 2D arrays into Object[,] is inconsistent (the
                # array sometimes arrives as a 1D PSObject array, leading to
                # mis-shaped output). Per-cell read is slower but always
                # produces a clean jagged array for JSON.
                $h = [int]$rng.Rows.Count
                $w = [int]$rng.Columns.Count
                $rows = New-Object 'object[][]' $h
                for ($i = 0; $i -lt $h; $i++) {
                    $rows[$i] = New-Object 'object[]' $w
                    for ($j = 0; $j -lt $w; $j++) {
                        $cell = $rng.Cells.Item($i + 1, $j + 1)
                        $rows[$i][$j] = if ($useFormulas) { $cell.Formula } else { $cell.Value2 }
                    }
                }
                Write-EventLine @{
                    t = 'range_read'
                    id = $cmd.id
                    sheet = $sheet.Name
                    range = $rng.Address($false, $false)
                    rows = $rows
                }
            } catch {
                Write-EventLine @{ t = 'range_read_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'write_range' {
            try {
                $sheetName = $cmd.sheet
                $addr      = $cmd.range
                $sheet = if ($sheetName) { $Wb.Sheets.Item($sheetName) } else { $Wb.ActiveSheet }
                $rng = $sheet.Range($addr)
                $rowsIn = @($cmd.values)
                $h = $rowsIn.Count
                $w = if ($h -gt 0) { @($rowsIn[0]).Count } else { 0 }
                if ($h -eq 0 -or $w -eq 0) { throw "write_range: values must be non-empty 2D array" }

                # Resize range if a single anchor cell was passed
                if ($rng.Rows.Count -eq 1 -and $rng.Columns.Count -eq 1) {
                    $rng = $rng.Resize($h, $w)
                }

                # Per-cell write via the sheet's absolute Cells(row, col)
                # property. Cell-by-cell is slower than batch Value2 but
                # reliable in PowerShell — batch COM marshalling of an
                # object[,] sometimes flattens into a 1D string array.
                # Per-cell write via Range.Offset($i, $j) from the anchor.
                # Other access patterns (Cells.Item, Range(<string>)) hit
                # PowerShell COM overload-resolution issues ("Cannot cast
                # Double to String"); Range.Offset is unambiguous.
                # PowerShell COM dispatch caches the chosen overload per
                # property based on the FIRST argument type. After three
                # string writes, a Double assignment to Value2 throws
                # "Cannot cast Double to String". InvokeMember bypasses
                # the cache and dispatches per-call.
                $anchor = $sheet.Range($addr)
                for ($i = 0; $i -lt $h; $i++) {
                    $rowArr = @($rowsIn[$i])
                    for ($j = 0; $j -lt $w; $j++) {
                        $val = $rowArr[$j]
                        if ($val -is [int64] -or $val -is [decimal]) { $val = [double]$val }
                        $cell = $anchor.Offset($i, $j)
                        [void]$cell.GetType().InvokeMember('Value2',
                            [System.Reflection.BindingFlags]::SetProperty,
                            $null, $cell, @($val))
                    }
                }

                Write-EventLine @{
                    t = 'range_written'
                    id = $cmd.id
                    sheet = $sheet.Name
                    range = $rng.Address($false, $false)
                    rows = $h
                    cols = $w
                }
            } catch {
                Write-EventLine @{ t = 'range_write_failed'; id = $cmd.id; error = $_.Exception.Message; trace = $_.ScriptStackTrace }
            }
        }
        'analyze_vba' {
            try {
                $op   = [string]$cmd.op
                $name = [string]$cmd.name
                if (-not $op) {
                    Write-EventLine @{ t = 'analyze_vba_failed'; id = $cmd.id; reason = 'missing_op' }
                    break
                }
                if ($op -ne 'definition_of' -and $op -ne 'read_module' -and $op -ne 'callers_of') {
                    Write-EventLine @{ t = 'analyze_vba_failed'; id = $cmd.id; reason = 'unknown_op'; op = $op }
                    break
                }
                if (-not $name) {
                    Write-EventLine @{ t = 'analyze_vba_failed'; id = $cmd.id; reason = 'missing_name'; op = $op }
                    break
                }
                $proj = $null
                try { $proj = Get-AccessibleVBProject $Wb }
                catch {
                    $msg = $_.Exception.Message
                    $reason = if ($msg -match 'password-locked') { 'vba_project_locked' } else { $msg }
                    Write-EventLine @{ t = 'analyze_vba_failed'; id = $cmd.id; reason = $reason }
                    break
                }
                switch ($op) {
                    'definition_of' { Invoke-VbaDefinitionOf -Proj $proj -Name $name -Cmd $cmd }
                    'read_module'   { Invoke-VbaReadModule   -Proj $proj -Name $name -Cmd $cmd }
                    'callers_of'    { Invoke-VbaCallersOf    -Proj $proj -Name $name -Cmd $cmd }
                }
            } catch {
                Write-EventLine @{ t = 'analyze_vba_failed'; id = $cmd.id; reason = $_.Exception.Message }
            }
        }
        'list_macros' {
            try {
                $macros = New-Object System.Collections.ArrayList
                $proj = Get-AccessibleVBProject $Wb
                foreach ($comp in $proj.VBComponents) {
                    if ($comp.CodeModule.CountOfLines -eq 0) { continue }
                    $cm = $comp.CodeModule
                    for ($ln = 1; $ln -le $cm.CountOfLines; $ln++) {
                        $line = $cm.Lines($ln, 1).Trim()
                        if ($line -match '^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\((.*?)\)') {
                            # Group 1 is optional — $Matches[1] is $null for
                            # a Sub/Function with no Public/Private/Friend
                            # modifier; coerce to string before .Trim().
                            $vis    = "$($Matches[1])".Trim()
                            $kind   = $Matches[2]
                            $name   = $Matches[3]
                            $argstr = $Matches[4].Trim()
                            $isPublic = ($vis -eq '' -or $vis -ieq 'Public')
                            [void]$macros.Add([pscustomobject]@{
                                module    = $comp.Name
                                kind      = $kind.ToLower()
                                name      = $name
                                args      = $argstr
                                public    = $isPublic
                                line      = $ln
                            })
                        }
                    }
                }
                Write-EventLine @{ t = 'macros_listed'; id = $cmd.id; macros = $macros }
            } catch {
                Write-EventLine @{ t = 'macros_list_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'list_sheets' {
            try {
                $sheets = New-Object System.Collections.ArrayList
                foreach ($sh in $Wb.Sheets) {
                    $tables = @()
                    try { foreach ($lo in $sh.ListObjects) { $tables += $lo.Name } } catch {}
                    $usedRange = ''
                    try { $usedRange = $sh.UsedRange.Address($false, $false) } catch {}
                    [void]$sheets.Add([pscustomobject]@{
                        name       = $sh.Name
                        index      = [int]$sh.Index
                        used_range = $usedRange
                        tables     = $tables
                        hidden     = ($sh.Visible -eq 0)  # xlSheetVisible=-1, xlHidden=0, xlVeryHidden=2
                        protected  = [bool]$sh.ProtectContents
                    })
                }
                Write-EventLine @{ t = 'sheets_listed'; id = $cmd.id; sheets = $sheets }
            } catch {
                Write-EventLine @{ t = 'sheets_list_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'get_workbook_info' {
            try {
                $tableCount = 0; $namedRangeCount = 0; $hasPivots = $false; $hasConnections = $false
                $vbaProt = 0; $sizeBytes = 0; $lastAuthor = ''; $lastSave = ''
                foreach ($sh in $Wb.Sheets) {
                    try { $tableCount += $sh.ListObjects.Count } catch {}
                    try { if ($sh.PivotTables.Count -gt 0) { $hasPivots = $true } } catch {}
                }
                try { $namedRangeCount = $Wb.Names.Count } catch {}
                try { $hasConnections = ($Wb.Connections.Count -gt 0) } catch {}
                try { $vbaProt = [int]$Wb.VBProject.Protection } catch {}
                try { $sizeBytes = (Get-Item -LiteralPath $Wb.FullName).Length } catch {}
                try { $lastAuthor = "$($Wb.BuiltinDocumentProperties.Item('Last Author').Value)" } catch {}
                try { $lastSave = "$($Wb.BuiltinDocumentProperties.Item('Last Save Time').Value)" } catch {}
                $info = [ordered]@{
                    name              = $Wb.Name
                    full_name         = $Wb.FullName
                    size_bytes        = $sizeBytes
                    sheet_count       = $Wb.Sheets.Count
                    table_count       = $tableCount
                    named_range_count = $namedRangeCount
                    has_vba           = [bool]$Wb.HasVBProject
                    vba_protection    = $vbaProt
                    has_pivots        = $hasPivots
                    has_connections   = $hasConnections
                    file_format       = [int]$Wb.FileFormat
                    last_author       = $lastAuthor
                    last_save_time    = $lastSave
                }
                Write-EventLine @{ t = 'workbook_info'; id = $cmd.id; info = $info }
            } catch {
                Write-EventLine @{ t = 'workbook_info_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'run_tests' {
            try {
                $filter = $cmd.filter
                # Discover Test_* Subs in any module
                $tests = New-Object System.Collections.ArrayList
                $proj = Get-AccessibleVBProject $Wb
                foreach ($comp in $proj.VBComponents) {
                    if ($comp.CodeModule.CountOfLines -eq 0) { continue }
                    $cm = $comp.CodeModule
                    for ($ln = 1; $ln -le $cm.CountOfLines; $ln++) {
                        $line = $cm.Lines($ln, 1).Trim()
                        if ($line -match '^\s*(Public\s+)?Sub\s+(Test_[A-Za-z0-9_]*)\s*\(') {
                            $tname = $Matches[2]
                            if ($filter -and $tname -notmatch $filter) { continue }
                            [void]$tests.Add(@{ module = $comp.Name; name = $tname })
                        }
                    }
                }

                $start = [DateTime]::UtcNow
                $passed = 0; $failed = 0; $errored = 0
                foreach ($t in $tests) {
                    $tStart = [DateTime]::UtcNow
                    $status = 'pass'
                    $errObj = $null
                    try {
                        $ref = "'$($Wb.Name)'!$($t.module).$($t.name)"
                        if ($script:Watcher) { $script:Watcher.State.MacroRunning = $true }
                        try { $null = $Xl.Run($ref) }
                        finally { if ($script:Watcher) { $script:Watcher.State.MacroRunning = $false } }
                    } catch {
                        $msg = $_.Exception.Message
                        $cur = $_.Exception
                        while ($cur.InnerException) { $cur = $cur.InnerException; if ($cur.Message) { $msg = $cur.Message } }
                        # Treat any Err.Raise from inside as a failure; we can't
                        # easily distinguish "asserted fail" from "unexpected error"
                        # via COM, so callers treat both as "test didn't pass".
                        $status = 'fail'
                        $errObj = @{ message = $msg; hresult = $_.Exception.HResult }
                    }
                    $duration = [int]([DateTime]::UtcNow - $tStart).TotalMilliseconds
                    $resultEv = @{
                        t        = 'test_result'
                        module   = $t.module
                        name     = $t.name
                        status   = $status
                        duration_ms = $duration
                    }
                    if ($errObj) { $resultEv.error = $errObj; $failed++ } else { $passed++ }
                    Write-EventLine $resultEv
                }
                Write-EventLine @{
                    t = 'tests_completed'
                    id = $cmd.id
                    total = $tests.Count
                    passed = $passed
                    failed = $failed
                    errored = $errored
                    duration_ms = [int]([DateTime]::UtcNow - $start).TotalMilliseconds
                }
            } catch {
                Write-EventLine @{ t = 'tests_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'save_workbook' {
            try {
                $Wb.Save()
                Write-EventLine @{ t = 'workbook_saved'; id = $cmd.id; path = $Wb.FullName }
            } catch {
                Write-EventLine @{ t = 'save_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'save_as' {
            try {
                $newPath = $cmd.path
                if (-not $newPath) { throw "save_as requires path" }
                $fileFormat = if ($cmd.file_format) { [int]$cmd.file_format } else { [int]$Wb.FileFormat }
                $Wb.SaveAs($newPath, $fileFormat)
                Write-EventLine @{ t = 'workbook_saved_as'; id = $cmd.id; path = $newPath; file_format = $fileFormat }
            } catch {
                Write-EventLine @{ t = 'save_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'calculate' {
            try {
                $start = [DateTime]::UtcNow
                $Xl.Calculate()
                Write-EventLine @{ t = 'calculated'; id = $cmd.id; duration_ms = [int]([DateTime]::UtcNow - $start).TotalMilliseconds }
            } catch {
                Write-EventLine @{ t = 'calculate_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'refresh_all' {
            try {
                $start = [DateTime]::UtcNow
                $Wb.RefreshAll()
                Write-EventLine @{ t = 'refreshed_all'; id = $cmd.id; duration_ms = [int]([DateTime]::UtcNow - $start).TotalMilliseconds }
            } catch {
                Write-EventLine @{ t = 'refresh_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'refresh_connection' {
            try {
                $name = $cmd.name
                if (-not $name) { throw "refresh_connection requires name" }
                $start = [DateTime]::UtcNow
                $conn = $Wb.Connections.Item($name)
                $conn.Refresh()
                Write-EventLine @{ t = 'connection_refreshed'; id = $cmd.id; name = $name; duration_ms = [int]([DateTime]::UtcNow - $start).TotalMilliseconds }
            } catch {
                Write-EventLine @{ t = 'connection_failed'; id = $cmd.id; name = $cmd.name; error = $_.Exception.Message }
            }
        }
        'create_workbook' {
            try {
                $newWb = $Xl.Workbooks.Add()
                $savedPath = $null
                if ($cmd.save_as) {
                    $fmt = if ($cmd.file_format) { [int]$cmd.file_format } else { 52 }  # default xlOpenXMLWorkbookMacroEnabled
                    $newWb.SaveAs($cmd.save_as, $fmt)
                    $savedPath = $cmd.save_as
                }
                Write-EventLine @{ t = 'workbook_created'; id = $cmd.id; name = $newWb.Name; saved_to = $savedPath }
            } catch {
                Write-EventLine @{ t = 'create_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'open_workbook' {
            try {
                $p = $cmd.path
                if (-not $p) { throw "open_workbook requires path" }
                $resolved = (Resolve-Path -LiteralPath $p).Path
                $newWb = $Xl.Workbooks.Open($resolved, [Type]::Missing, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
                Write-EventLine @{ t = 'workbook_opened'; id = $cmd.id; name = $newWb.Name; path = $resolved }
            } catch {
                Write-EventLine @{ t = 'open_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'close_workbook' {
            try {
                $name = $cmd.name
                if (-not $name) { throw "close_workbook requires name (workbook name)" }
                $target = $Xl.Workbooks.Item($name)
                $target.Close([bool]$cmd.save)
                Write-EventLine @{ t = 'workbook_closed_cmd'; id = $cmd.id; name = $name }
            } catch {
                Write-EventLine @{ t = 'close_workbook_failed'; id = $cmd.id; error = $_.Exception.Message }
            }
        }
        'import' { Invoke-VbaSyncOp -Xl $Xl -Wb $Wb -cmd $cmd }
        'export' { Invoke-VbaSyncOp -Xl $Xl -Wb $Wb -cmd $cmd }
        'session_status' {
            # Single-call health snapshot. Consolidates state.json + watcher
            # state + a fresh COM-heartbeat liveness probe so an agent that
            # "feels something is off" gets one coherent view in one round
            # trip rather than tailing events.jsonl + reading state.json +
            # running Get-Process.
            #
            # Re-probe liveness fresh — Test-ExcelAlive emits excel_crashed
            # as a side effect if the probe flips, and snaps LastCrashedEvent.
            # The report below then reflects the just-updated state.
            [void](Test-ExcelAlive -Xl $Xl -ComHeartbeat)

            $strangers  = @()
            $armedCount = 0
            if ($script:Watcher) {
                try { $strangers  = @($script:Watcher.State.ReportedStrangerPids) } catch {}
                try { $armedCount = [int]$script:Watcher.State.ArmedRulesCount }   catch {}
            }

            $eventsSincePrev = 0
            if ($script:LastCommandSnapshot -and $null -ne $script:LastCommandSnapshot.events_since_prev) {
                $eventsSincePrev = [int]$script:LastCommandSnapshot.events_since_prev
            }

            $report = [ordered]@{
                t                          = 'session_status_report'
                id                         = $cmd.id
                report_version             = 1
                status                     = $(if ($script:ExcelDead) { 'crashed' } else { 'ready' })
                host_alive                 = $true
                excel_alive                = -not $script:ExcelDead
                host_pid                   = $PID
                excel_pid                  = $script:ExcelPid
                workbook                   = $resolvedWb
                session_id                 = $SessionId
                started_at                 = $script:StartedAt
                last_command               = $script:LastCommandSnapshot
                last_event                 = $script:LastEventSnapshot
                events_since_last_command  = $eventsSincePrev
                stranger_excel_pids        = $strangers
                armed_responses_active     = $armedCount
                idle_seconds               = [int]([DateTime]::UtcNow - $script:LastCommandTime).TotalSeconds
                last_excel_check_ts        = $script:LastExcelCheckTs
            }
            if ($script:ExcelDead -and $script:LastCrashedEvent) {
                $report['crashed_event'] = $script:LastCrashedEvent
            }
            Write-EventLine $report
        }
        'close_recovery' {
            # Terminate a stranger EXCEL.EXE the watcher surfaced via
            # `recovery_instance_detected`. Validation gate: pid must be in
            # the watcher's ReportedStrangerPids — refuses to kill arbitrary
            # processes (only Excels we already told the agent about).
            $targetPid = $null
            try { $targetPid = [int]$cmd.pid } catch {}
            if (-not $targetPid) {
                Write-EventLine @{ t = 'recovery_close_failed'; id = $cmd.id; pid = $cmd.pid; reason = 'missing_or_invalid_pid' }
                break
            }
            $reported = $false
            if ($script:Watcher -and $script:Watcher.State.ReportedStrangerPids) {
                $reported = $script:Watcher.State.ReportedStrangerPids.Contains($targetPid)
            }
            if (-not $reported) {
                Write-EventLine @{ t = 'recovery_close_failed'; id = $cmd.id; pid = $targetPid; reason = 'unknown_or_unreported_pid' }
                break
            }
            try {
                Stop-Process -Id $targetPid -Force -ErrorAction Stop
                Write-EventLine @{ t = 'recovery_closed'; id = $cmd.id; pid = $targetPid; ok = $true }
            } catch {
                Write-EventLine @{ t = 'recovery_close_failed'; id = $cmd.id; pid = $targetPid; reason = $_.Exception.Message }
            }
        }
        'close' {
            Write-EventLine @{ t = 'closing'; id = $cmd.id }
            $script:Stop = $true
        }
        default {
            Write-EventLine @{ t = 'unknown_command'; id = $cmd.id; cmd = $cmd.cmd }
        }
    }
}

# ---------- main ----------

# Snapshot existing EXCEL.EXE PIDs so we can identify the one we create
# even when $xl.Hwnd is unavailable (a hidden COM Excel can report Hwnd 0).
$excelPidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = [bool]$Visible
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
# Low = let macros run (default ForceDisable would block xl.Run)
$xl.AutomationSecurity = 1
$wb = $null

try {
    # IgnoreReadOnlyRecommended=True so workbooks with that flag don't open RO.
    $wb = $xl.Workbooks.Open($resolvedWb, [Type]::Missing, $false, [Type]::Missing,
                             [Type]::Missing, [Type]::Missing, $true)

    # Determine Excel's PID (not our $PID) so the watcher targets the right process.
    # Recorded in state.json so tests / cleanup scripts can identify which
    # EXCEL.EXE belongs to this session — never blanket-kill all Excel.
    # A hidden COM Excel can report Hwnd = 0/null, so [IntPtr]$xl.Hwnd would
    # throw "Cannot convert null to type System.IntPtr" — guard it and fall
    # back to diffing the EXCEL.EXE process list captured before creation.
    #
    # MUST happen before the started event + Save-State below, so that any
    # consumer reading state.json on `started` sees a populated excel_pid.
    $excelPid = [uint32]0
    $xlHwnd = $null
    try { $xlHwnd = $xl.Hwnd } catch {}
    if ($xlHwnd) {
        [XcSession.Win32]::GetWindowThreadProcessId([IntPtr][int]$xlHwnd, [ref]$excelPid) | Out-Null
    }
    if (-not $excelPid) {
        $excelPidNew = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
                         Where-Object { $excelPidsBefore -notcontains $_.Id } |
                         Select-Object -ExpandProperty Id)
        if ($excelPidNew.Count -ge 1) { $excelPid = [uint32]$excelPidNew[0] }
    }
    $script:ExcelPid = [int]$excelPid

    Save-State 'ready'
    Write-EventLine @{ t = 'started'; pid = $PID; excel_pid = $script:ExcelPid; workbook = $resolvedWb; session_id = $SessionId; visible = [bool]$Visible }
    Write-Host "Session $SessionId started (pid=$PID, excel_pid=$($script:ExcelPid), workbook=$resolvedWb, visible=$([bool]$Visible))" -ForegroundColor Cyan

    # Bug C: a password-locked VBA project is inaccessible over COM
    # (VBComponents is null) until unlocked, and there is no object-model
    # method to enter the password. Resolve the password (param > env >
    # gitignored file) and, if the project is locked, unlock it once here —
    # before the dialog watcher starts, so the watcher can't race the
    # VBAProject Password dialog. The unlock persists for the COM session.
    $vbaPw = $VbaPassword
    if (-not $vbaPw) { $vbaPw = $env:XC_VBA_PASSWORD }
    if (-not $vbaPw) {
        $pwFile = Join-Path $PSScriptRoot '.vba-password'
        if (Test-Path -LiteralPath $pwFile) {
            $vbaPw = (Get-Content -LiteralPath $pwFile -Raw).Trim()
        }
    }
    $projLocked = $false
    try { $projLocked = ([int]$wb.VBProject.Protection -ne 0) } catch {}
    if ($projLocked) {
        if ($vbaPw) {
            $unlocked = $false
            try {
                $unlocked = Unlock-VbaProject -Xl $xl -Password $vbaPw -ExcelPid $excelPid
            } catch {
                Write-EventLine @{ t = 'vba_unlock_failed'; error = $_.Exception.Message }
            }
            if ($unlocked) {
                Write-EventLine @{ t = 'vba_unlocked' }
                Write-Host "VBA project unlocked." -ForegroundColor Cyan
            } elseif (-not $unlocked) {
                Write-EventLine @{ t = 'vba_unlock_failed'; error = 'Unlock-VbaProject returned false — wrong password, or the password dialog could not be driven.' }
            }
        } else {
            Write-EventLine @{ t = 'vba_unlock_failed'; error = 'VBA project is password-locked but no password was supplied (-VbaPassword / $env:XC_VBA_PASSWORD / a .vba-password file beside the harness scripts). sync_vba / compile_check / run_tests / list_macros will fail.' }
        }
    }

    $watcher = Start-SessionDialogWatcher `
        -ProcessId         $excelPid `
        -EventsFile        $eventsFile `
        -CommandsFile      $commandsFile `
        -CapturesDir       $capturesDir `
        -Visible           ([bool]$Visible) `
        -BaselineExcelPids $excelPidsBefore
    $script:Watcher     = $watcher
    $script:CapturesDir = $capturesDir
    $script:SessionDir  = $sessionDir

    while (-not $script:Stop) {
        # Per-loop process-liveness check (item A). On exit it flips
        # ExcelDead and emits excel_crashed exactly once.
        if (-not $script:ExcelDead) { [void](Test-ExcelAlive -Xl $xl) }

        $cmds = Read-NewCommands
        foreach ($entry in $cmds) {
            $cmd = $entry.obj
            # respond_dialog / set_form_control / arm_response /
            # arm_form_control are owned by the watcher runspace (they act
            # on a live modal while the main loop is blocked in $Xl.Run) —
            # skip them here.
            if ($cmd.cmd -in 'respond_dialog','set_form_control','arm_response','arm_form_control') { continue }
            $script:LastCommandTime = [DateTime]::UtcNow

            # Pre-dispatch COM heartbeat (item C). Catches a race where the
            # process survives the per-loop check but the COM ref is gone.
            if (-not $script:ExcelDead) { [void](Test-ExcelAlive -Xl $xl -ComHeartbeat) }

            # In `crashed`, only close / close_recovery are accepted.
            # Anything else emits command_rejected (single new event shape,
            # not per-<cmd>_failed variants — agreed on the collab channel).
            if ($script:ExcelDead -and $cmd.cmd -notin 'close','close_recovery') {
                Write-EventLine @{ t = 'command_rejected'; id = $cmd.id; cmd = $cmd.cmd; reason = 'excel_crashed' }
                continue
            }

            # Snap the events-since counter BEFORE command_ack zeroes it,
            # then build LastCommandSnapshot so session_status can report
            # "events between the previous command and this one".
            $snappedEventsSincePrev = $script:EventsSinceLastCommand
            $ackAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
            $script:LastCommandSnapshot = @{
                id = $cmd.id
                cmd = $cmd.cmd
                at = $ackAt
                events_since_prev = $snappedEventsSincePrev
            }
            Write-EventLine @{ t = 'command_ack'; id = $cmd.id; cmd = $cmd.cmd; ts = $ackAt }
            Save-State 'busy'
            Invoke-SessionCommand -Xl $xl -Wb $wb -cmd $cmd
            if ($script:ExcelDead) { Save-State 'crashed' } else { Save-State 'ready' }
            if ($script:Stop) { break }
        }
        if (-not $script:ExcelDead) { Save-State 'ready' }

        # Zombie-host idle timeout in `crashed` (10 min). Prevents a dead
        # session lingering forever if no agent reads the channel.
        if ($script:ExcelDead) {
            $idleMs = ([DateTime]::UtcNow - $script:LastCommandTime).TotalMilliseconds
            if ($idleMs -gt $script:CrashedIdleTimeoutMs) {
                Write-EventLine @{ t = 'crashed_idle_timeout'; idle_ms = [int]$idleMs }
                $script:Stop = $true
            }
        }

        if (-not $script:Stop) { Start-Sleep -Milliseconds $PollMs }
    }
}
catch {
    Write-EventLine @{ t = 'session_error'; error = $_.Exception.Message; stack = $_.ScriptStackTrace }
    Save-State 'crashed'
    Write-Host "Session $SessionId crashed: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($watcher) { try { Stop-SessionDialogWatcher $watcher } catch {} }
    try { if ($wb) { $wb.Close($false) } } catch {}
    try { $xl.Quit() } catch {}
    # Release every COM ref now; leaving any (the workbook was) to the GC
    # finalizer makes EXCEL.EXE teardown stall when another Excel is open.
    foreach ($o in @($wb, $xl)) {
        if ($o) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null } catch {} }
    }
    $wb = $null; $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Save-State 'closed'
    Write-EventLine @{ t = 'closed' }
    Write-Host "Session $SessionId closed" -ForegroundColor Cyan
}
