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
    [string]$SessionsRoot = (Join-Path $PSScriptRoot 'sessions'),
    [int]$PollMs = 250
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'session-dialog-watcher.ps1')

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

$script:LastOffset = 0
$script:Stop       = $false
$script:StartedAt  = (Get-Date).ToUniversalTime().ToString('o')

# Resume offset if state.json says so (allows restart-after-crash to skip
# already-consumed commands)
if (Test-Path $stateFile) {
    try {
        $prev = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json
        if ($prev.last_command_offset) { $script:LastOffset = [int64]$prev.last_command_offset }
    } catch {}
}

function Write-EventLine([hashtable]$Event) {
    $line = ConvertTo-Json -Compress -Depth 10 -InputObject $Event
    Add-Content -LiteralPath $eventsFile -Value $line -Encoding UTF8
}

function Save-State([string]$Status) {
    $s = [ordered]@{
        pid                 = $PID
        workbook            = $Workbook
        session_id          = $SessionId
        status              = $Status
        started_at          = $script:StartedAt
        last_command_offset = $script:LastOffset
    }
    Set-Content -LiteralPath $stateFile -Value (ConvertTo-Json -Depth 5 -InputObject $s) -Encoding UTF8
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
                    Write-EventLine @{ t = 'command_error'; error = "Invalid JSON: $($_.Exception.Message)"; raw = $trim }
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
function Invoke-Macro($Xl, [string]$MacroName, [object[]]$Args) {
    $all = ,$MacroName + ($Args | ForEach-Object { $_ })
    return $Xl.GetType().InvokeMember(
        'Run',
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null, $Xl, $all)
}

function Handle-Command($Xl, $Wb, $cmd) {
    switch ($cmd.cmd) {
        'respond_dialog' {
            # Owned by the dialog watcher (separate runspace). Main loop
            # already skipped these — this branch is here as a safety net
            # in case the dispatch routing changes.
            return
        }
        'run_macro' {
            $start = [DateTime]::UtcNow
            try {
                $args = @()
                if ($null -ne $cmd.args) { $args = @($cmd.args) }
                # Qualify with workbook name to avoid "macro not available"
                # when multiple workbooks are open or modules don't resolve
                # globally. Excel.Application.Run accepts "'wbname'!Macro".
                $macroRef = "'$($Wb.Name)'!$($cmd.name)"
                $result = Invoke-Macro -Xl $Xl -MacroName $macroRef -Args $args
                $duration = [int]([DateTime]::UtcNow - $start).TotalMilliseconds
                Write-EventLine @{
                    t = 'macro_completed'
                    id = $cmd.id
                    name = $cmd.name
                    result = $result
                    duration_ms = $duration
                }
            } catch {
                Write-EventLine @{
                    t = 'macro_failed'
                    id = $cmd.id
                    name = $cmd.name
                    error = $_.Exception.Message
                    error_type = $_.Exception.GetType().FullName
                }
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

$resolvedWb = (Resolve-Path -LiteralPath $Workbook).Path
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
# Low = let macros run (default ForceDisable would block xl.Run)
$xl.AutomationSecurity = 1
$wb = $null

try {
    # IgnoreReadOnlyRecommended=True so workbooks with that flag don't open RO.
    $wb = $xl.Workbooks.Open($resolvedWb, [Type]::Missing, $false, [Type]::Missing,
                             [Type]::Missing, [Type]::Missing, $true)

    Save-State 'ready'
    Write-EventLine @{ t = 'started'; pid = $PID; workbook = $resolvedWb; session_id = $SessionId }
    Write-Host "Session $SessionId started (pid=$PID, workbook=$resolvedWb)" -ForegroundColor Cyan

    # Determine Excel's PID (not our $PID) so the watcher targets the right process
    $excelPid = [uint32]0
    [XcSession.Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid) | Out-Null

    $watcher = Start-SessionDialogWatcher `
        -ProcessId    $excelPid `
        -EventsFile   $eventsFile `
        -CommandsFile $commandsFile

    while (-not $script:Stop) {
        $cmds = Read-NewCommands
        foreach ($entry in $cmds) {
            $cmd = $entry.obj
            # respond_dialog is owned by the watcher runspace — skip in main
            if ($cmd.cmd -eq 'respond_dialog') { continue }
            Write-EventLine @{ t = 'command_ack'; id = $cmd.id; cmd = $cmd.cmd }
            Save-State 'busy'
            Handle-Command -Xl $xl -Wb $wb -cmd $cmd
            Save-State 'ready'
            if ($script:Stop) { break }
        }
        Save-State 'ready'
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
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Save-State 'closed'
    Write-EventLine @{ t = 'closed' }
    Write-Host "Session $SessionId closed" -ForegroundColor Cyan
}
