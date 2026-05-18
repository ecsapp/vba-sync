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
. (Join-Path $PSScriptRoot 'capture.ps1')

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
                $macroArgs = @()
                if ($null -ne $cmd.args) { $macroArgs = @($cmd.args) }
                $macroRef = "'$($Wb.Name)'!$($cmd.name)"
                if ($macroArgs.Count -eq 0) {
                    $result = $Xl.Run($macroRef)
                } else {
                    $result = Invoke-Macro -Xl $Xl -MacroName $macroRef -Args $macroArgs
                }
                Write-EventLine @{
                    t = 'macro_completed'
                    id = $cmd.id
                    name = $cmd.name
                    result = $result
                    duration_ms = [int]([DateTime]::UtcNow - $start).TotalMilliseconds
                }
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
            }
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
                        Capture-Window -Hwnd $hwnd -Path $outPath | Out-Null
                    }
                    '^worksheet:(.+)$' {
                        $sheetName = $Matches[1]
                        $sheet = $Wb.Sheets.Item($sheetName)
                        $sheet.Activate()
                        Start-Sleep -Milliseconds 200
                        $hwnd = [int64]$Xl.Hwnd
                        Capture-Window -Hwnd $hwnd -Path $outPath | Out-Null
                    }
                    '^(dialog|form):(.+)$' {
                        $dlgId = $Matches[2]
                        $info  = $script:Watcher.State.DialogInfo[$dlgId]
                        if (-not $info) { throw "Unknown dialog/form id: $dlgId" }
                        Capture-Window -Hwnd $info.Hwnd -Path $outPath | Out-Null
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

                # Optional unlock — defer the import here, no .xlam dependency.
                if ($cmd.vba_password) {
                    throw "VBA password unlock is not yet implemented in sync_vba (deferred to a later phase). Unlock manually for now."
                }

                $imported = @()
                $removed  = @()

                # Strip user components (anything not Document, type != 100)
                $proj = $Wb.VBProject
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
                $values = if ($useFormulas) { $rng.Formula } else { $rng.Value() }
                # Range.Value returns scalar for single cell, 2D array otherwise.
                # Build a jagged object[][] for JSON-friendly nesting.
                # Use Rank/GetUpperBound rather than `-is [object[,]]` — under
                # PowerShell COM marshalling the runtime type isn't always
                # [object[,]] even when the value is a 2D array.
                $rows = $null
                $rank = 0
                try { $rank = $values.Rank } catch { $rank = 0 }
                if ($rank -eq 2) {
                    $r1 = $values.GetLowerBound(0); $r2 = $values.GetUpperBound(0)
                    $c1 = $values.GetLowerBound(1); $c2 = $values.GetUpperBound(1)
                    $h = $r2 - $r1 + 1
                    $w = $c2 - $c1 + 1
                    $rows = New-Object 'object[][]' $h
                    for ($i = 0; $i -lt $h; $i++) {
                        $rows[$i] = New-Object 'object[]' $w
                        for ($j = 0; $j -lt $w; $j++) {
                            $rows[$i][$j] = $values[$r1 + $i, $c1 + $j]
                        }
                    }
                } else {
                    $rows = New-Object 'object[][]' 1
                    $rows[0] = ,$values
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
                # Build the absolute address per cell. Cells.Item(row, col)
                # ran into PowerShell COM overload-resolution issues
                # ("cannot cast Double to String"); going through
                # Range("A1") is unambiguous.
                $startRow = [int]$rng.Row
                $startCol = [int]$rng.Column
                for ($i = 0; $i -lt $h; $i++) {
                    $rowArr = @($rowsIn[$i])
                    for ($j = 0; $j -lt $w; $j++) {
                        $val = $rowArr[$j]
                        if ($val -is [int64] -or $val -is [decimal]) { $val = [double]$val }
                        $colLetter = ''
                        $cn = $startCol + $j
                        while ($cn -gt 0) {
                            $rem = (($cn - 1) % 26)
                            $colLetter = [char]([byte][char]'A' + $rem) + $colLetter
                            $cn = [int](($cn - $rem - 1) / 26)
                        }
                        $cellAddr = "$colLetter$($startRow + $i)"
                        $sheet.Range($cellAddr).Value2 = $val
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
        -CommandsFile $commandsFile `
        -CapturesDir  $capturesDir
    $script:Watcher     = $watcher
    $script:CapturesDir = $capturesDir
    $script:SessionDir  = $sessionDir

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
