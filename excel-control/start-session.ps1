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
    [int]$PollMs = 250,
    # -Visible: show Excel on the desktop so you can watch the agent
    # work. Worksheet screenshots become meaningful (headless mode
    # renders mostly-empty main window). Side-effects: any UI element
    # the agent triggers steals foreground focus from other windows.
    [switch]$Visible,
    # -VbaPassword: password for a locked VBA project. Without it the
    # harness cannot sync/compile/test a password-protected VBA project.
    # Precedence: this param > $env:XC_VBA_PASSWORD > a gitignored
    # <tools>/.vba-password file. Never commit the secret.
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

function Write-EventLine([hashtable]$EventObj) {
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
        pid                 = $PID
        excel_pid           = $script:ExcelPid
        workbook            = $Workbook
        session_id          = $SessionId
        status              = $Status
        visible             = [bool]$Visible
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
              "via -VbaPassword / `$env:XC_VBA_PASSWORD / tools/.vba-password " +
              "so the session auto-unlocks the project at startup."
    }
    return $proj
}

function Invoke-SessionCommand($Xl, $Wb, $cmd) {
    switch ($cmd.cmd) {
        'respond_dialog' {
            # Owned by the dialog watcher (separate runspace). Main loop
            # already skipped these — this branch is here as a safety net
            # in case the dispatch routing changes.
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
                $result = switch ($macroArgs.Count) {
                    0 { $Xl.Run($macroRef) }
                    1 { $Xl.Run($macroRef, $macroArgs[0]) }
                    2 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1]) }
                    3 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1], $macroArgs[2]) }
                    4 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1], $macroArgs[2], $macroArgs[3]) }
                    5 { $Xl.Run($macroRef, $macroArgs[0], $macroArgs[1], $macroArgs[2], $macroArgs[3], $macroArgs[4]) }
                    default { Invoke-Macro -Xl $Xl -MacroName $macroRef -MacroArgsArr $macroArgs }
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
                        $null = $Xl.Run($ref)
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

    Save-State 'ready'
    Write-EventLine @{ t = 'started'; pid = $PID; workbook = $resolvedWb; session_id = $SessionId; visible = [bool]$Visible }
    Write-Host "Session $SessionId started (pid=$PID, workbook=$resolvedWb, visible=$([bool]$Visible))" -ForegroundColor Cyan

    # Determine Excel's PID (not our $PID) so the watcher targets the right process.
    # Recorded in state.json so tests / cleanup scripts can identify which
    # EXCEL.EXE belongs to this session — never blanket-kill all Excel.
    # A hidden COM Excel can report Hwnd = 0/null, so [IntPtr]$xl.Hwnd would
    # throw "Cannot convert null to type System.IntPtr" — guard it and fall
    # back to diffing the EXCEL.EXE process list captured before creation.
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
    Save-State 'ready'  # re-save now that we know excel_pid

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
            Write-EventLine @{ t = 'vba_unlock_failed'; error = 'VBA project is password-locked but no password was supplied (-VbaPassword / $env:XC_VBA_PASSWORD / tools/.vba-password). sync_vba / compile_check / run_tests / list_macros will fail.' }
        }
    }

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
            Invoke-SessionCommand -Xl $xl -Wb $wb -cmd $cmd
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
