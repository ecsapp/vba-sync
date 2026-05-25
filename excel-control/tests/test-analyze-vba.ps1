# analyze_vba regression — three sub-tests in one session for the v1
# ops: definition_of, read_module, callers_of. Syncs a known small
# fixture so we can assert exact module/line/name/caller results.

[CmdletBinding()]
param(
    [string]$SessionId = "test-vba-$(Get-Random -Maximum 99999)"
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

Write-Host "Test: analyze_vba (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $workbook `
    -SessionId $SessionId -SessionsRoot $sessions

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

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no 'started' event" }
    Write-Host "  started" -ForegroundColor Green

    # Seed fixture with three small modules. Layout designed to exercise:
    #   - definition_of: BuildSql defined in modSql; case-insensitive lookup
    #     ("buildsql" should return canonical "BuildSql"); also a same-named
    #     Sub in a second module (modOther.BuildSql) to assert matches[] is
    #     an array.
    #   - callers_of: BuildSql is called from RunReport in modCore (line 8)
    #     and twice from RunBatch in modCore (lines 13 and 14). NOT called
    #     from modSql.BuildSql itself (that's the def line).
    #   - read_module: source must contain the Attribute VB_Name header and
    #     the known body.
    $srcDir = Join-Path $sessionDir '_src'
    New-Item -ItemType Directory -Force -Path (Join-Path $srcDir 'Modules') | Out-Null

    $modCore = @'
Attribute VB_Name = "modCore"
Option Explicit

Public Sub RunReport()
    Dim s As String
    ' below-line call site (will appear at line 8 inside RunReport's body)
    s = BuildSql("a", "b")
End Sub

Public Sub RunBatch()
    Dim s As String
    s = BuildSql("x", "y")
    s = BuildSql("p", "q")
End Sub
'@
    $modSql = @'
Attribute VB_Name = "modSql"
Option Explicit

Public Function BuildSql(k As String, v As String) As String
    BuildSql = "SET " & k & "=" & v
End Function
'@
    $modOther = @'
Attribute VB_Name = "modOther"
Option Explicit

Private Sub BuildSql()
    ' same name, different module, no args — definition_of should
    ' return both with canonical casing
End Sub
'@
    # Regression for the multi-line-signature false-negative (HARNESS_IMPROVEMENTS
    # P1.5). VBA's `_` line-continuation can split a Sub/Function signature
    # across physical lines. The Get-VbaProcedures scanner must join those
    # before regex-matching, otherwise the procedure is invisible to
    # definition_of and callers_of.
    $modMulti = @'
Attribute VB_Name = "modMulti"
Option Explicit

Private Function WideSig(a As String, b As String, _
                         c As String, d As String, _
                         e As String) As String
    WideSig = a & b & c & d & e
End Function

Public Sub TinyOne()
End Sub
'@
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modCore.bas')  -Value $modCore  -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modSql.bas')   -Value $modSql   -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modOther.bas') -Value $modOther -Encoding ASCII
    Set-Content -LiteralPath (Join-Path $srcDir 'Modules\modMulti.bas') -Value $modMulti -Encoding ASCII

    Add-Content -LiteralPath $commandsFile -Value (@{ id='c0'; cmd='sync_vba'; source_dir=$srcDir } | ConvertTo-Json -Compress) -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'sync_completed' 15 -IdMatch 'c0')) { throw "no sync_completed" }
    Write-Host "  sync_completed (3 modules)" -ForegroundColor Green

    # ---------- (1) definition_of ----------
    # Case-insensitive lookup: "buildsql" should hit both modSql.BuildSql
    # (Function) and modOther.BuildSql (Sub), with canonical casing.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c1","cmd":"analyze_vba","op":"definition_of","name":"buildsql"}' -Encoding UTF8
    $r1 = Wait-ForEvent $eventsFile 'vba_analysis_result' 10 -IdMatch 'c1'
    if (-not $r1) { throw "no vba_analysis_result for c1" }
    if ($r1.op -ne 'definition_of') { throw "op expected definition_of, got '$($r1.op)'" }
    $m1 = @($r1.matches)
    if ($m1.Count -ne 2) { throw "expected 2 matches for buildsql, got $($m1.Count)" }
    foreach ($m in $m1) {
        if ($m.name -ne 'BuildSql') { throw "canonical casing not preserved: got '$($m.name)' (expected 'BuildSql')" }
    }
    $byMod = @{}
    foreach ($m in $m1) { $byMod[$m.module] = $m }
    if (-not $byMod.ContainsKey('modSql'))   { throw "modSql.BuildSql not found in matches" }
    if (-not $byMod.ContainsKey('modOther')) { throw "modOther.BuildSql not found in matches" }
    if ($byMod['modSql'].kind   -ne 'function') { throw "modSql.BuildSql kind expected 'function', got '$($byMod['modSql'].kind)'" }
    if ($byMod['modOther'].kind -ne 'sub')      { throw "modOther.BuildSql kind expected 'sub', got '$($byMod['modOther'].kind)'" }
    if (-not $byMod['modSql'].public)           { throw "modSql.BuildSql public expected true" }
    if ($byMod['modOther'].public)              { throw "modOther.BuildSql public expected false (Private)" }
    Write-Host "  (1) definition_of: 2 matches, canonical casing 'BuildSql', kinds + visibility correct" -ForegroundColor Green

    # ---------- (2) callers_of ----------
    # RunReport calls BuildSql once; RunBatch calls it twice. modSql.BuildSql
    # itself should NOT appear as a caller (the def line is excluded).
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c2","cmd":"analyze_vba","op":"callers_of","name":"BuildSql"}' -Encoding UTF8
    $r2 = Wait-ForEvent $eventsFile 'vba_analysis_result' 10 -IdMatch 'c2'
    if (-not $r2) { throw "no vba_analysis_result for c2" }
    $callers = @($r2.callers)
    if ($callers.Count -ne 2) { throw "expected 2 callers for BuildSql, got $($callers.Count); list: $($callers | ForEach-Object { $_.module + '.' + $_.name } | Join-String -Separator ', ')" }
    $callByName = @{}
    foreach ($c in $callers) { $callByName[$c.name] = $c }
    if (-not $callByName.ContainsKey('RunReport')) { throw "RunReport not found in callers" }
    if (-not $callByName.ContainsKey('RunBatch'))  { throw "RunBatch not found in callers" }
    # RunReport calls it once; RunBatch calls it twice.
    if (@($callByName['RunReport'].call_lines).Count -ne 1) { throw "RunReport call_lines count expected 1, got $(@($callByName['RunReport'].call_lines).Count)" }
    if (@($callByName['RunBatch'].call_lines).Count  -ne 2) { throw "RunBatch call_lines count expected 2, got $(@($callByName['RunBatch'].call_lines).Count)" }
    # The defining Sub (modSql.BuildSql) must NOT appear as a self-caller.
    if ($callByName.ContainsKey('BuildSql')) { throw "modSql.BuildSql should NOT appear as a caller of itself" }
    Write-Host "  (2) callers_of: RunReport (1 call) + RunBatch (2 calls); self-call excluded" -ForegroundColor Green

    # ---------- (3) read_module ----------
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"analyze_vba","op":"read_module","name":"modSql"}' -Encoding UTF8
    $r3 = Wait-ForEvent $eventsFile 'vba_analysis_result' 10 -IdMatch 'c3'
    if (-not $r3) { throw "no vba_analysis_result for c3" }
    if ($r3.module.name -ne 'modSql')   { throw "read_module name expected 'modSql', got '$($r3.module.name)'" }
    if ($r3.module.kind -ne 'standard') { throw "read_module kind expected 'standard', got '$($r3.module.kind)'" }
    if ([int]$r3.module.line_count -lt 3) { throw "line_count expected >=3, got $($r3.module.line_count)" }
    $src = [string]$r3.module.source
    # VBE's CodeModule.Lines returns code-body only (no Attribute VB_Name
    # header — that's metadata, not a code line). The body must contain
    # Option Explicit and the function signature.
    if ($src -notmatch 'Option Explicit')              { throw "source missing 'Option Explicit'" }
    if ($src -notmatch 'Public Function BuildSql')     { throw "source missing BuildSql signature" }
    if ($src -notmatch 'BuildSql = "SET "')            { throw "source missing function body" }
    Write-Host "  (3) read_module: kind=standard line_count=$($r3.module.line_count); body present (no Attribute header per VBE)" -ForegroundColor Green

    # ---------- (4) failure modes — quick sanity ----------
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c4","cmd":"analyze_vba","op":"definition_of","name":"DoesNotExist"}' -Encoding UTF8
    $r4 = Wait-ForEvent $eventsFile 'vba_analysis_result' 10 -IdMatch 'c4'
    if (-not $r4) { throw "no vba_analysis_result for c4 (not_found)" }
    if (@($r4.matches).Count -ne 0) { throw "not_found should return empty matches[], got $(@($r4.matches).Count)" }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c5","cmd":"analyze_vba","op":"frobnicate","name":"x"}' -Encoding UTF8
    $r5 = Wait-ForEvent $eventsFile 'analyze_vba_failed' 10 -IdMatch 'c5'
    if (-not $r5) { throw "no analyze_vba_failed for c5 (unknown_op)" }
    if ($r5.reason -ne 'unknown_op') { throw "unknown_op reason expected 'unknown_op', got '$($r5.reason)'" }

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c6","cmd":"analyze_vba","op":"definition_of"}' -Encoding UTF8
    $r6 = Wait-ForEvent $eventsFile 'analyze_vba_failed' 10 -IdMatch 'c6'
    if (-not $r6) { throw "no analyze_vba_failed for c6 (missing_name)" }
    if ($r6.reason -ne 'missing_name') { throw "missing_name reason expected 'missing_name', got '$($r6.reason)'" }
    Write-Host "  (4) failure modes: not_found returns empty matches[]; unknown_op + missing_name surface correctly" -ForegroundColor Green

    # ---------- (5) multi-line signature regression (P1.5) ----------
    # WideSig in modMulti has a 3-physical-line signature joined by VBA's
    # `_` line-continuation. Pre-P1.5 the line-by-line regex never matched
    # the declaration line (no closing paren on it) and the procedure was
    # invisible to definition_of. Post-fix: the scanner joins continuation
    # lines before matching, so WideSig surfaces with Line pointing at the
    # declaration line and Signature carrying the joined logical signature.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c7","cmd":"analyze_vba","op":"definition_of","name":"WideSig"}' -Encoding UTF8
    $r7 = Wait-ForEvent $eventsFile 'vba_analysis_result' 10 -IdMatch 'c7'
    if (-not $r7) { throw "no vba_analysis_result for c7 (multi-line WideSig)" }
    $m7 = @($r7.matches)
    if ($m7.Count -ne 1) { throw "expected 1 match for WideSig, got $($m7.Count) — multi-line signature regression?" }
    if ($m7[0].module -ne 'modMulti') { throw "WideSig module expected 'modMulti', got '$($m7[0].module)'" }
    if ($m7[0].kind   -ne 'function') { throw "WideSig kind expected 'function', got '$($m7[0].kind)'" }
    if ($m7[0].public)                { throw "WideSig public expected false (Private)" }
    # Joined signature should contain all 5 arg names and the As String return.
    foreach ($needle in @('a As String','b As String','c As String','d As String','e As String','As String')) {
        if ($m7[0].signature -notmatch [regex]::Escape($needle)) {
            throw "WideSig signature missing '$needle' — got: $($m7[0].signature)"
        }
    }
    # Sanity: the sibling TinyOne in the same module is still discoverable
    # (the scanner's line-skip after consuming continuation lines must not
    # skip past unrelated procedures).
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c8","cmd":"analyze_vba","op":"definition_of","name":"TinyOne"}' -Encoding UTF8
    $r8 = Wait-ForEvent $eventsFile 'vba_analysis_result' 10 -IdMatch 'c8'
    if (-not $r8) { throw "no vba_analysis_result for c8 (TinyOne)" }
    $m8 = @($r8.matches)
    if ($m8.Count -ne 1) { throw "expected 1 match for TinyOne, got $($m8.Count) — continuation-line skip over-advanced?" }
    if ($m8[0].module -ne 'modMulti') { throw "TinyOne module expected 'modMulti', got '$($m8[0].module)'" }
    Write-Host "  (5) multi-line signature: WideSig discoverable with joined signature; TinyOne sibling unaffected" -ForegroundColor Green

    # Close
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c9","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no 'closed' event" }
    Write-Host "  closed" -ForegroundColor Green

    $proc.WaitForExit(8000) | Out-Null
    if (-not $proc.HasExited) { throw "host did not exit cleanly" }

    Write-Host "PASS test-analyze-vba" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "FAIL test-analyze-vba: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) {
        Write-Host "Events:" -ForegroundColor Yellow
        Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" }
    }
    exit 1
}
finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
