# Phase 6.5 test — read_range + write_range.
#
# Three parts:
#   (1) happy-path 3x2 round-trip
#   (2) P2.11 — write_range refuses a ListObject right-edge crossing
#       with reason:"listobject_autoexpand_risk" (no cell modified)
#   (3) P2.11 — trailing-null values columns emit range_write_warning

[CmdletBinding()]
param([string]$SessionId = "test-rio-$(Get-Random -Maximum 99999)")
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
$repo = Resolve-Path (Join-Path $here '..\..')
$startSess = Join-Path $repo 'excel-control\start-session.ps1'
$fixture   = Join-Path $here 'fixtures\empty.xlsm'
$sessions  = Join-Path $repo 'excel-control\sessions'
$sessionDir = Join-Path $sessions $SessionId
if (Test-Path $sessionDir) { Remove-Item -Recurse -Force $sessionDir }

. (Join-Path $PSScriptRoot '_helpers.ps1')
Clear-OrphanSessionExcels $sessions
Start-Sleep -Milliseconds 200

# Stage a small setup macro that builds a ListObject on Sheet1 ranging
# D4:R200 (header in D4:R4, data D5:R200) — header has 15 named cols.
# After the LO is built, test write_range with crossings + trailing nulls.
$bas = @'
Attribute VB_Name = "modLoSetup"
Option Explicit

Public Sub CreateRoListObject()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    ' Wipe any existing LO first (re-runs)
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Delete
    Next lo
    Dim i As Long
    For i = 1 To 15
        ws.Cells(4, 3 + i).Value = "Col" & i  ' headers in D4..R4 (cols 4..18)
    Next i
    ' Add a couple of data rows so the LO has data range > 0
    ws.Cells(5, 4).Value = "row1"
    ws.Cells(6, 4).Value = "row2"
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("D4:R6"), , xlYes)
    lo.Name = "SyncData"
End Sub
'@
$stagedWb = New-StagedWorkbook -Fixture $fixture -Modules @{ modLoSetup = $bas }

Write-Host "Test: range I/O + P2.11 (SessionId=$SessionId)" -ForegroundColor Cyan

$proc = Start-SessionHost -StartSession $startSess -Workbook $stagedWb `
    -SessionId $SessionId -SessionsRoot $sessions

function Read-Events($p) {
    if (-not (Test-Path $p)) { return @() }
    Get-Content -LiteralPath $p | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json }
}
function Wait-ForEvent($p, $T, $S=15, $Id=$null) {
    $end = (Get-Date).AddSeconds($S)
    while ((Get-Date) -lt $end) {
        foreach ($e in Read-Events $p) {
            if ($e.t -eq $T -and (-not $Id -or $e.id -eq $Id)) { return $e }
        }
        Start-Sleep -Milliseconds 200
    }
    return $null
}

try {
    $eventsFile   = Join-Path $sessionDir 'events.jsonl'
    $commandsFile = Join-Path $sessionDir 'commands.jsonl'

    if (-not (Wait-ForEvent $eventsFile 'started' 30)) { throw "no started" }

    # Write a 3x2 grid
    $writeCmd = @{
        id='c1'; cmd='write_range'; sheet='Sheet1'; range='A1'
        values = @( @('Name','Age'), @('Alice',30), @('Bob',25) )
    } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $writeCmd -Encoding UTF8
    $w = Wait-ForEvent $eventsFile 'range_written' 10 -Id 'c1'
    if (-not $w) { throw "no range_written" }
    Write-Host "  wrote $($w.rows)x$($w.cols) into $($w.range)" -ForegroundColor Green

    # Read it back
    $readCmd = @{ id='c2'; cmd='read_range'; sheet='Sheet1'; range='A1:B3' } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $commandsFile -Value $readCmd -Encoding UTF8
    $r = Wait-ForEvent $eventsFile 'range_read' 10 -Id 'c2'
    if (-not $r) { throw "no range_read" }
    if ($r.rows[0][0] -ne 'Name') { throw "A1 expected 'Name', got '$($r.rows[0][0])'" }
    if ($r.rows[1][1] -ne 30)     { throw "B2 expected 30, got '$($r.rows[1][1])'" }
    if ($r.rows[2][0] -ne 'Bob')  { throw "A3 expected 'Bob', got '$($r.rows[2][0])'" }
    Write-Host "  read back: $($r.rows[0][0])/$($r.rows[0][1]), $($r.rows[1][0])=$($r.rows[1][1]), $($r.rows[2][0])=$($r.rows[2][1])" -ForegroundColor Green

    # ---------- (2) P2.11 — refuse a ListObject right-edge crossing ----------
    # Build the LO first via the staged setup macro.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c-imp","cmd":"import"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'import_completed' 60 -Id 'c-imp')) {
        $f = Wait-ForEvent $eventsFile 'import_failed' 2 -Id 'c-imp'
        throw ('no import_completed' + $(if ($f) { " (import_failed: $($f.error))" } else { '' }))
    }
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c-rm","cmd":"run_macro","name":"CreateRoListObject"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'macro_completed' 15 -Id 'c-rm')) {
        $mf = Wait-ForEvent $eventsFile 'macro_failed' 2 -Id 'c-rm'
        throw ('no macro_completed for CreateRoListObject' +
               $(if ($mf) { " (macro_failed: $($mf.error.message))" } else { '' }))
    }
    Write-Host "  ListObject 'SyncData' built at D4:R6 (Cols 4..18, 2 data rows)" -ForegroundColor Green

    # Cross the right edge: target row 5 (inside LO's row span 4-6),
    # write 20 cells anchored at D5 — last write col = D+20-1 = W (col 23),
    # past LO last col R (col 18) by 5. Should be refused with
    # listobject_autoexpand_risk + crossing_edge:'right' + no cell modified.
    $vals20 = ,@(1..20 | ForEach-Object { "v$_" })
    $crossCmd = @{ id='c-x'; cmd='write_range'; sheet='Sheet1'; range='D5'; values=$vals20 } |
                ConvertTo-Json -Compress -Depth 4
    Add-Content -LiteralPath $commandsFile -Value $crossCmd -Encoding UTF8
    $cx = Wait-ForEvent $eventsFile 'range_write_failed' 15 -Id 'c-x'
    if (-not $cx) { throw "no range_write_failed for right-edge crossing (P2.11 guard regressed?)" }
    if ($cx.reason -ne 'listobject_autoexpand_risk') {
        throw "range_write_failed.reason expected 'listobject_autoexpand_risk', got '$($cx.reason)' (error=$($cx.error))"
    }
    if ($cx.crossing_edge -notmatch 'right') {
        throw "crossing_edge expected to include 'right', got '$($cx.crossing_edge)'"
    }
    if ($cx.listobject -ne 'SyncData') {
        throw "listobject expected 'SyncData', got '$($cx.listobject)'"
    }
    Write-Host "  P2.11 right-edge crossing refused (listobject='$($cx.listobject)', edge='$($cx.crossing_edge)')" -ForegroundColor Green

    # Verify nothing actually got written: read W5 (the rightmost cell the
    # blocked write targeted) and expect it empty.
    Add-Content -LiteralPath $commandsFile -Value '{"id":"c-rd","cmd":"read_range","sheet":"Sheet1","range":"W5"}' -Encoding UTF8
    $rd = Wait-ForEvent $eventsFile 'range_read' 10 -Id 'c-rd'
    if (-not $rd) { throw "no range_read for W5 verification" }
    if ($rd.rows[0][0]) { throw "W5 was modified despite refusal — got '$($rd.rows[0][0])' (P2.11 broken)" }
    Write-Host "  P2.11 verified — W5 untouched after refusal" -ForegroundColor Green

    # Force bypass with force_listobject_extend:true should succeed even
    # for the same 20-cell crossing write — proves the guard was skipped.
    $forceCmd = @{ id='c-f'; cmd='write_range'; sheet='Sheet1'; range='D5'
                   values=$vals20; force_listobject_extend=$true } |
                ConvertTo-Json -Compress -Depth 4
    Add-Content -LiteralPath $commandsFile -Value $forceCmd -Encoding UTF8
    $fw = Wait-ForEvent $eventsFile 'range_written' 15 -Id 'c-f'
    if (-not $fw) { throw "force_listobject_extend bypass did not produce range_written" }
    Write-Host "  P2.11 force_listobject_extend bypass works (20-cell write proceeded)" -ForegroundColor Green

    # ---------- (3) P2.11 — trailing-null warning, no refusal ----------
    # Write fully inside LO bounds (D5:F5) but with 3 trailing-null columns
    # in the values array (total 6 cols at D5). The write proceeds; emit a
    # range_write_warning reason:'trailing_null_columns'.
    # The first 3 cols are populated, last 3 are null. Effective write
    # footprint D5..I5 — col I is past LO right edge R? No, R=col 18, I=9.
    # So this is fully inside the LO column span (D=4 to R=18). It just
    # has trailing-null values that suggest array mis-shape.
    $nullsCmd = @{ id='c-n'; cmd='write_range'; sheet='Sheet1'; range='D5'
                   values=,@('p','q','r',$null,$null,$null) } |
                ConvertTo-Json -Compress -Depth 4
    Add-Content -LiteralPath $commandsFile -Value $nullsCmd -Encoding UTF8
    $rw = Wait-ForEvent $eventsFile 'range_written' 10 -Id 'c-n'
    if (-not $rw) { throw "no range_written for trailing-null inside-bounds write" }
    $rwwarn = Wait-ForEvent $eventsFile 'range_write_warning' 5 -Id 'c-n'
    if (-not $rwwarn) { throw "no range_write_warning for trailing-null inside-bounds write (P2.11 warning regressed?)" }
    if ($rwwarn.reason -ne 'trailing_null_columns') {
        throw "warning reason expected 'trailing_null_columns', got '$($rwwarn.reason)'"
    }
    if ([int]$rwwarn.trailing_null_cols -ne 3) {
        throw "trailing_null_cols expected 3, got $($rwwarn.trailing_null_cols)"
    }
    Write-Host "  P2.11 trailing-null warning emitted ($($rwwarn.trailing_null_cols) cols)" -ForegroundColor Green

    Add-Content -LiteralPath $commandsFile -Value '{"id":"c3","cmd":"close"}' -Encoding UTF8
    if (-not (Wait-ForEvent $eventsFile 'closed' 15)) { throw "no closed" }
    $proc.WaitForExit(8000) | Out-Null

    Write-Host "PASS test-range-io" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "FAIL test-range-io: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path $eventsFile) { Get-Content -LiteralPath $eventsFile | ForEach-Object { Write-Host "  $_" } }
    exit 1
} finally {
    Stop-SessionProcesses -HostProc $proc -SessionDir $sessionDir
}
