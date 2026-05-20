---
name: excel-control
description: Drive this Excel workbook end-to-end — run macros, read/write cells, handle dialogs and runtime errors, refresh data connections, write and run a VBA test suite. Use whenever the user wants to execute, inspect, debug, or test the workbook beyond static source editing.
---

# excel-control — agent guide

This workbook ships with a PowerShell session harness that lets you
drive Excel programmatically. Long-running PS process holds a COM
instance open; you communicate via two append-only JSONL files:

- `tools/sessions/<id>/commands.jsonl` — you append, harness consumes
- `tools/sessions/<id>/events.jsonl` — harness appends, you watch

`tools/INTERFACE.md` is the full API reference. This file is the guide:
how to do common things.

## Paths

All paths in this guide are **relative to the workbook root** — the
folder vba-sync Export wrote, where the harness lives at `./tools/`. If
you are working in the **excel-control dev tree** instead (the source
repo), the scripts are at `excel-control/` and sessions at
`excel-control/sessions/` — substitute accordingly. The `pwsh` launch
runs from the workbook root.

## Quickstart

1. **Launch the session** (Bash, `run_in_background: true`). Headless by
   default — Excel runs invisibly:
   ```
   pwsh tools/start-session.ps1 -Workbook X.xlsm -SessionId s1
   ```
   Add `-Visible` to show Excel on the desktop — useful when you need
   meaningful worksheet screenshots (a headless automation Excel renders
   a mostly-empty main window). Headless is otherwise fine for every
   command, Export included.
2. **Wait for readiness.** The harness emits
   `{"t":"started","pid":...,"excel_pid":...,"session_id":...}` to
   `tools/sessions/s1/events.jsonl` once Excel has opened the workbook —
   this takes a few seconds. **Do not send a command before `started`
   arrives.** If it has not arrived in ~15s, read
   `tools/sessions/s1/state.json`: `status:crashed` means startup failed
   — check the `events.jsonl` tail for a `session_error`.
3. **Watch events** — point the Monitor tool at
   `tools/sessions/s1/events.jsonl` (one notification per appended line).
4. **Send commands** — append one JSON object per line to
   `tools/sessions/s1/commands.jsonl` (append only; never rewrite it).
5. **End the session** — append `{"id":"cN","cmd":"close"}`.

**Concurrent Excel use**: the session creates a separate `EXCEL.EXE`
process. It is safe to run alongside the user's interactive Excel
provided:
- Don't open the same workbook in both (file-lock conflict).
- MsgBoxes / runtime-error dialogs from the harness pop on the user's
  desktop — the watcher auto-dismisses them fast, but a user click can
  race the watcher.
- With `-Visible`, agent operations may steal foreground focus.

Every command needs an `id` (any unique string per session — typically `c1`,
`c2`, …) and a `cmd`. You get one `command_ack` event per command, then the
result event(s).

## Common workflows

### Run a macro with no dialogs expected

```jsonc
// commands.jsonl:
{"id":"c1","cmd":"run_macro","name":"BuildReport","args":[2026,"Q1"]}

// events.jsonl:
{"t":"command_ack","id":"c1","cmd":"run_macro"}
{"t":"macro_completed","id":"c1","name":"BuildReport","result":1247,"duration_ms":4210}
```

`args` crosses the COM boundary, which constrains what you can pass:
scalars (numbers, strings, booleans, dates) and arrays of those work;
`Range`, `Workbook`, or other object references do **not**; `ByRef`
out-parameters are not returned — only the function's return value
comes back in `result`. Keep macro signatures to value arguments.

### Run a macro that pops a dialog

The dialog watcher reports every modal as an event with text, buttons,
and a screenshot. Decide which button, send `respond_dialog`, the
macro continues.

```jsonc
{"id":"c1","cmd":"run_macro","name":"RefreshFromDB"}
// →
{"t":"command_ack","id":"c1","cmd":"run_macro"}
{"t":"dialog_appeared","id":"d1","title":"Microsoft Excel","text":"Overwrite existing data?","buttons":["Yes","No","Cancel"],"screenshot":"tools/sessions/s1/captures/d1.png"}
// you read the screenshot if needed, then:
{"id":"c2","cmd":"respond_dialog","dialog_id":"d1","button":"Yes"}
// →
{"t":"command_ack","id":"c2","cmd":"respond_dialog"}
{"t":"dialog_dismissed","id":"c2","dialog_id":"d1","button":"Yes"}
{"t":"macro_completed","id":"c1","name":"RefreshFromDB","result":null}
```

### Diagnose and fix a runtime error

When a macro hits an unhandled error, Excel's "Microsoft Visual Basic"
dialog appears. The watcher auto-dismisses it (End) and emits a
structured `runtime_error` event with the number and description.

```jsonc
{"id":"c1","cmd":"run_macro","name":"BuildReport"}
// →
{"t":"runtime_error","id":"d1","number":9,"description":"Subscript out of range","screenshot":"..."}
{"t":"macro_failed","id":"c1","name":"BuildReport","error":"..."}
```

Fix loop:
1. Read the offending module to find the error site (you may need to
   `list_macros` first to know which modules exist, or `sync_vba` from
   your source-on-disk and then `compile_check`).
2. Edit the `.bas` / `.cls` on disk.
3. `sync_vba` to push the fix into the running Excel.
4. `compile_check` — confirm no parse errors.
5. `run_macro` again.

```jsonc
{"id":"c2","cmd":"sync_vba","source_dir":"."}
// → {"t":"sync_completed","id":"c2","imported":["modSales"],"removed":[]}

{"id":"c3","cmd":"compile_check"}
// → {"t":"compile_result","id":"c3","ok":true}

{"id":"c4","cmd":"run_macro","name":"BuildReport"}
// → {"t":"macro_completed","id":"c4","result":1247}
```

### Write and run a test suite

The harness has a discover-and-run convention for VBA tests:

- Any `Sub Test_<Name>()` in any module gets discovered by `run_tests`
  — the `Public` keyword is optional. Only a `Sub` is matched: a
  `Function` named `Test_*`, or a `Private Sub Test_*`, is **not**
  discovered and silently never runs.
- A test **passes** if it runs to completion without raising an error.
- A test **fails** if it raises any error (via `Err.Raise` or an
  unhandled runtime error).

The shipped `tools/clsAssert.cls` is a minimal assertion library. Import
it into your workbook (via `sync_vba`) and use its `Assert.AreEqual`,
`Assert.IsTrue`, `Assert.IsFalse`, `Assert.IsNothing`, `Assert.Fail`
methods. Each raises a vbObjectError on mismatch, which the test
runner catches.

Example test module (drop this in `<your-repo>/Modules/modSalesTests.bas`):

```vba
Attribute VB_Name = "modSalesTests"
Option Explicit

Public Sub Test_Sales_BasicTotal()
    Dim total As Long
    total = ComputeTotal(50, 50)
    Assert.AreEqual 100, total
End Sub

Public Sub Test_Sales_RejectsNegatives()
    On Error Resume Next
    ComputeTotal -1, 0
    If Err.Number = 0 Then Assert.Fail "expected error on negative qty"
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub Test_Sales_HappyPath_Refresh()
    Refresh_Sales_Sheet
    Assert.IsTrue Range("Summary!B12").Value > 0, "expected non-zero in B12"
End Sub
```

To run them:

```jsonc
// Sync your source (includes clsAssert.cls + your test modules):
{"id":"c1","cmd":"sync_vba","source_dir":"."}
// → {"t":"sync_completed","id":"c1","imported":["clsAssert","modSalesTests",...],"removed":[]}

// Optional: compile first to catch typos
{"id":"c2","cmd":"compile_check"}
// → {"t":"compile_result","id":"c2","ok":true}

// Run everything matching Test_Sales_*
{"id":"c3","cmd":"run_tests","filter":"^Test_Sales_"}
// →
{"t":"test_result","module":"modSalesTests","name":"Test_Sales_BasicTotal","status":"pass","duration_ms":12}
{"t":"test_result","module":"modSalesTests","name":"Test_Sales_RejectsNegatives","status":"pass","duration_ms":8}
{"t":"test_result","module":"modSalesTests","name":"Test_Sales_HappyPath_Refresh","status":"fail","duration_ms":420,"error":{"message":"AreEqual failed: expected=100 actual=0","hresult":-2147221503}}
{"t":"tests_completed","id":"c3","total":3,"passed":2,"failed":1,"errored":0,"duration_ms":520}
```

The emitter always includes `errored` in `tests_completed`; it is
currently always `0` (the runner has no separate error category — see
the next note), so rely on `passed`/`failed`.

Interpreting:
- `tests_completed.passed`/`failed` is your summary
- Each `test_result.status` is `pass` or `fail` (no separate `error` category;
  unhandled errors look the same as failed assertions to the runner)
- `test_result.error.message` carries the assertion message (or the
  VBA error description if it was an unhandled error)

### Inspect a workbook you've never seen

When meeting an unfamiliar workbook, three commands give you the lay of
the land:

```jsonc
{"id":"c1","cmd":"get_workbook_info"}
// → {"t":"workbook_info","id":"c1","info":{"name":"X.xlsm","size_bytes":118000,
//      "sheet_count":12,"table_count":4,"has_vba":true,"has_pivots":true,
//      "has_connections":true,"vba_protection":0,"last_author":"alice"}}

{"id":"c2","cmd":"list_sheets"}
// → {"t":"sheets_listed","id":"c2","sheets":[
//      {"name":"Inputs","used_range":"A1:Z48","tables":["tblScenarios"]},
//      {"name":"Calc","used_range":"A1:CC2400","hidden":true,"protected":true},
//      ...]}

{"id":"c3","cmd":"list_macros"}
// → {"t":"macros_listed","id":"c3","macros":[
//      {"module":"modSales","kind":"sub","name":"BuildReport","args":"year As Long, qtr As String","public":true},
//      ...]}
```

The event payloads above are **abridged** — the emitters send more
fields than shown (`workbook_info` adds `full_name`, `file_format`,
`named_range_count`, …; `list_sheets` adds `index`; `list_macros` adds
`line`). See `tools/INTERFACE.md` for the full shape of every event.

### Refresh data and verify

```jsonc
{"id":"c1","cmd":"refresh_connection","name":"OrdersDB"}
// → {"t":"connection_refreshed","id":"c1","name":"OrdersDB","duration_ms":4200}
// OR if it failed:
// → {"t":"connection_failed","id":"c1","name":"OrdersDB","error":"Login failed for user 'svc_acct'"}

{"id":"c2","cmd":"calculate"}
// → {"t":"calculated","id":"c2","duration_ms":1820}

{"id":"c3","cmd":"read_range","sheet":"Summary","range":"B12"}
// → {"t":"range_read","id":"c3","sheet":"Summary","range":"B12","rows":[[1247.5]]}
```

### Scaffold inputs for a scenario

`write_range` writes a 2D array of values to a sheet — useful for
setting up test inputs or scenario assumptions:

```jsonc
{"id":"c1","cmd":"write_range","sheet":"Inputs","range":"A1",
 "values":[["Period","Value"],["Q1",100],["Q2",120],["Q3",95]]}
// → {"t":"range_written","id":"c1","sheet":"Inputs","range":"A1:B4","rows":4,"cols":2}
```

Single-cell anchor (`A1`) auto-resizes to the input dimensions.

## When NOT to use the harness

- **Just reading VBA source** — that's a vba-sync Export; no session
  needed.
- **Pure source edits** that don't need execution — use vba-sync's
  Import (faster, no session overhead).
- **SharePoint-hosted workbooks** that open from a URL — the harness
  works only against local or synced files.
- **Reading/writing cell *formatting*** (fill, font, borders) — the
  harness exposes values and formulas only. For formatting use VBA
  inside the workbook.
- **Driving the ribbon or menus** — out of scope. Dialogs and
  UserForms triggered by code are in scope.

## Anti-patterns (will cause headaches)

- **Don't open the workbook manually in Excel** while a session has it
  open — file lock conflict.
- **Always send `close`** before killing the session process; otherwise
  Excel may flag the .xlsm as "needing recovery" on next open.
- **After editing `.bas`/`.cls` on disk during a live session, always
  `sync_vba`** before `compile_check` / `run_macro` — the harness only
  sees synced source; the in-memory VBA project does not track disk.
- **Don't open the same workbook twice** — use one session per
  workbook (multi-workbook within a session is fine via
  `open_workbook` / `close_workbook`).

## Debugging stuck or unhappy sessions

Look at `tools/sessions/<id>/state.json`:

```json
{
  "pid": 12345,
  "excel_pid": 23320,
  "workbook": "C:\\path\\X.xlsm",
  "session_id": "s1",
  "status": "busy",
  "visible": false,
  "started_at": "2026-05-19T10:30:00Z",
  "last_command_offset": 482
}
```

- `pid` — the PowerShell **host** process.
- `excel_pid` — the **spawned Excel** process. This is the one to kill
  if you ever must force-terminate — never `Stop-Process -Name EXCEL`,
  which would also kill the user's interactive Excel.
- `status: ready` — idle, waiting for commands
- `status: busy` — executing a command (likely run_macro)
- `status: closed` — clean shutdown
- `status: crashed` — host died unexpectedly; check `events.jsonl`
  tail for `session_error`

**Session hangs (no events arriving):**
1. Check `state.json` — if `status: busy` for a long time, the macro
   is probably stuck on a modal you haven't responded to. Look at the
   last `dialog_appeared` / `userform_appeared` event — does it
   have a `dialog_id` you should respond to?
2. If the dialog is something unexpected (e.g., Excel's "file recovery"
   dialog from a prior crash), screenshot the host's main window via
   `{"id":"cN","cmd":"screenshot","target":"window"}` to see what's
   on screen.
3. Last resort: if no response to commands at all and the host is
   alive but unresponsive, terminate `state.json.pid` (the host) and,
   if it is orphaned, `state.json.excel_pid` (the spawned Excel) — never
   all of Excel — then start a fresh session.

**Resuming after a crash:** a session directory persists after the host
dies. `start-session.ps1` resumes from `last_command_offset` if you
reuse an existing `SessionId`. To pick up where a crashed session left
off, reuse its id; to start clean, `ls tools/sessions/` and choose an
unused one.

**Commands not being processed:**
- `commands.jsonl` is read by byte offset (tracked in
  `state.last_command_offset`). If you accidentally rewrote the file
  instead of appending, the harness may skip your new commands or
  re-read old ones. Always **append**, never overwrite.

**`macro_failed` with cryptic COM error:**
- Look for a paired `runtime_error` event preceding it — that's the
  real VBA error number and description. The `macro_failed` HRESULT
  is COM's translation of the same.

**`respond_dialog` returns `respond_failed`:**
- `unknown_or_closed_dialog` — the dialog vanished before your command
  arrived (user clicked it, or another macro closed it). Watch for
  `dialog_closed_externally` events.
- `dialog_did_not_close` — the watcher's click dispatch chain
  (WM_COMMAND → BM_CLICK → VK_RETURN → WM_CLOSE) failed. Rare.
  Screenshot the dialog and report.

## Known limitations

- **`Debug.Print` capture** — in headless Excel COM, `Debug.Print`
  output doesn't surface in `VBE.Windows("Immediate").CodePane`. The
  capture code is in place but yields no events. Use a custom file
  logger if you need runtime output:

  ```vba
  Public Sub Log(s As String)
      Open ThisWorkbook.Path & "\macro.log" For Append As #1
      Print #1, Format(Now, "hh:nn:ss.000"); " "; s
      Close #1
  End Sub
  ```

- **UserForm field introspection** — modal UserForm `Controls(name).Value`
  isn't reachable from outside the VBA runtime. `respond_dialog` fires
  the form's Default button via VK_RETURN. For forms with input
  fields, either (a) wrap the form-show in a VBA Sub that pre-fills
  fields, or (b) take a screenshot via `screenshot target=form:<id>`
  and read it visually.

- **`.frx` non-determinism** — Excel's UserForm binary writer is
  non-deterministic; .frx files differ between consecutive saves
  even when the form hasn't changed. Tracked separately.

## Reference

Full command and event reference: [tools/INTERFACE.md](../../../tools/INTERFACE.md)

Source: [excel-control](https://github.com/ecsapp/vba-sync/tree/main/excel-control) in the vba-sync repo.
