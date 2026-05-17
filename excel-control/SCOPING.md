# excel-control — scoping

## What this is

A PowerShell-based session interface that lets an AI agent drive any
Excel workbook end-to-end. Reads/writes cells, runs macros, captures
every dialog and UserForm Excel pops as a structured event with a
screenshot, handles runtime errors with module:line, refreshes data
connections, takes screenshots on demand. Works on macro-heavy
workbooks and pure-spreadsheet ones equally.

Bundled into every workbook exported by vba-sync (lives in `tools/`).
Standalone install is supported but vba-sync's Export is the primary
distribution channel.

## Vision

Excel is the most-used software-development surface on Earth. Workbooks
are programs, ranges are state, COM is the runtime. But it's opaque to
AI on the *execute* side — existing integrations operate at the
surface (read a cell, write a cell). They can't run-observe-iterate or
operate a workbook through its real workflow including the dialogs,
forms, and prompts a real user encounters.

excel-control is the **execute-and-observe** half of the AI-Excel loop.
Combined with vba-sync's read-and-edit half, an agent gets the full
development cycle on any Excel workbook:

```
read source → reason → edit → execute → observe → iterate → commit
[vba-sync]                    [excel-control]              [git]
```

## Architecture

A **session** is a long-running PowerShell process that holds an Excel
instance open and exchanges events with the agent over append-only
JSONL files. The agent can be Claude Code (using Bash + Monitor +
Write), plain bash + tail, or anything that can read/write files.

```
sessions/<id>/
├── state.json          pid, workbook path, started-at, status
├── commands.jsonl      agent appends; harness consumes in order
├── events.jsonl        harness appends; agent watches
└── captures/           PNG screenshots referenced by events
```

**Harness reads commands.jsonl** via `FileSystemWatcher` (near-zero
latency on append) with a 2s polling backstop (covers the rare miss).
Tracks last consumed byte offset in `state.json` for resume.

**Agent reads events.jsonl** however its toolchain supports — for
Claude Code that's the `Monitor` tool (push notifications, no polling
on the agent side).

**Dialog handling** is the "reporter, not dismisser" model. Every
modal Excel pops (#32770 or UserForm) becomes a `dialog_appeared` or
`userform_appeared` event with text, buttons, and a screenshot. The
agent decides which button via `respond_dialog` / `respond_form`. The
watcher dispatches the click via WM_COMMAND → BM_CLICK → VK_RETURN →
WM_CLOSE fallback chain with verify-close polling.

**Screenshots:** `PrintWindow` for individual dialog/form HWNDs (works
when obscured); `CopyFromScreen` for the Excel main window.

**No MCP wrapper, no Skill packaging.** Documentation
(`tools/INTERFACE.md`, regenerated on every Export) is the discovery
mechanism. Any agent that can append to a file and tail another file
can drive this.

## Features

### Any workbook (no VBA required)

| Command | Purpose |
|---------|---------|
| `run_macro` | Execute a Sub/Function, get result + event stream |
| `read_range` | Read cell values/formulas from a range |
| `write_range` | Write values to a range |
| `screenshot` | Capture dialog / form / Excel window / specific sheet |
| `respond_dialog` | Click a named button on an open modal |
| `respond_form` | Fill UserForm fields and submit |
| `save_workbook` / `save_as` | Explicit save control |
| `calculate` / `refresh_all` | Application.Calculate / Workbook.RefreshAll |
| `refresh_connection` | Refresh a specific data connection; capture errors |
| `list_sheets` | Sheet name, index, used range, tables, hidden, protected |
| `get_workbook_info` | Size, sheet count, has-vba, has-pivots, has-connections, etc. |
| `open_workbook` / `close_workbook` | Multi-workbook sessions |
| `create_workbook` | New workbook from scratch |
| `abort` / `close` | Cancel running macro / end session |

### Spontaneous events (any workbook)

Excel can surface things the agent didn't ask for:

| Event | When |
|-------|------|
| `dialog_appeared` | Unsolicited modal (link update prompt, etc.) |
| `userform_appeared` | A macro showed a UserForm |
| `connection_failed` | A data refresh failed |
| `workbook_closed` | User closed Excel manually |
| `status_changed` | `Application.StatusBar` updated |

### With VBA (additional commands)

| Command | Purpose |
|---------|---------|
| `compile_check` | VBE compile, returns pass or module:line:message |
| `sync_vba` | Strip + re-import .bas/.cls/.frm from disk (no .xlam dep) |
| `run_tests` | Discover + run `Test_*` Subs, per-test pass/fail/error |
| `list_macros` | All public Subs/Functions across modules with signatures |

### With VBA (additional events)

| Event | When |
|-------|------|
| `runtime_error` | Macro raised an error; includes number, description, module, line, screenshot |
| `debug_print` | `Debug.Print` output drained from the VBE Immediate window |

## Distribution and dependencies

**Bundled via vba-sync.** When the user clicks **Export**, `modSync.WriteHarness`
copies `excel-control/*.ps1` into `<exported workbook>/tools/`, alongside
the regenerated `tools/INTERFACE.md`. Zero setup for downstream agents.

**Standalone install supported** — clone the repo, run scripts directly
against any workbook. No vba-sync dependency.

**No `VBA Sync.xlam` dependency at runtime.** The `sync_vba` command
reimplements the strip-and-import logic directly via COM in PowerShell
(~20 lines) rather than shelling out to `'VBA Sync.xlam'!modSync.ImportProject`.

**Conventions shared with vba-sync** (no formal contract, just a
sensible folder layout):
- `Modules/*.bas` — standard modules
- `ClassModules/*.cls` — class modules
- `Forms/*.frm` (+ `.frx`) — UserForms

This is where vba-sync exports VBA source, and where excel-control
imports it back from. Same folder shape, no spec doc needed.

## Documentation model

- **Repo `README.md`** — mentions excel-control as a feature of vba-sync's Export.
- **Workbook `README.md`** (the one `modSync.WriteReadme` already generates per
  export) — written once, never overwritten. References `tools/INTERFACE.md`
  for the agent-facing harness docs.
- **`tools/INTERFACE.md`** — full agent-facing command + event reference.
  Overwritten on every Export so it always reflects the bundled harness
  version. This is what the agent reads to learn the interface.
- **`tools/.claude/settings.json`** — pre-approved permissions so Claude
  Code doesn't prompt for every Bash/Read/Write/Monitor call against
  the session files.

## Build plan

One feature per PR. Each phase delivers a working command, a fixture
test, and an INTERFACE.md update.

| # | Feature | Effort |
|---|---------|--------|
| 1 | Test fixture workbooks | 1-2h |
| 2 | Session host + `run_macro` (JSONL + FileSystemWatcher + polling backstop) | 4-6h |
| 3 | `respond_dialog` + dialog reporter | 3-4h |
| 4 | Screenshots (PrintWindow + CopyFromScreen) | 3-4h |
| 5 | `runtime_error` capture | 2-3h |
| 6 | `compile_check` + `sync_vba` (PS-native, no .xlam) | 3h |
| 6.5 | `read_range` + `write_range` | 1-2h |
| 6.7 | `Debug.Print` capture | 2h |
| 6.8 | `list_macros` + `list_sheets` + `get_workbook_info` | 2h |
| 7 | `run_tests` + `clsAssert.cls` | 3-4h |
| 7.5 | `save_workbook` + `calculate` + `refresh_all` + `refresh_connection` | 2-3h |
| 8 | `respond_form` + UserForm introspection | 4-5h |
| 8.5 | `open_workbook` + `close_workbook` | 1-2h |
| 8.7 | `create_workbook` | 1h |
| 9 | `WriteHarness` in `modSync.bas` + `INTERFACE.md` + `.claude/settings.json` + bundling | 2-3h |

Total: ~34-44h.

## Test strategy

`excel-control/tests/fixtures/` — 6 generated `.xlsm` files exercising
the dialog/form/error/lock cases. Built reproducibly by a generator
script.

`excel-control/tests/test-*.ps1` — one ps1 per scenario. Self-contained
(spawns its own Excel PID, asserts via throw on failure, kills orphan
EXCEL.EXE on teardown).

`excel-control/tests/run-all.ps1` — discovers and runs all `test-*.ps1`
serially (Excel COM doesn't parallelise well). Per-test timeout. Exit
code = fail count.

No CI initially — local-run only. GitHub Actions Windows runner is
viable later but flaky-Excel-in-CI is a known pain.

## Decisions locked in

- **Session-based**, not one-shot CLI commands. The dialog/form flow
  requires a long-running channel.
- **File-based JSONL append-only protocol.** Universal across agent
  toolchains. No MCP, no Skill metadata, no special IPC libraries.
- **`FileSystemWatcher` + 2s polling backstop** for command reads on
  the harness side.
- **`PrintWindow` for dialog/form capture**, `CopyFromScreen` for the
  main window.
- **Reporter, not dismisser.** The watcher reports every modal as an
  event and only acts on explicit `respond_*` commands.
- **One session = one spawned Excel PID.** No attaching to existing
  Excel instances. Multiple workbooks in one session = yes.
- **Hard fail on missing dependencies** — PowerShell, tar.exe,
  System.Drawing not present → exit visibly, no silent degradation.
- **PowerShell scripts**, not a compiled binary. Lowest install
  friction; ships on every Windows 10+ box.
- **`sync_vba` reimplements import in pure PowerShell** — no
  `VBA Sync.xlam` runtime dependency.
- **Same repo as vba-sync**, two directories, one top-level README.

## Out of scope

- MCP server wrapper (documented interface + file IPC is sufficient)
- Claude Code Skill metadata (same)
- Read/write of cell formatting (fill, font, borders) — belongs in
  VBA; out-of-VBA cell IO is for values + formulas only
- Breakpoints / step debugging in VBE — too much surface area
- UI interaction beyond dialogs/forms (no ribbon clicks, no menu drive)
- macOS support (Windows COM only)
- General Office automation (Word, PowerPoint, etc.)
- Recording / replaying interactive sessions
- Pivot tables / Power Query / DAX automation (use sbroenne/mcp-server-excel
  if you need that; excel-control is for session-based workbook driving)

## Open questions

- **Session lifetime on workbook close.** If the user closes the workbook
  in Excel directly, session should emit `workbook_closed` and stay
  alive or auto-shut down?
- **Concurrent commands.** Block while one command is running, or queue
  with command_ack ordering? Probably block — Excel COM doesn't like
  reentry.
- **Screenshot redaction.** Opt-in `redact_text` flag (default off) or
  always rely on agent discretion? Probably the flag — useful for known
  password fields.
- **Events.jsonl rotation.** Defer until it actually bites; document
  10MB as a soft cap.

## Current state (as of this commit)

Three working PowerShell scripts on the `excel-control` branch:

- `bidirectional-dialog-watcher.ps1` — the v2 "reporter" watcher.
  Single ModalFile/ActionFile IPC. Will be subsumed by the JSONL
  session protocol in Phase 3 (the single-file IPC goes away
  entirely — no backward compat).
- `unlock-vba-project.ps1` — drives the VBE Tools menu (control id
  2578) to surface the VBAProject Password dialog, fills via
  WM_SETTEXT + WM_COMMAND IDOK. Generic utility, kept.
- `sync-vba-to-workbook.ps1` — composed flow that opens a workbook,
  optionally unlocks VBA, runs vba-sync's `ImportProject` via COM,
  saves, closes. Will be rewritten in Phase 6 as the `sync_vba`
  session command (no `.xlam` dependency).

Prototype scripts (`harness.ps1`, `dialog-watcher.ps1`,
`find-compile-error.ps1`) remain for the vba-sync inner-loop work.
They sit alongside but are independent of the session protocol.

## Next: Phase 1

Build the test fixture workbooks. Six `.xlsm` files generated by one
reproducible PowerShell script:

- `empty.xlsm` — baseline
- `msgbox.xlsm` — Sub that pops MsgBox vbOKCancel
- `userform.xlsm` — UserForm with Username/Password/RememberMe + OK/Cancel
- `runtime-error.xlsm` — Sub that raises Err 9
- `password-locked.xlsm` — same as msgbox but with VBProject password = "test"
- `obscured.xlsm` — same as msgbox; "obscured" condition is created at test
  time by the driver opening another window over Excel

Generator: `excel-control/tests/fixtures/build-fixtures.ps1`. Commits
both the generator and the resulting `.xlsm` files. Re-runnable to
regenerate.
