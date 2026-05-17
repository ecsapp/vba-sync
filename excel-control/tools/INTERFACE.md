# excel-control — agent interface

A long-running PowerShell session holds an Excel instance open and
exchanges events with you (the agent) over append-only JSONL files. You
send commands by appending to `commands.jsonl`; the harness writes
events to `events.jsonl`.

## Quickstart (Claude Code)

```
# Start a session in the background
Bash run_in_background=true: pwsh tools/start-session.ps1 -Workbook X.xlsm -SessionId s1

# Tail events.jsonl — each new line is a push notification
Monitor: tools/sessions/s1/events.jsonl

# Send commands by appending to commands.jsonl
Write/Edit append: tools/sessions/s1/commands.jsonl
   {"id":"c1","cmd":"run_macro","name":"BuildReport"}

# End the session
   {"id":"cN","cmd":"close"}
```

## Session layout

```
sessions/<id>/
├── state.json          pid, workbook, status, last_command_offset
├── commands.jsonl      agent appends; harness consumes in order
├── events.jsonl        harness appends; agent watches
└── captures/           PNG screenshots referenced by events
```

`state.json` is the agent's source of truth for session health.
`status` is one of `ready`, `busy`, `closed`, `crashed`.

## Commands

### `run_macro`

Execute a Sub or Function and get its result.

```json
{"id":"c1","cmd":"run_macro","name":"BuildReport","args":[2026,"Q1"]}
```

Event stream:

```json
{"t":"command_ack","id":"c1","cmd":"run_macro"}
{"t":"macro_completed","id":"c1","name":"BuildReport","result":1247,"duration_ms":4210}
```

If the macro raises an error: `macro_failed` with `error` + `error_type`.

**Phase 2 limitation:** a macro that pops a modal (`MsgBox`, UserForm,
runtime error dialog) will block the COM thread. Phase 3 wires the
dialog watcher; until then, only call macros that you know don't pop
modals.

### `close`

End the session cleanly.

```json
{"id":"cN","cmd":"close"}
```

Event stream:

```json
{"t":"command_ack","id":"cN","cmd":"close"}
{"t":"closing","id":"cN"}
{"t":"closed"}
```

## Events

| Event | Fields | When |
|-------|--------|------|
| `started` | `pid`, `workbook`, `session_id` | Session opened the workbook successfully |
| `command_ack` | `id`, `cmd` | Command received and parsed |
| `macro_completed` | `id`, `name`, `result`, `duration_ms` | Macro returned (Phase 2) |
| `macro_failed` | `id`, `name`, `error`, `error_type` | Macro raised an error (Phase 2) |
| `closing` | `id` | Close command received |
| `closed` | — | Session shut down cleanly |
| `session_error` | `error`, `stack` | Session crashed |
| `command_error` | `error`, `raw` | A line in commands.jsonl wasn't valid JSON |
| `unknown_command` | `id`, `cmd` | Command name not recognised |
| `dialog_appeared` | `id`, `title`, `text`, `buttons`, `class` | A `#32770` dialog (MsgBox, alert, error) appeared (Phase 3) |
| `userform_appeared` | same shape | A UserForm appeared (Phase 3; control-level introspection lands in Phase 8) |
| `dialog_dismissed` | `id` (command id), `dialog_id`, `button` | The dialog watcher clicked your chosen button and the dialog closed |
| `dialog_closed_externally` | `id` (dialog id) | A dialog vanished without a `respond_dialog` command (user closed it, macro ended) |
| `respond_failed` | `id`, `dialog_id`, `reason` | Watcher could not dispatch the click (unknown id, dialog didn't close) |

## `respond_dialog`

Click a named button on an open dialog.

```json
{"id":"c2","cmd":"respond_dialog","dialog_id":"d1","button":"OK"}
```

Event stream:

```json
{"t":"command_ack","id":"c2","cmd":"respond_dialog"}
{"t":"dialog_dismissed","id":"c2","dialog_id":"d1","button":"OK"}
```

If the named button doesn't match (case-insensitive), the watcher
falls back to: the first button, then VK_RETURN, then WM_CLOSE. If
none works it emits `respond_failed`.

UserForms (`*_appeared` with `class: "ThunderDFrame"`) expose their
`Forms.CommandButton` controls only via the form's COM model, not as
Win32 buttons — so `buttons` may be empty. In that case the watcher
sends VK_RETURN, which fires the form's Default button click handler.
Per-control introspection lands in Phase 8.

More commands and events land in subsequent phases.

## Gotchas

- **Never** modify the workbook in Excel manually while a session has it open
- **Always** send a `close` command before killing the session process (so the .xlsm doesn't go into a "recover this file?" state)
- Event ordering: `command_ack` always precedes any other event for the same `id`
- The session is single-threaded; commands are processed in append order. Don't expect concurrency.

## Phase status

This document tracks the harness as it grows. Current: Phase 2 (session
host + `run_macro`). See `excel-control/SCOPING.md` for the full plan.
