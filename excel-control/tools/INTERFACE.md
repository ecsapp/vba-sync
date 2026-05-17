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

More commands and events land in subsequent phases.

## Gotchas

- **Never** modify the workbook in Excel manually while a session has it open
- **Always** send a `close` command before killing the session process (so the .xlsm doesn't go into a "recover this file?" state)
- Event ordering: `command_ack` always precedes any other event for the same `id`
- The session is single-threaded; commands are processed in append order. Don't expect concurrency.

## Phase status

This document tracks the harness as it grows. Current: Phase 2 (session
host + `run_macro`). See `excel-control/SCOPING.md` for the full plan.
