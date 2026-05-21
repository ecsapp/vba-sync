# excel-control — agent interface

A long-running PowerShell session holds an Excel instance open and
exchanges events with you (the agent) over append-only JSONL files.
You send commands by appending to `commands.jsonl`; the harness writes
events to `events.jsonl`.

## Quickstart (Claude Code)

```
# Start a session in the background (headless Excel by default)
Bash run_in_background=true: pwsh tools/start-session.ps1 -Workbook X.xlsm -SessionId s1

# OR visible Excel (user can watch the agent work; worksheet
# screenshots become meaningful)
Bash run_in_background=true: pwsh tools/start-session.ps1 -Workbook X.xlsm -SessionId s1 -Visible

# Tail events.jsonl — each new line is a push notification
Monitor: tools/sessions/s1/events.jsonl

# Send commands by appending to commands.jsonl
Write/Edit append: tools/sessions/s1/commands.jsonl
   {"id":"c1","cmd":"run_macro","name":"BuildReport"}

# End the session
   {"id":"cN","cmd":"close"}
```

## start-session.ps1 parameters

| Param | Default | Purpose |
|-------|---------|---------|
| `-Workbook` | (required) | Path to .xlsm/.xlsx/.xlam to open |
| `-SessionId` | (required) | Folder name under `sessions/` |
| `-SessionsRoot` | `tools/sessions/` | Override the sessions folder location |
| `-Visible` | off (headless) | Show Excel on the desktop. Trade-off: agent operations may steal foreground focus, but `screenshot target=worksheet:<name>` becomes meaningful (headless Excel main window doesn't render content) |
| `-PollMs` | 250 | Commands.jsonl polling cadence (ms) |

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

Each command requires an `id` (caller's choice — usually a counter)
and `cmd` (the command name).

### Execution

| Command | Body | Result event |
|---------|------|--------------|
| `run_macro` | `name`, optional `args[]` | `macro_completed` with `result`, `duration_ms` |
| `compile_check` | — | `compile_result` with `ok` and on fail: `module`, `line`, `source_context` |
| `run_tests` | optional `filter` (regex on Test_* name) | one `test_result` per test + `tests_completed` summary |

### Cell I/O

| Command | Body | Result event |
|---------|------|--------------|
| `read_range` | `sheet`, `range`, optional `include_formulas` | `range_read` with `rows` (2D array) |
| `write_range` | `sheet`, `range`, `values` (2D array) | `range_written` with `rows`, `cols` |

`range` can be a single anchor cell (e.g., `"A1"`) — `write_range`
auto-resizes to the input dimensions.

### Dialog & UserForm response

| Command | Body | Result event |
|---------|------|--------------|
| `respond_dialog` | `dialog_id`, `button` | `dialog_dismissed` or `respond_failed` |
| `set_form_control` | `dialog_id`, `control`, `value` | `form_control_set` or `form_control_failed` |

`respond_dialog` clicks a button by caption — it works for both `#32770`
dialogs and `ThunderDFrame` UserForms (MSForms controls are windowless,
so they are driven via MSAA / IAccessible — the accessibility layer, not
window messages). Match a caption from the `buttons[]` of the
`dialog_appeared` / `userform_appeared` event; `&` accelerators are
ignored.

`set_form_control` sets one UserForm control — typically before you
click OK. A radio button (`control` = its caption) is selected; a
checkbox is toggled to `value` (`true`/`false`); a text box / combobox
is set to `value` (best-effort). Read the `controls[]` list on
`userform_appeared` to see every control's name, role and state.

### Workbook state

| Command | Body | Result event |
|---------|------|--------------|
| `screenshot` | `target` (`window` / `worksheet:<name>` / `dialog:<id>` / `form:<id>`) | `screenshot_captured` with `path`, `width`, `height` |
| `save_workbook` | — | `workbook_saved` |
| `save_as` | `path`, optional `file_format` | `workbook_saved_as` |
| `calculate` | — | `calculated` with `duration_ms` |
| `refresh_all` | — | `refreshed_all` |
| `refresh_connection` | `name` | `connection_refreshed` or `connection_failed` |
| `create_workbook` | optional `save_as`, optional `file_format` | `workbook_created` |
| `open_workbook` | `path` | `workbook_opened` |
| `close_workbook` | `name`, optional `save` | `workbook_closed_cmd` |

### Introspection

| Command | Body | Result event |
|---------|------|--------------|
| `list_macros` | — | `macros_listed` with `macros[]` (`module`, `kind`, `name`, `args`, `public`, `line`) |
| `list_sheets` | — | `sheets_listed` with `sheets[]` (`name`, `index`, `used_range`, `tables`, `hidden`, `protected`) |
| `get_workbook_info` | — | `workbook_info` with `info` (size, sheet count, has_vba, has_pivots, has_connections, …) |

### VBA source

| Command | Body | Result event |
|---------|------|--------------|
| `sync_vba` | `source_dir` | `sync_completed` with `imported[]`, `removed[]` |

Strip user VBComponents and re-import `.bas`/`.cls`/`.frm` from
`source_dir` following the vba-sync convention (`Modules/`,
`ClassModules/`, `Forms/`, `Objects/`). Pure-PowerShell — no
`VBA Sync.xlam` runtime dependency.

### Lifecycle

| Command | Body | Result event |
|---------|------|--------------|
| `close` | — | `closing` → `closed` |

## Events

Every command produces a `command_ack` first.

| Event | Fields | When |
|-------|--------|------|
| `started` | `pid`, `workbook`, `session_id` | Session opened the workbook |
| `command_ack` | `id`, `cmd` | Command received and parsed |
| `macro_completed` / `macro_failed` | `id`, `name`, `result` / `error` | run_macro result |
| `compile_result` | `id`, `ok`, on fail: `module`, `line`, `column`, `source_context` | compile_check |
| `test_result` | `module`, `name`, `status` (pass/fail), `duration_ms`, on fail: `error` | per test in run_tests |
| `tests_completed` | `id`, `total`, `passed`, `failed`, `duration_ms` | end of run_tests |
| `range_read` / `range_written` | see above | read/write_range |
| `screenshot_captured` / `screenshot_failed` | see above | screenshot |
| `workbook_*` events | varies | workbook lifecycle |
| `calculated` / `refreshed_all` / `connection_refreshed` / `connection_failed` | varies | calc/refresh |
| `macros_listed` / `sheets_listed` / `workbook_info` | see above | introspection |
| `sync_completed` / `sync_failed` | see above | sync_vba |
| `dialog_appeared` / `userform_appeared` | `id`, `title`, `text`, `buttons[]`, `class`, `screenshot`; UserForms also `controls[]` — each `name`, `role`, `value`, `checked`, `enabled` | unsolicited modal observed |
| `dialog_dismissed` | `id` (command id), `dialog_id`, `button` | watcher dispatched click |
| `dialog_closed_externally` | `id` (dialog id) | dialog vanished without `respond_dialog` |
| `dialog_activated` | `polls` (consecutive deferred polls) | `-Visible` mode only — watcher brought Excel to the foreground to surface a modal it had deferred |
| `respond_failed` | `id`, `dialog_id`, `reason`, on a missing button `available[]` | dispatch couldn't close dialog |
| `form_control_set` / `form_control_failed` | `id`, `dialog_id`, `control`; on success `role`/`checked`/`value`; on fail `reason` (+ `available[]`) | `set_form_control` result |
| `runtime_error` | `id` (dialog id), `number`, `description`, `text`, `screenshot` | VBA runtime-error dialog auto-End'd |
| `closing` / `closed` | `id` (for `closing`) | end-of-session |
| `session_error` | `error`, `stack` | session crashed |
| `command_error` | `error`, `raw` | invalid JSON in commands.jsonl |
| `unknown_command` | `id`, `cmd` | command name not recognised |

## Patterns

### Run a macro that pops a dialog

```jsonc
{"id":"c1","cmd":"run_macro","name":"AskOverwrite"}
// watch for:
{"t":"dialog_appeared","id":"d1","title":"...","text":"Overwrite existing?","buttons":["Yes","No"],"screenshot":"..."}
// decide, send:
{"id":"c2","cmd":"respond_dialog","dialog_id":"d1","button":"Yes"}
// then:
{"t":"dialog_dismissed","id":"c2","dialog_id":"d1","button":"Yes"}
{"t":"macro_completed","id":"c1","name":"AskOverwrite"}
```

### Drive a UserForm (pick an option, click OK)

```jsonc
{"id":"c1","cmd":"run_macro","name":"SyncToolPush"}
// watch for the UserForm, with its controls:
{"t":"userform_appeared","id":"d1","title":"Resource Units Mode","class":"ThunderDFrame",
 "buttons":["OK","Cancel"],
 "controls":[{"name":"Units","role":"radiobutton","checked":false,"enabled":true},
             {"name":"Units/Time","role":"radiobutton","checked":false,"enabled":true},
             {"name":"OK","role":"button","checked":false,"enabled":true}]}
// select the option, then click OK:
{"id":"c2","cmd":"set_form_control","dialog_id":"d1","control":"Units","value":"true"}
// → {"t":"form_control_set","id":"c2","dialog_id":"d1","control":"Units","role":"radiobutton","checked":true}
{"id":"c3","cmd":"respond_dialog","dialog_id":"d1","button":"OK"}
// → {"t":"dialog_dismissed","id":"c3","dialog_id":"d1","button":"OK"}
{"t":"macro_completed","id":"c1","name":"SyncToolPush"}
```

### Diagnose a compile error, fix, retry

```jsonc
{"id":"c1","cmd":"compile_check"}
// → {"t":"compile_result","id":"c1","ok":false,"module":"modSales","line":47,"source_context":[...]}
// fix source on disk, then:
{"id":"c2","cmd":"sync_vba","source_dir":"."}
{"id":"c3","cmd":"compile_check"}
// → {"t":"compile_result","id":"c3","ok":true}
```

## Gotchas + limitations

- **Never** modify the workbook in Excel manually while a session has it open.
- **Always** send `close` before killing the session process (otherwise the .xlsm may go into "recover this file?" state).
- Event ordering: `command_ack` always precedes any other event for the same `id`.
- Session is single-threaded; commands process in append order.

### Running alongside the user's interactive Excel

The session host spawns its **own** `EXCEL.EXE` process (separate PID
from any Excel the user has open). The two are mostly isolated:

- ✅ Separate VBA runtime, memory, command bars
- ⚠️  **File locks** — if both processes open the same `.xlsm` read-write, the OS denies the second. Don't open the same workbook in both
- ⚠️  **Add-ins** auto-load in BOTH processes from `%APPDATA%\Microsoft\AddIns\` — usually harmless
- ⚠️  **MsgBox / runtime-error dialogs** from the headless session POP ON THE USER'S DESKTOP — Win32 dispatches them to the foreground session. The dialog watcher auto-dismisses these via WM_COMMAND but a fast user click can race the watcher
- ⚠️  **Foreground focus** — if `-Visible` is set, agent operations may steal focus from windows the user is working in

If concurrent use is essential, prefer `-Visible` so the user can SEE
what the agent is doing instead of being surprised by stray dialogs.

### Known limitations (deferred)

- **`Debug.Print` capture** — in headless Excel COM, `Debug.Print`
  output doesn't surface in `VBE.Windows("Immediate").CodePane`. The
  capture code is in place but yields no events in headless mode.
  Use a custom logger that writes to a file if you need runtime output.
- **UserForm text boxes** — buttons, radio buttons and checkboxes on a
  modal UserForm are fully driven via MSAA (see `set_form_control`).
  Setting a **text box / combobox** value goes through `accValue`, which
  some MSForms edit controls expose read-only — `set_form_control`
  reports `form_control_failed` (`set_failed: …`) when that happens. For
  a form that must be pre-filled with text, a VBA wrapper that seeds the
  fields before `.Show` is still the most reliable route.
- **`.frx` non-determinism** — Excel's UserForm binary writer is
  non-deterministic. Separate work item being tracked outside this
  harness.
