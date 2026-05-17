# excel-control — scoping doc

Working scoping doc for the `excel-control` branch. Captures the
decisions made so far, what's locked, what's open. Update as we go.

## Vision

Excel is the most-used software development surface on Earth — VBA is a
real language, workbooks are programs, macros are functions, ranges are
state, COM is the runtime. But it's opaque to AI on the **execute** side.
Existing integrations (Office MCP, Copilot) operate at the surface: read a
cell, write a cell. They can't refactor VBA, can't run-observe-iterate,
can't actually *operate* a workbook through its real workflow — including
the dialogs, forms, and prompts a real user encounters.

`excel-control` is the **execute-and-observe** half of the AI-Excel loop.

Combined with vba-sync's read-and-edit half, an AI agent gets the full
development cycle on any Excel workbook:

```
read source → reason → edit → execute → observe → iterate → commit
[vba-sync]                    [excel-control]              [git]
```

## The core abstraction: a session, not a CLI

Initial framing was "one-shot verb scripts: run-macro, compile-check,
run-tests". That's too narrow. Real Excel workflows have **dialogs,
UserForms, runtime errors, save-changes prompts, connection-credential
prompts**. A one-shot script either has to dismiss all of those blindly
(losing information and breaking the workflow) or wait for user input
(impossible headless). Neither matches what an AI agent actually needs.

The correct abstraction is a **session**: a long-running PowerShell
process that holds an Excel instance open and exchanges events with the
agent over an append-only log. Verb scripts (`run-macro.ps1`,
`compile-check.ps1`, `run-tests.ps1`) become thin convenience wrappers
that spin up a one-shot session for simple cases. The real interface is
the session itself.

## Bidirectional event flow

The session is a true two-way channel:

**Agent → Excel (commands):** run a macro, set a property, compile-check,
close the workbook, respond to an open dialog/form.

**Excel → Agent (events):** macro completed (with result), runtime error
raised (with line + description), a dialog appeared (with title, text,
buttons, screenshot), a UserForm opened (with name, control list,
screenshot), a status-bar message changed.

**Agent → Excel (responses to events):** click a specific button on the
dialog Excel just popped, fill a UserForm field with a value, abort the
running macro, retry.

This means **every dialog Excel shows is now an interface to the AI**,
not an obstacle. The Win32 dialog watcher's role flips from "dismiss
everything" to "report everything; act only on explicit response from the
agent". The same shift applies to UserForms — currently invisible to
automation, soon a first-class event the agent can introspect and fill.

## Screenshots are first-class

Text introspection of dialogs and forms (window class, control labels,
button captions) is necessary but not sufficient. The AI needs visual
context to handle:

- Custom controls that don't expose readable text via Win32
- Charts visible in the Excel window
- Conditional formatting and layout cues
- Dialogs whose text is the same but visual state differs
- Confirmation that what the agent *thinks* is on screen actually is

So every dialog event, every UserForm event, and the Excel main window
itself ships with a **screenshot saved to a PNG**. The event JSON
includes the path. Agent reads text + looks at image via its vision
model.

Mechanism: PowerShell + `System.Drawing.Graphics.CopyFromScreen` for the
Excel main window; Win32 `PrintWindow` API for individual dialog/form
HWNDs (works even when the window is partially obscured). Screenshots
land in `sessions/<id>/captures/` with predictable names.

## Session protocol (file-based JSONL)

A session lives in a folder. Both sides poll append-only files. Universal
across agent tools — works with Claude Code Monitor, Cursor, OpenClaw, or
any script that can read/write a file.

```
sessions/abc123/
├── state.json          (pid, workbook path, started-at, status)
├── commands.jsonl      (agent appends; harness consumes in order)
├── events.jsonl        (harness appends; agent watches via Monitor)
└── captures/           (PNG screenshots referenced by events)
```

### Command shapes (agent → Excel)

```json
{"id": "c1", "cmd": "run_macro", "name": "DoStuff", "args": [42]}
{"id": "c2", "cmd": "compile_check"}
{"id": "c3", "cmd": "respond_dialog", "dialog_id": "d7", "button": "OK"}
{"id": "c4", "cmd": "respond_form", "form_id": "d9",
 "fill": {"Username": "arnaud", "Password": "***"}, "submit": "OK"}
{"id": "c5", "cmd": "screenshot", "target": "window"}
{"id": "c6", "cmd": "abort"}
{"id": "c7", "cmd": "close"}
```

### Event shapes (Excel → agent)

```json
{"t": "started", "pid": 12345, "workbook": "X.xlsm"}
{"t": "command_ack", "id": "c1"}
{"t": "dialog_appeared", "id": "d7", "title": "Microsoft Excel",
 "text": "Save changes?", "buttons": ["Yes", "No", "Cancel"],
 "screenshot": "captures/d7.png"}
{"t": "userform_appeared", "id": "d9", "name": "LoginForm",
 "fields": [{"name": "Username", "type": "TextBox"},
            {"name": "Password", "type": "TextBox", "password": true}],
 "buttons": ["OK", "Cancel"],
 "screenshot": "captures/d9.png"}
{"t": "runtime_error", "id": "c1", "number": 9,
 "description": "Subscript out of range", "module": "modSales", "line": 47,
 "screenshot": "captures/err_c1.png"}
{"t": "macro_completed", "id": "c1", "result": "OK", "duration_ms": 1340}
{"t": "status_changed", "text": "Processing 50000 rows..."}
{"t": "window_screenshot", "id": "c5", "path": "captures/w_c5.png"}
{"t": "closed"}
```

Every event that involves visible state includes a `screenshot` field
pointing to a PNG. The agent can choose to read it or skip it.

## Scope: what's IN

| Capability                          | Why it's here                       |
|-------------------------------------|--------------------------------------|
| Session start/stop                  | The core abstraction                 |
| Run a named macro                   | Execute VBA from outside             |
| Compile-check the VBA project       | Quick gate before running anything   |
| Report dialogs as events            | The bidirectional channel            |
| Report UserForms as events          | Same — extended to custom forms      |
| Respond to dialogs/forms            | Closes the loop                      |
| Screenshot dialogs / forms / window | Visual context for vision models     |
| Capture runtime errors with context | Error-aware iteration                |
| Test runner                         | Define + run test suites end-to-end  |

Convenience wrappers (verb scripts on top of the session):
- `run-macro.ps1` — open session, run one macro, return result, close
- `compile-check.ps1` — open session, compile, surface errors, close
- `run-tests.ps1` — open session, run a test suite, report, close

The verb scripts cover the simple cases; agents that want interactive
control speak to the session directly.

## Scope: what's OUT

**Read/write/eval of cell data is intentionally out of scope.** VBA inside
the workbook already does this. If an agent needs to read a range, it
should call a macro that returns the data. We don't duplicate VBA's job
from outside.

Also out:

- UI interaction beyond dialogs/forms (no clicking ribbon, no menu navigation)
- Replacement for the vba-sync addin (humans still click "Export" in Excel)
- Abstraction layer over the Excel object model (no DSL)
- Parsing OOXML directly (we drive live Excel via COM)
- macOS support (Windows COM only)
- A general-purpose Office automation framework (Word, PowerPoint, etc.)
- Recording / replaying user macros — separate problem space

## Distribution

Two channels, in order:

1. **Bundled in every vba-sync export.** When a user clicks "Export
   VBA", the resulting workbook repo gets a `tools/` folder containing
   the excel-control scripts. Zero install, ships with the source.
2. **Standalone (later).** Cloneable / installable independently for use
   against any .xlsm without vba-sync involvement.

## Primary user

- **AI agents** (Claude Code, GPT, Cursor, OpenClaw, etc.) operating on
  a workbook repo end-to-end
- Secondary: humans writing CI / regression pipelines for Excel workbooks
- Tertiary: humans debugging a macro interactively from a terminal

## Success criteria

- An AI agent given a workbook repo can: open a session, run a macro,
  see the dialog the macro pops, decide based on its text and
  screenshot, click the right button, observe the result, fix a compile
  error, re-run, run a test suite — all from one session, with zero
  manual intervention.
- Every dialog the workbook can produce becomes a reportable event
  before the agent has to know about it. No "I didn't see that dialog"
  surprises.
- Screenshot capture works on dialogs whether visible or fully obscured
  (PrintWindow handles the offscreen case).
- Reliable every time on Windows 10+ / Office 365 — no flaky 1004s, no
  race conditions, no manual dismissal.
- Each verb script is a thin (~100-line) wrapper over the session;
  session itself is the substantive code (estimated ~800 lines).

## Claude Code integration patterns

Specifically how a Claude Code session uses this:

1. **Start session in background:**
   `Bash` with `run_in_background=true` invokes
   `tools/start-session.ps1 -Workbook X.xlsm -SessionId abc123`.
   The process stays alive; Claude Code returns immediately with the
   session id.

2. **Watch events via `Monitor`:** Monitor tails
   `tools/sessions/abc123/events.jsonl`. Each new line is a notification
   to the agent. The agent reads `t` to dispatch on event type, reads
   `screenshot` path if visual context is needed.

3. **Send commands via `Write`/`Edit`:** the agent appends to
   `tools/sessions/abc123/commands.jsonl`. The harness's own polling
   loop picks them up.

4. **End the session:** append `{"cmd": "close"}` and unmonitor.

This is symmetric file-polling. No special pipes, no IPC libraries, no
Excel-specific protocols on the agent side. Anything that can read and
write files can drive this.

## Architecture decisions (locked)

- **PowerShell scripts**, not a compiled binary. Lowest install
  friction; PS5+ ships on every Windows 10+ box.
- **Session is primary; verb scripts are wrappers.** Don't design the
  scripts first and bolt on a session — the other way around.
- **File-based JSONL append-only protocol** for the bidirectional
  channel. Universal across agent toolchains.
- **One screenshot per visible-state event.** PNG, lossless, in
  `sessions/<id>/captures/`. Path included in the event JSON.
- **Win32 PrintWindow for dialog/form capture** — works even when the
  target is fully obscured. CopyFromScreen for the main Excel window.
- **Synchronous extraction primitive** (already proven in vba-sync via
  `tar.exe`) for any zip work — no race conditions.
- **Hard fail on missing dependencies.** If PowerShell, tar.exe, or
  System.Drawing is missing/blocked, fail visibly. No silent
  degradation.
- **Win32 dialog watcher = bidirectional event source.** Not a
  "dismisser". It reports everything and only acts when told.

## Open questions

- **Session lifetime.** One session per workbook open, or can one
  session host multiple workbooks?
- **Concurrent sessions.** Can two sessions run against the same Excel
  instance? Likely no (Excel is single-threaded). Different instances?
- **Screenshot redaction.** UserForms may contain passwords or
  sensitive data. Should we offer a `redact_text` flag, or rely on the
  agent to decide?
- **Test runner shape.** What does a test look like? A VBA `Sub` with a
  naming convention (`Test_*`)? A separate `.tests.bas`? A test
  manifest? Need to design.
- **Standalone install path.** When (not if) we support standalone,
  where? `%APPDATA%`? `~/.excel-control/`? `Program Files`?
- **Versioning.** Independent of vba-sync, or aligned? Probably
  independent.
- **MCP wrapper (later).** Wrap the session protocol as MCP tools so
  any MCP-capable client (Claude desktop, Cursor) can call it natively.
- **Auth/credentials handling.** Connection prompts will ask for
  passwords. Pass through? Vault integration? Out of scope?

## Current state

`harness.ps1`, `dialog-watcher.ps1`, `find-compile-error.ps1` in this
folder are the **prototype** built during vba-sync's
`excel-export-rewrite` work. They were built to drive vba-sync
development from CLI; they happen to demonstrate every fundamental
capability listed above except screenshots, session protocol, and
form-response.

`PROTOTYPE_README.md` is the original developer-facing README for the
prototype. Kept as reference; will be superseded by per-script + session
docs.

The productization work this branch will do:

1. **Design the session protocol fully** (JSONL schema, error paths,
   timeout behaviour, file rotation if events.jsonl grows large)
2. **Implement `start-session.ps1`** — the core long-running process
3. **Add screenshot capture** for dialogs, UserForms, Excel main window
4. **Extend dialog watcher** from dismisser to reporter; add
   form-introspection (enumerate UserForm controls)
5. **Build verb wrappers** `run-macro.ps1`, `compile-check.ps1`,
   `run-tests.ps1` on top of the session
6. **Design and implement the test runner**
7. **Add `WriteHarness` to vba-sync's `modSync.bas`** so the export
   emits `tools/` into every workbook repo
8. **Documentation**: per-script user docs, agent-integration guide

## Out of scope for THIS branch

- MCP server wrapper (v2)
- PS module packaging (v2)
- Standalone install path (later)
- macOS support (never)
- General Office automation (never)
- Recording / replaying interactive sessions for replay testing
