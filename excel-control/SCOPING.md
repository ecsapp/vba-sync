# excel-control — scoping doc

Working scoping doc for the `excel-control` branch. Captures the decisions
made so far, what's locked, what's open. Update as we go.

## Vision

Excel is the most-used software development surface on Earth — VBA is a
real language, workbooks are programs, macros are functions, ranges are
state, COM is the runtime. But it's opaque to AI on the **execute** side.
Existing integrations (Office MCP, Copilot) operate at the surface: read a
cell, write a cell. They can't refactor VBA, can't run-observe-iterate,
can't actually *operate* a workbook.

`excel-control` is the **execute-and-observe** half of the AI-Excel loop.

Combined with vba-sync's read-and-edit half, an AI agent gets the full
development cycle on any Excel workbook:

```
read source → reason → edit → execute → observe → iterate → commit
[vba-sync]                    [excel-control]              [git]
```

## What we're building

A small set of PowerShell scripts that drive Excel headlessly via COM. One
verb per script. Each script:

- Opens a workbook (no UI)
- Does one thing
- Returns a result (stdout / exit code / JSON)
- Survives Excel's modal nature (dialogs, popups, runtime-error boxes)
- Closes the workbook cleanly

That's it. No abstraction layer over the Excel object model. No DSL. Each
script is ≤200 lines, documented at the top, callable from any AI agent
harness or CI system.

## Scope: what's IN

Only things VBA itself **cannot do from inside the workbook**:

| Script              | Purpose                                                      |
|---------------------|--------------------------------------------------------------|
| `run-macro.ps1`     | Open workbook, run a named macro, capture result + errors    |
| `compile-check.ps1` | Verify VBA project compiles cleanly; surface line of any error |
| `run-tests.ps1`     | Define + run a test suite for the workbook                   |

Internal libraries (used by the above, not direct entry points):

- `dialog-watcher.ps1` — Win32 modal-dismissal helper

## Scope: what's OUT

**Read/write/eval operations are intentionally out of scope.** VBA inside
the workbook can already read and write cells, evaluate formulas, manipulate
ranges. If an AI agent needs to read or write data, the right move is to
ask the workbook's own VBA to do it (via `run-macro.ps1`). We don't
duplicate VBA's job from the outside.

Also out:

- UI interaction (this is headless-only)
- Replacement for the vba-sync addin (humans still click "Export" in Excel)
- An abstraction layer over the Excel object model (no DSL, no wrappers)
- Parsing OOXML directly (we drive live Excel via COM)
- macOS support (Windows COM only)
- A general-purpose Office automation framework (Word, PowerPoint, etc.)

## Distribution

Two channels, in order:

1. **Bundled in every vba-sync export.** When a user clicks "Export VBA",
   the resulting workbook repo gets a `tools/` folder containing the
   excel-control scripts. Zero install, ships with the source.
2. **Standalone (later).** Cloneable / installable independently for use
   against any .xlsm without vba-sync involvement.

## Primary user

- **AI agents** (Claude, GPT, Cursor, etc.) operating on a workbook repo
  end-to-end
- Secondary: humans writing CI / regression pipelines for Excel workbooks

## Success criteria

- An AI agent given a workbook repo can: run an arbitrary macro, observe
  its output, fix a compile error, run a test suite — all from one session,
  with zero manual intervention.
- Reliable every time on Windows 10+ / Office 365 — no flaky 1004s, no
  race conditions, no manual dialog dismissal.
- Each script is self-documenting (header comment explains usage) and
  callable from any agent harness or CI runner.

## Architecture decisions (locked)

- **PowerShell scripts**, not a compiled binary. Lowest install friction;
  PS5+ ships on every Windows 10+ box.
- **One verb per script**, not a monolithic CLI. Easier for an agent to
  pick the right tool by name; easier for a human to read the source.
- **Synchronous extraction** primitive (already proven in vba-sync via
  `tar.exe`) is the model — if a script needs to wait, it waits properly,
  not via heuristic polling.
- **Hard fail on missing dependencies.** If PowerShell or tar.exe is
  missing/blocked, the script raises a visible error. No silent
  degradation.
- **Win32 dialog watcher** as a shared internal library (`dialog-watcher.ps1`)
  used by every script that opens Excel.

## Open questions

- **PS scripts vs PS module?** MVP is scripts. Later: do we package as a
  PS module (`Import-Module excel-control`)?
- **MCP server?** Wrap the scripts as an MCP server so AI clients can call
  them directly via the protocol. Probably v2.
- **Test runner shape.** What does a test look like? A function in a
  user's VBA module marked with a convention? A separate `.tests.ps1`
  file? A `Tests/` folder with one `.bas` per test? Need to design.
- **Standalone install path.** When (not if) we support standalone,
  where does the install live? `%APPDATA%`? `~/.excel-control/`?
  `Program Files`?
- **Versioning.** Excel-control versioned independently of vba-sync, or
  aligned? Probably independent, since the audience overlaps but isn't
  identical.
- **Naming convention for scripts.** `run-macro.ps1`, `compile-check.ps1`
  — verb-noun PowerShell style. Locked.

## Current state

`harness.ps1`, `dialog-watcher.ps1`, `find-compile-error.ps1` in this
folder are the **prototype** built during vba-sync's `excel-export-rewrite`
work. They were built to drive vba-sync development from CLI; they happen
to demonstrate every capability listed above. They are NOT the productized
shape — they're hardcoded to my dev paths and one specific workbook.

`PROTOTYPE_README.md` is the original developer-facing README for the
prototype. Kept as reference; will be superseded by per-script docs.

The productization work this branch will do:

1. Split `harness.ps1` (current monolith — does rebuild + open + run +
   report) into focused scripts: `run-macro.ps1`, `compile-check.ps1`,
   plus a new `run-tests.ps1`
2. Remove hardcoded paths; make defaults relative-to-repo
3. Define and document the test-runner shape
4. Add a `WriteHarness` sub to `modSync.bas` so the vba-sync export emits
   these scripts into every exported workbook's `tools/` folder
5. Update vba-sync README + this branch's README to introduce the tool to
   end users

## Out of scope for this branch

- MCP server (v2)
- PS module packaging (v2)
- Standalone install path (later)
- macOS support (never)
- General Office automation (never)
