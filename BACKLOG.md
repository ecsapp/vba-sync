# Backlog

Open work for vba-sync. Items here are deferred features, known gaps, and
TODOs lifted out of source comments. Review before starting any of them —
some may have been addressed since they were filed.

## In-source TODOs

- **Cell styles via OOXML** — `modExcelExport.bas` currently leaves
  `cellStyles` empty. Read `xl/styles.xml` + `xl/worksheets/sheetN.xml`
  from the OOXML extract (already produced by `BuildDrawingImageMap`)
  and build a per-cell xfid lookup. No per-cell COM needed.
- **Standalone chart sheets** — `wb.Charts` (separate chart-sheet tabs,
  not `ChartObject`s embedded in worksheets) are not currently emitted.

## Deferred features

### Pivot tables

Parse `xl/pivotTables/pivotTable*.xml` for source ref, fields on rows/
cols/values, filters. Output `worksheets/<NN> - <Name>/pivot_tables/<PivotName>/{definition.json, data.tsv}`.
The `data.tsv` snapshots the rendered cached values — useful diff signal
even when the underlying data hasn't changed. Skip `xl/pivotCache/` —
opaque and large; Excel regenerates from source.

Hold for a real workbook with pivots before writing the parser.

### `.vba-sync/config.json`

Per-repo export options. Created on first export, never overwritten.
Schema sketch:

```json
{
  "export": {
    "stripVolatileMetadata": true,
    "includeFormulaCachedValues": false,
    "includeRevisionUIDs": false,
    "includeViewState": false,
    "includeLastModifiedBy": true,
    "data": {
      "defaultMaxRows": null,
      "perSheetOverrides": {
        "Big Dump Sheet": { "maxRows": 500 }
      }
    },
    "drawings": { "copyAssets": true }
  }
}
```

Everything is hardcoded sensible defaults today. Add this when a user
first asks for an option that should be per-repo (e.g. "skip cached
formula values to reduce churn").

### Round-trip import

Major undertaking. Approach: don't write OOXML — drive Excel via VBA +
the object model. For each text file the export wrote, apply it back to
a live workbook:

- `data.tsv` → `Range.Value2 = ...`
- `formulas.json` → `Range.Formula = ...`
- `styles.json` → walk and apply to ranges
- `_meta.json` → `ws.Tab.Color = ...`, freeze panes, etc.
- `tables/*/definition.json` → `ws.ListObjects.Add` then column setup
- `drawings/shapes.json` → `ws.Shapes.AddShape` with OnAction assignment
- `drawings/_assets/*.png` → `ws.Pictures.Insert`
- `validations.json`, `conditional_formats.json`, `comments.json` →
  respective Excel object model APIs

VBA becomes the importer. The importer can be incremental — start with
just `data.tsv` (cell values + formulas), since most of the time that's
what a VS Code edit wants back in Excel.

## Known gaps

Spot-checked 2026-05-21 — these held up. (`Cell styles`, the table
marker row, and the continuation-limit item were verified directly
against current `modExcelExport.bas` / `modSync.bas`.)

- **`tabColor` indexed/theme handling.** Currently handles `rgb=` only.
  Indexed (`indexed=`) and theme (`theme=N`) colours become `null` in
  MANIFEST. Needs `theme1.xml` lookup + indexed-palette table.
- **`SafeFileName` truncation.** Truncates at 100 chars without
  uniqueness check. Excel limits sheet names to 31 chars so collision
  is vanishingly unlikely, but it's a real correctness gap.
- **`tabColor` alpha stripping.** `Substring(2)` drops the `AA` byte
  from `AARRGGBB`. Tabs don't support transparency so it's fine in
  practice; document explicitly if revisiting colour handling.
- **`sheetProtection` element bodies.** Secret attributes (`hashValue`,
  `saltValue`) are stripped but the element remains with non-secret
  attributes (`lockStructure`, `selectLockedCells`). On round-trip
  import Excel would treat these as "protected with no password" —
  possibly surprising. Consider a `protectionStripped: true` flag.
- **Table marker row format** in sheet `data.tsv` uses
  `[table:Name ref=R1:R2 -> tables/Name/]` in column A. If a sheet's
  column A starts with `[` for legitimate data, this could be visually
  confusing. Consider moving the marker to `_meta.json`.
- **`<extLst>` extension lists** in worksheets carry x14
  conditional-formatting and data-validation rules that the current
  parsers miss (they only look at the main namespace).
- **Static export templates risk the VBA 25-continuation limit.**
  `WriteGitIgnore` and `WriteReadme` build static text via long `& _`
  continuation chains. VBA caps line-continuations at 25 per logical
  statement; crossing it makes the VBE reject the whole module on Import
  with a bare `0x800A9D00` — and it compiles fine until it doesn't (the
  `.gitignore` builder hit exactly this on 2026-05-20). Proposed fix:
  move the static templates into an embedded payload part — same
  `CustomXMLPart` mechanism as the harness payload, but a separate part
  (`urn:vba-sync:core-templates:v1`), since `.gitignore`/README are
  vba-sync core export artifacts, not excel-control harness files. At
  emit time, load the template and token-substitute the dynamic bits
  (the `Format(Now …)` timestamp, the secret-file line). Removes the
  continuation hazard for good and makes the templates diffable text
  instead of escaped VBA concatenation.

## excel-control harness — deferred issues

**Re-verified against the current `excel-control/*.ps1` on 2026-05-21.**
An earlier read-only audit logged ~27 findings; each was re-checked
against the current code. Several were stale (line numbers predated the
Phase-0 rewrite of `start-session.ps1`) or wrong. Below is only what
still holds, with current line references. `D/E/F`, `H5/M1`, `M2`, `L3`
were already fixed and are not relisted.

**Dropped on re-verification — checked, do not re-add:** the prior
audit's `C1`/`C2` (`commands.jsonl` dual-consumer "lost commands" — each
reader keeps its own offset and scans the whole file independently, so
nothing is lost; the two-reader split is intentional, it lets the
watcher read commands while the main loop is blocked in a modal);
`C3` (the `WM_COMMAND → BM_CLICK → VK_RETURN → WM_CLOSE` fallback chain
already covers owner-drawn buttons; `Wait-Closed` tracks a specific
HWND); `H7`/`M9` (a closed dialog's HWND is dead — the
`dialog_appeared` / `runtime_error` event already carries the
screenshot, so retaining the id buys nothing); `M8` (`sync_vba` *does*
re-import `Objects/` — `VBComponents.Import` updates the matching
Document module in place); `L4` (the per-cell write now uses
`InvokeMember` per call, which removes the overload-cache hazard).

### High

- **Contract drift — `SCOPING.md` vs reality.** `SCOPING.md` documents
  a `FileSystemWatcher` command reader (`:49,204` — the code is a fixed
  `Start-Sleep $PollMs` poll loop, `start-session.ps1:829-843`), an
  `abort` command (`:91` — no such `switch` case), `respond_form`
  (`:60,83,174` — not implemented), and spontaneous `workbook_closed` /
  `status_changed` events (`:102-103` — never emitted, no COM event
  sinks). Implement them, or correct `SCOPING.md` — an agent following
  the doc sends commands that just return `unknown_command`.
- **Non-English Office breaks runtime-error handling.**
  `session-dialog-watcher.ps1` matches the English caption
  `"Microsoft Visual Basic"` (`:304`), the regex `Run-time error '…'`
  (`:307`), and the `&End` button caption (`:345`) literally. On a
  non-English Office the VBA runtime-error dialog isn't classified as
  `runtime_error`, isn't auto-End'd, and the blocked macro hangs the
  session. Match by window class `#32770` + layout; target End by
  control id.
- **`commands.jsonl` offset reset replays every command.** If the file
  is ever truncated/rewritten, the byte offset resets to 0 and every
  command re-executes (`start-session.ps1:118`,
  `session-dialog-watcher.ps1:217`). Dedupe by consumed command `id`.

### Medium

- **`read_range` / `write_range` are per-cell COM calls**
  (`start-session.ps1:419-428`, `470-480`). A large range blocks for a
  long time with no progress event. Size cap + clear error, or a
  heartbeat event.
- **`run_tests` reporting gaps** — the `errored` counter is always 0
  (`start-session.ps1:604,639`); `test_result` events carry no `id` to
  tie them to the `run_tests` command (`:623-631`).
- **`screenshot worksheet:<name>` in headless mode** captures the
  unrendered main window yet still reports `screenshot_captured`
  (`start-session.ps1:281-288`). Refuse or warn when `-Visible` is off.
- **`compile_check` timing** — a fixed 600ms `Start-Sleep` after
  `Execute()` then an `ActiveCodePane` check (`start-session.ps1:
  324-350`); too short for a large project, and `ActiveCodePane` is an
  imperfect failure signal. (The "already compiled" disabled-control
  case *is* handled.) Poll with a timeout; detect the compile-error
  `#32770` dialog.
- **`sync_vba` strip-then-import is not atomic** — the strip loop
  removes every user component before the import loop runs
  (`start-session.ps1:377-395`); an import failure mid-loop leaves the
  project stripped. Import-verify-swap, or report the partial state.
- **`state.json` written non-atomically** — `Set-Content` truncates
  then writes (`start-session.ps1:97-109`); a crash mid-write corrupts
  it. Write `.tmp` then `Move-Item -Force`.

### Low

- **No `#requires`** in `start-session.ps1`; `pwsh` vs `powershell.exe`
  is not pinned (the docs invoke `pwsh`, which isn't on every box).
- **`Resolve-Path` on `-Workbook`** runs before the `try` block
  (`start-session.ps1:743`) — a bad path throws a raw error with no
  `session_error` event and possibly no `state.json`.
- **Prototype scripts** (`harness.ps1`, `dialog-watcher.ps1`,
  `find-compile-error.ps1` — not in the shipped manifest) — hardcoded
  `C:\Users\ArnaudLavignolle\…` param defaults, PID-kill without a
  `ProcessName -eq 'EXCEL'` guard, and `Invoke-Expression`'d function
  bodies (`dialog-watcher.ps1:238-243`). Harden or delete;
  `dialog-watcher.ps1` is superseded by `session-dialog-watcher.ps1`.
- **Fixed 200ms delay** after `Activate` before a worksheet screenshot
  (`start-session.ps1:285`) — may be too short on a slow machine.

## excel-control SKILL.md — deferred doc-review items

From the `skill/SKILL.md` review (2026-05-20). 13 of 15 findings were
fixed in `a605f1a` / `b178c78`; these two are deferred — they need
`INTERFACE.md` edits, out of the SKILL.md-only review scope:

- **M4** — `SKILL.md` and `tools/INTERFACE.md` each maintain a separate
  Quickstart and both repeat the Debug.Print / UserForm / `.frx`
  limitations near-verbatim. They will drift. Split cleanly: SKILL.md =
  workflows + decisions; INTERFACE.md = command/event reference. Link,
  don't repeat.
- **L5** — the Debug.Print limitation is stated twice (`SKILL.md` and
  `INTERFACE.md`); pick one home (folds into M4).

## What not to do

- Don't write parsers for OOXML features without a workbook that
  exercises them. The repo has a cautionary tale here (see closed PR
  history).
- Don't add config flags for "old layout vs new layout". Hard cutover.
- Don't add a build system. Still VBA + a minimal PowerShell harness.
- Don't start by removing things you don't understand. The current
  baseline works; treat it with respect.
- Don't propose 5 changes in one PR. Closed PRs in this repo's history
  show that pattern fails.

## Test workbook shapes that exercise edge cases

When picking workbooks to validate a change, look for ones that combine
these properties (test names anonymised):

- Many sheets (25+) with tables and LAMBDA defined names; sheet
  protections enabled on most
- A single very large sheet (~20MB of cell data); many sheet
  protections
- Workbook-open password set; multiple protected sheets; embedded
  drawings; shape-assigned macros

Don't hardcode workbook paths into test scripts — accept a `-Target`
parameter and document the shape it should exercise.
