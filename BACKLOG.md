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

## excel-control harness — deferred audit findings

From a full read-only audit of `excel-control/` (2026-05-20). The
harness-self-sufficiency effort (payload embedding + unlock wiring) folded
in only `H5/M1`, `M2`, `L3`, `L5`; bugs `D/E/F` were already fixed.
Everything below was **deliberately deferred** to keep that effort focused.
Severity and `file:line` are from the audit at time of writing — verify
against current code before acting.

### Critical — `commands.jsonl` dual-consumer design

- **C1** — `commands.jsonl` consumed by two readers (main loop + dialog
  watcher) with independent byte offsets, no coordination
  (`start-session.ps1:100-127`, `session-dialog-watcher.ps1:212-235`).
  `respond_dialog` delivery is timing-dependent. Fix: single consumer —
  main loop reads all, hands `respond_dialog` to the watcher via a
  `ConcurrentQueue`; or a dedicated `dialog-commands.jsonl`.
- **C2** — `command_ack` for `respond_dialog` emitted cross-thread, no
  ordering guarantee vs `dialog_appeared`. Fixed by the C1 redesign.
- **C3** — MsgBox blocks the COM thread; `Wait-Closed` races a
  replacement modal; owner-drawn buttons return ctrl-id 0 and `WM_COMMAND`
  with a standard id is never tried (`session-dialog-watcher.ps1:237-266`).
  Fix: map button label → standard control id (OK=1, Cancel=2, Yes=6,
  No=7); never free a dialog id until `IsWindow` is false.

### High — contract drift + robustness

- **H1** — `FileSystemWatcher` promised (`SCOPING.md:50-52,204-205`); code
  is a fixed 250ms poll loop. Implement it or correct the docs.
- **H2** — `abort` command documented (`SCOPING.md:91`) but not
  implemented; no way to interrupt a running macro. Implement or strike.
- **H3** — `respond_form` documented (`SCOPING.md:60,83,174`) but not
  implemented; three docs disagree. Strike or mark deferred.
- **H4** — spontaneous events `workbook_closed` / `status_changed` never
  emitted (`SCOPING.md:97-104`); no COM event sinks wired. Wire
  `Register-ObjectEvent` sinks or remove the rows.
- **H6** — offset-reset-on-shrink silently replays every command if
  `commands.jsonl` is truncated/rewritten (`start-session.ps1:105`). Fix:
  dedupe by consumed command `id` persisted in `state.json`.
- **H7** — `runtime_error` dialog auto-End'd and the id freed before the
  agent can re-screenshot (`session-dialog-watcher.ps1:343-348`). Keep
  `DialogInfo` for a grace TTL after auto-End.
- **H8** — `&End` button label and `'Microsoft Visual Basic'` title match
  are English-only (`session-dialog-watcher.ps1:304,345`); non-English
  Office hangs. Match by window class `#32770` + standard control id.

### Medium

- **M3** — `read_range`/`write_range` are per-cell COM calls; large ranges
  block for minutes with no progress event (`start-session.ps1:360-366,
  408-418`). Size cap + clear error, batched `Value2`, or a heartbeat.
- **M4** — `run_tests` `errored` counter always 0; `test_result` has no
  `id` (`start-session.ps1:537,563,572`).
- **M5** — `screenshot worksheet:<name>` renders blank in headless mode
  but reports success (`start-session.ps1:233-240`).
- **M6** — `compile_check` uses a fixed 600ms sleep + `ActiveCodePane`
  heuristic → false results (`start-session.ps1:269-276`). Detect the
  compile-error `#32770` dialog; poll with timeout.
- **M7** — `sync_vba` strip-then-import is not atomic; a mid-loop import
  failure leaves the project gutted (`start-session.ps1:316-333`).
- **M8** — `sync_vba` cannot update `Document` modules (`ThisWorkbook`,
  sheet code); type-100 skipped (`start-session.ps1:317-318`). For
  `Objects/`, replace `CodeModule` contents line-by-line (the technique
  `modSync.bas` import already uses).
- **M9** — dialog screenshot after close throws `Unknown dialog/form id`
  (`session-dialog-watcher.ps1:284-291`). Retain info for a TTL.
- **M10** — runtime-error number regex English-only
  (`session-dialog-watcher.ps1:307`).
- **M11** — prototype scripts kill by PID with weak guards
  (`harness.ps1:210-212`, `find-compile-error.ps1:104`). Assert
  `ProcessName -eq 'EXCEL'` before `Stop-Process`.
- **M12** — `state.json` written non-atomically (`start-session.ps1:
  84-96`). Write `.tmp` then `Move-Item -Force`.
- **M13** — fixed 200ms delay after `Activate` before screenshot
  (`start-session.ps1:236-239`). Minor.

### Low

- **L1** — `pwsh` vs `powershell.exe` not pinned; no `#requires`.
- **L2** — `harness.ps1:31-32`, `find-compile-error.ps1:10-11` hardcode
  `C:\Users\ArnaudLavignolle\...` defaults.
- **L4** — `write_range` numeric coercion misses `[int32]`
  (`start-session.ps1:412`).
- **L6** — `dialog-watcher.ps1:238-243` `Invoke-Expression`s function
  bodies into the runspace; superseded by `session-dialog-watcher.ps1` —
  delete once `harness.ps1` / `find-compile-error.ps1` migrate.
- **L7** — `start-session.ps1:676` `Resolve-Path` on `-Workbook` runs
  before the `try` block; a bad path throws a raw stack trace with no
  `session_error` event. Validate + emit a structured failure first.

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
