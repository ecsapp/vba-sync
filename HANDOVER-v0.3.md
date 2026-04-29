# Handover — vba-sync v0.3 Excel export rewrite

A briefing for the next agent (or human contributor) picking up the second half of the Excel-export rewrite. **Read this before touching anything in `Normalize-ExcelXml.ps1` or `modSync.bas`'s `ExtractExcelStructure`.**

---

## What v0.2 already shipped

The current branch (`excel-export-rewrite`) replaces the old "raw single-line XML dump" with:

- **Stable, semantic filenames.** `worksheets/02 - Project Settings.xml` instead of `worksheets/sheet5.xml`. Tables named after the table (`HardLogic.xml`) instead of the rId (`table10.xml`).
- **Multi-line, indented XML** with attributes alphabetised per-element. Diff-friendly.
- **Volatile metadata stripped.** All revision GUIDs, window positions, scroll positions, calc build IDs, the local filesystem path Excel embeds (privacy leak), and Excel-build feature advertisements.
- **Password hashes redacted** from `<fileSharing>`, `<workbookProtection>`, and per-sheet `<sheetProtection>` elements. Logged so the user knows.
- **LAMBDA dedup.** A LAMBDA defined name that gets copied per-sheet by Excel collapses into one `lambdas/<Name>.lambda` file plus a single reference in MANIFEST.
- **MANIFEST.json** replaces the monolithic `workbook.xml`. Sheet list, defined names, lambdas, calc settings, all in structured JSON.

Architecture: a PowerShell 5.1 script (`Normalize-ExcelXml.ps1`, ships next to the .xlam) does the XML work via `[System.Xml.Linq.XDocument]`. The VBA in `ExtractExcelStructure` orchestrates: copy .xlsm → temp .zip, expand, shell the normaliser, walk output, mark exported. If the script is missing, falls back to the legacy raw-copy path so an export still produces *something*.

**What v0.2 doesn't do:** any of the per-worksheet split into `data.tsv` / `formulas.json` / `styles.json` / etc. The worksheet XML files are still raw normalised OOXML, just much cleaner. Tables are still XML inside `tables/` at the workbook root, not under their host sheet.

---

## What v0.3 needs to do

The full layout that was specified in the design conversation but deliberately deferred for incremental shipping. Open `excel-export-rewrite` PR description on GitHub for the full design discussion if it has been opened by the time you read this.

### Target layout

```
<workbook>/Excel/
├── MANIFEST.json
├── .vba-sync/config.json                    # NEW: per-repo export options
├── lambdas/
│   └── *.lambda
├── worksheets/
│   └── 36 - Project Settings/               # Folder per sheet, not flat XML
│       ├── _meta.json                       # Tab colour, frozen panes, col widths,
│       │                                    # hosted-tables ref, hidden state
│       ├── data.tsv                         # Cached values for non-table cells,
│       │                                    # with table-marker rows at table
│       │                                    # positions (Option A from spec)
│       ├── formulas.json                    # R1C1, range-collapsed
│       ├── styles.json                      # Resolved dxfId, range-merged
│       ├── validations.json
│       ├── conditional_formats.json
│       ├── comments.json
│       ├── tables/
│       │   └── HardLogic/
│       │       ├── definition.json          # Schema, formulas, overrides
│       │       └── data.tsv                 # Clean dataset
│       ├── pivot_tables/
│       │   └── PivotTable1/
│       │       ├── definition.json
│       │       └── data.tsv
│       ├── charts/
│       │   ├── Chart 1.json
│       │   └── Chart 1.png
│       └── drawings/
│           ├── shapes.json                  # All shapes: name, type, anchor,
│           │                                # text, macro (OnAction), visibility
│           └── _assets/
│               └── *.png
└── STRUCTURE_SUMMARY.md
```

### Migration policy: hard cutover

When v0.3 ships, the export wipes the old `Excel/` folder layout and writes the new one. Users review the migration diff in their own repos, commit it once, move on. **Do not implement a config flag for layout v1 vs v2.** It's a maintenance trap.

`PruneStaleFiles` in modSync.bas is already recursive under `Excel/`, so it will naturally clean up stale v0.2 files when v0.3 writes new ones — no special migration code needed.

---

## Decisions already made (don't relitigate)

These came out of the design conversation. The agent should not re-open them without strong reason.

1. **Tables nest under their host worksheet folder** — not at the Excel root. Same for pivot tables, charts, drawings. Only `lambdas/` and `MANIFEST.json` live at the Excel root because they're cross-sheet.

2. **Sheet folders are `<NN> - <SheetName>/`** with the same sheetId-prefix-then-name convention v0.2 uses for files. Renames stay rename-stable; sheet ordering is preserved by sheetId pad.

3. **`data.tsv` includes formula cell *values* (cached calc results) plus literal cell values.** Formulas live separately in `formulas.json`. Two questions, two files. Power users can opt out of the cached-results write via `.vba-sync/config.json` to avoid recalc-only churn.

4. **Table data is *excluded* from the host sheet's `data.tsv`** (lives in the table's own `data.tsv` instead). The sheet's data.tsv has a marker row at the table's first row (Option A from the spec): `[table:HardLogic ref=A8:I108 — see tables/HardLogic/]`. This is *not* a comment header; it's a real row in the TSV that occupies the table's range so positional context is preserved.

5. **Formulas: R1C1 format, range-collapsed** when adjacent cells share an R1C1 expression. Schema:
   ```json
   {
     "ranges": { "B2:B100": "=XLOOKUP(RC1, ...)", "C2:C100": "=RC2*RC[-1]" },
     "cells":  { "F1": "=SUM(F2:F100)" }
   }
   ```
   Empty `ranges` or `cells` blocks are omitted.

6. **Table calculated-column formulas live in `tables/<Name>/definition.json`** — never repeated per row in the table's `data.tsv`, never duplicated in the host sheet's `formulas.json`. The definition.json includes an `overrides` block keyed by table-row-index for cells that have been hand-edited away from the column's default formula.

7. **Drawings are first-class.** Buttons that trigger macros are essential to understanding UI flow. Schema captures shape name, type (`button`/`rectangle`/`textBox`/`image`/`group`/`formControl`/`activeXControl`), anchor in cell terms, text content, **OnAction macro reference**, group hierarchy. Images stored as PNG/JPG in `drawings/_assets/`. Skipping: pixel positions, rotations, gradients, SmartArt internals, embedded OLE.

8. **Charts: PNG snapshot + JSON manifest of structural data** (data ranges, series, axes, title, chart type). PNG diff is meaningless but tells you *that* it changed; JSON diff tells you *what data* changed.

9. **Style resolution: only emit non-default attributes.** Don't write 40 lines per range when only `fill` differs. `dxfId` indirection is resolved (e.g. `dxfId=137` becomes `{ "fill": "#FFE699" }`).

10. **Range merging: greedy rectangles.** Disjoint sets like `A1:A10,C1:C10` get split into two entries, not merged. More entries, but each is a clear range.

11. **Round-trip import is OUT OF SCOPE for v0.3.** That's v0.4+. The user explicitly said "would be awesome though, especially if you can get control to import/export and run vba yourself. but let's leave for later". Don't try to implement import in v0.3.

12. **`xlsm` itself stays committed alongside the export folder for v0.3.** Excel binary in git, just like v0.2. Re-evaluate once round-trip lands.

---

## Open questions for the next agent

When you start work, post your responses to these to the user before implementing:

1. **PowerShell vs VBA for the new work.** The v0.2 normaliser is PowerShell because it has good XML/JSON tools. v0.3 needs to read live Excel state for some things (drawings/shape OnAction, pivot table sources, chart series) that are easier from VBA via the Excel object model. Pick one of: (a) VBA does everything new, PS handles only what it already does; (b) extend PS to parse OOXML drawings/charts; (c) split — VBA gathers, writes JSON to disk, PS still handles the workbook-level XML normalisation. My read: **(a)**, because v0.3's value is the schema, not the parsing pyrotechnics. Confirm with user.

2. **TSV format for data.tsv with `Row | A | B | C ...` index column** — already in spec, but worth flagging. Cells with tabs/newlines: the user said vote was TSV (assume no tabs/newlines in cells). When you encounter one, what to do — escape, skip, fall back to JSON? Likely escape with `\t`/`\n` and document. Get user sign-off.

3. **Drawing schema** — the v0.3 spec is sketched in the design doc but not finalised. Confirm with the user before building the JSON schema. Real shapes from his repos will inform it; sample 3-4 of his workbooks first.

4. **Chart PNG generation.** Excel's `Chart.Export` method writes PNG/JPG/GIF. Path of least resistance. But it requires the chart to render fully — some charts depend on data that comes from external connections. Test on at least one workbook with a complex chart before committing to PNG-as-source-of-truth.

5. **Column-width preservation in `_meta.json`.** Excel default is 8.43. Most user sheets have custom widths. Per-column or per-range? Per-column is simpler; per-range with run-length encoding is smaller. Lean per-column.

6. **What about pivot table source data?** A pivot reading from an internal table is fine — record the source table name. A pivot reading from an external connection is in `xl/connections.xml` and `xl/pivotCaches/`. v0.2 doesn't touch these. v0.3 needs to at least preserve enough info to recreate the pivot: source ref, fields-on-rows/cols/values, filters. Out of scope: cached pivot data (will regenerate from source anyway).

---

## Existing code you'll touch

Files in `c:/Users/ArnaudLavignolle/AppData/Roaming/Microsoft/AddIns/`:

- **`Normalize-ExcelXml.ps1`** — extend or fork. The current script writes worksheet XML to `worksheets/<NN> - <Name>.xml`. v0.3 needs to instead create the folder + sub-files. One option: rename the v0.2 output to `worksheets/<NN> - <Name>/_raw.xml` as an intermediate, then have a second pass split it. Cleaner option: rewrite the worksheet path to do the split from the start.

- **`VBA Sync/Modules/modSync.bas`** — `ExtractExcelStructure` (~line 220) is the orchestrator. The shell call to PS, the legacy fallback, and the post-processing all live there. `CreateExcelStructureSummary` (~line 395) needs an update to point at the new layout. `PruneStaleFiles` (~line 567) is already recursive, no change needed.

- **`README.md`** — "What Gets Exported" tree (~line 24) needs the new tree.

- **`HANDOVER-v0.3.md`** — this file. Update or delete once v0.3 ships.

---

## Test workbooks

The user's real workbooks are the only test fixtures that matter. They cover edge cases no synthetic test will:

| Path | What it stresses |
|---|---|
| `~/Dev/p6-excel-tools/P6 Tools.xlsm` | 25 worksheets, 26 tables, 15 LAMBDA copies, sheet protections, complex defined names |
| `~/Dev/Axiom Forecast/Axiom Forecast.xlsm` | Large data sheets (22MB worksheet), many sheet protections |
| `~/Dev/gec-quote-system/Quote System.xlsm` | Workbook open password, structure protection, 13 protected sheets |
| `~/Dev/zinfra-mcr/MCR Template.xlsm` | (was empty in old export — verify it works under v0.2 first, then v0.3) |

Always test against at least 3 of these before declaring v0.3 done. The standalone PS test pattern from v0.2 is documented in this branch's commit history and worth replicating.

---

## What not to do

- **Don't add config flags for "v1 layout vs v2 layout vs v3 layout".** Hard cutover. Migration commits are the documentation.
- **Don't merge `Arrays.cls`-style heroics** into vba-sync. Keep DataModel concerns separate.
- **Don't try to round-trip in v0.3.** Explicitly out of scope.
- **Don't add a build system.** This is still VBA + a single PowerShell script.
- **Don't commit cached pivot data, calc chain XML, shared strings table, or styles.xml** to the export. Those are derived.
- **Don't auto-write `.vba-sync/config.json` with sensitive defaults.** Created on first export with conservative defaults; user opts into anything privacy-affecting.

---

## Provenance

This handover written by an AI agent that just shipped v0.2 on this branch. Earlier in the same session, the same agent built the `vscode-encoding-hint` PR (#3). Earlier still, the abandoned `utf8-encoding` PR (#2) — its closure is a useful cautionary tale about agents that try to fix too much in one PR. **If you find yourself proposing 5 changes for v0.3, you're doing it wrong.** Pick one of the 6 sub-features (worksheet split / tables-under-sheet / drawings / charts / pivots / config) and ship it standalone. The hard cutover lets you ship them one at a time without worrying about half-states.
