# Handover — vba-sync v1.1+ Excel export

A briefing for the next agent (or human contributor) picking up the next round of vba-sync Excel-export work. Read this before touching `Normalize-ExcelXml.ps1`.

---

## What v1.0 ships (current state)

The Excel export was rewritten end-to-end on the `excel-export-rewrite` branch. Per-sheet folder layout:

```
Excel/
├── MANIFEST.json                              schemaVersion 1, generatorVersion "1.0"
├── lambdas/<Name>.lambda                      Deduplicated LAMBDA defined names
└── worksheets/<NN> - <SheetName>/
    ├── data.tsv                               Cell values (sst-resolved)
    ├── formulas.json                          {ranges:{...}, cells:{...}, sharedFormulas:{...}}
    ├── styles.json                            {range: {fill, font, border, numFmt}, ...}
    ├── _meta.json                             tabColor, frozenPanes, columns, mergedCells, hosted-tables
    ├── validations.json                       (only if any rules)
    ├── conditional_formats.json               (only if any rules)
    ├── comments.json                          (only if any comments)
    ├── tables/<TableName>/
    │   ├── definition.json                    Schema, columns, calc formulas, overrides
    │   └── data.tsv                           Clean dataset
    └── drawings/
        ├── shapes.json                        Shapes, pictures, **with OnAction macro names**
        └── _assets/                           Embedded images
```

The cell range a table occupies is removed from the host sheet's `data.tsv`; a marker row (`[table:Foo ref=A1:I100 -> tables/Foo/]`) marks where it sat. This preserves positional context without duplicating cells.

Ships with `Normalize-ExcelXml.ps1` next to the .xlam. PowerShell 5.1 compatible (System.Xml.Linq + ConvertTo-Json + ADODB.Stream). VBA wiring in `modSync.bas` ExtractExcelStructure: copy → expand → shell PS → mark exported → cleanup. **No legacy fallback**: if the PS script is missing, export errors out clearly rather than producing v0.1-style raw XML.

Test workbooks all clean:

| Workbook | v0.2 size | v1.0 size |
|---|---|---|
| Axiom Forecast | 25.2 MB | ~5 MB |
| P6 Tools | 12.8 MB | ~5 MB |
| Quote System | 2.3 MB | ~2 MB |

---

## What v1.1+ should add

### 1. Charts (D2 deferred from v1.0)

None of v1.0's test workbooks have charts, so writing the parser blind is exactly the kind of mistake the project's history warns against (see PR #2 cautionary tale). Hold for a workbook with real charts to test against.

When implementing:

- **JSON manifest** from `xl/charts/chart*.xml`: chart type (bar/line/pie/etc.), data ranges, series names, axis labels, title. Resolved via the worksheet's `_rels` to `xl/drawings/drawing*.xml` to get the chart's anchor on the host sheet.
- **PNG snapshot** is the visual source of truth, but `Chart.Export()` requires Excel COM. Two options:
  - Add a VBA pre-step in `ExtractExcelStructure` that opens the workbook, calls `Charts.Export` for each chart into a temp PNG, then shells PS as today. PS picks up the PNGs from a known location.
  - Give up on PNG and rely on JSON-only. Diff signal is "the data ranges changed", not "the chart shape changed". Pragmatic but loses something.
- Output: `worksheets/<NN> - <Name>/charts/<ChartName>.json` plus optionally `<ChartName>.png`.

### 2. Pivot tables (D3 deferred from v1.0)

Same reason for deferral. None of the test workbooks have pivots.

When implementing:

- Parse `xl/pivotTables/pivotTable*.xml` for source ref, fields-on-rows/cols/values, filters
- Output: `worksheets/<NN> - <Name>/pivot_tables/<PivotName>/{definition.json, data.tsv}`
- The `data.tsv` is a snapshot of the rendered pivot (the cached values), useful for diff signal even if the underlying data hasn't changed
- Cached pivot data in `xl/pivotCache/` is opaque-and-large; skip it (Excel regenerates from source)

### 3. .vba-sync/config.json

Per-repo export options. Created on first export, never overwritten. Schema sketch:

```json
{
  "$schema": "https://...",
  "export": {
    "stripVolatileMetadata": true,
    "includeFormulaCachedValues": false,
    "includeRevisionUIDs": false,
    "includeViewState": false,
    "includeLastModifiedBy": true,
    "data": {
      "defaultMaxRows": null,
      "perSheetOverrides": {
        "Raw P6 Dump": { "maxRows": 500 }
      }
    },
    "drawings": {
      "copyAssets": true
    }
  }
}
```

Currently everything's hardcoded sensible defaults. Add this when a user first asks for an option that should be per-repo (e.g. "skip cached formula values to reduce churn").

### 4. Round-trip import (v0.4-ish era from old plan, now v1.2+)

Major undertaking. The user's preferred angle: don't write OOXML from scratch — instead drive Excel via VBA + the object model. For each text file we wrote, apply it back to a live workbook:

- `data.tsv` → `Range.Value2 = ...`
- `formulas.json` → `Range.Formula = ...`
- `styles.json` → walk and apply to ranges
- `_meta.json` → `ws.Tab.Color = ...`, freeze panes, etc.
- `tables/*/definition.json` → `ws.ListObjects.Add` then column setup
- `drawings/shapes.json` → `ws.Shapes.AddShape` with OnAction assignment
- `drawings/_assets/*.png` → `ws.Pictures.Insert`
- `validations.json`, `conditional_formats.json`, `comments.json` → respective Excel object model APIs

The VBA side becomes the importer. PS is still used for export. The importer can be incremental (just modules first, like vba-sync's existing VBA import) and add features one at a time.

The killer subset: **import the data.tsv changes only** (cell values + formulas). Most of the time that's what someone editing a workbook in VS Code wants — they touched a value, want it back in Excel. Style changes, drawing changes, etc. can be later.

---

## Known v1.0 gaps to fix in v1.1

- **`tabColor` only handles `rgb=`**. Indexed (`indexed=`) and theme (`theme=N`) colours silently become `null` in MANIFEST. Need theme1.xml lookup + indexed-palette table. ECMA-376 Part 1 has both lookup tables.

- **`Get-SafeFileName` truncates at 100 chars without uniqueness check.** Two long sheet names sharing a 100-char prefix would produce identical filenames. Vanishingly unlikely (Excel limits sheet names to 31 chars) but a real correctness gap.

- **`tabColor` capture strips alpha.** `Substring(2)` drops the `AA` byte from `AARRGGBB`. Tabs don't support transparency so this is fine in practice; document it explicitly when reworking colour handling.

- **Stale `<workbookProtection>`/`<sheetProtection>` element bodies** remain after hash strip. We strip the secret attributes (hashValue, saltValue, etc.) but leave the element with its non-secret attributes (lockStructure, selectLockedCells). On round-trip import, Excel will treat these as "protected with no password" — possibly surprising. Consider documenting in MANIFEST or adding a "protectionStripped: true" flag.

- **The marker row in sheet `data.tsv` for tables** uses a fixed format `[table:Name ref=R1:R2 -> tables/Name/]` in column A. If a sheet has a table whose first column actually contains data starting with `[`, this could be visually confusing. Consider a non-printable prefix or moving the marker to `_meta.json` instead.

- **`Format-Json` collapses `{}` and `[]` correctly** but other empty-container patterns (an empty `definedNames` object inside a non-empty manifest) might still emit a blank line. Audit if it bites.

- **No explicit handling of `<extLst>` (extension lists)** in worksheets. These can carry x14 conditional-formatting and data-validation rules that v1.0 misses. The validations.json and conditional_formats.json parsers only look at the main namespace. Worth checking how often Arnaud's workbooks use x14 rules.

- **The PS5.1-vs-PS7 boundary**: if v1.1 ever needs `ForEach-Object -Parallel` for performance, that's PS7-only. v1.0 sticks to 5.1 for corporate-machine compatibility. Re-evaluate when the user complains about export time.

---

## Decisions already made (don't relitigate)

These were debated extensively in the design conversation. Don't reopen unless you have new information.

1. **Tables nest under host worksheet folder**, not at Excel root. Same for pivots, charts, drawings.

2. **Sheet folders are `<NN> - <SheetName>/`** with sheetId-prefix-then-name convention.

3. **`data.tsv` includes formula cached values** plus literal cell values. Formulas live separately in `formulas.json`.

4. **Table data is excluded from host sheet's `data.tsv`** (lives in the table's own data.tsv). The sheet's data.tsv has a marker row at the table's start.

5. **Formulas in A1 form, range-collapsed**. R1C1 was considered and rejected — we'd need a real formula parser to translate, which is non-trivial.

6. **Table calc-column formulas live in `tables/<Name>/definition.json`** — never per-row in the table's data.tsv.

7. **Drawings are first-class.** Buttons + macros are essential UI context. Skipped: pixel positions, rotations, gradients, SmartArt.

8. **Style resolution: only emit non-default attributes.** Don't write 40 lines per range when only `fill` differs.

9. **Range merging: greedy rectangles.** Disjoint sets get split into multiple entries.

10. **Round-trip import is OUT OF SCOPE for v1.0.** Deferred to v1.2+ via the VBA-driven approach.

11. **`xlsm` itself stays committed alongside the export folder.** Excel binary in git.

12. **Hard cutover at v1.0.** No "layout v0.x vs layout v1.0" config flag. Migration commits are the documentation.

13. **PowerShell 5.1, no module installs.** Compatible with locked-down corporate machines.

14. **Password hashes are silently stripped** with summary in STRUCTURE_SUMMARY.md. Not configurable.

---

## Existing code you'll touch

Files in `c:/Users/ArnaudLavignolle/AppData/Roaming/Microsoft/AddIns/`:

- **`Normalize-ExcelXml.ps1`** (~2000 lines) — main script. Process-Worksheet is the entry per sheet. Save-* functions one per output file. Helpers near the top: A1 ref math, sharedStrings/styles loaders, range merger.

- **`VBA Sync/Modules/modSync.bas`** — `ExtractExcelStructure` (~line 220) is the orchestrator. Mostly stable; v1.1 may need a Chart.Export pre-step here.

- **`README.md`** — "What Gets Exported" tree should be updated for any new output files.

- **`HANDOVER-v1.1.md`** — this file. Update or delete as work lands.

---

## Test workbooks

| Path | Stresses |
|---|---|
| `~/Dev/p6-excel-tools/P6 Tools.xlsm` | 25 sheets, 26 tables, 15 LAMBDA copies, sheet protections |
| `~/Dev/Axiom Forecast/Axiom Forecast.xlsm` | 22MB Timesheets sheet, many sheet protections |
| `~/Dev/gec-quote-system/Quote System.xlsm` | Workbook open password, 13 protected sheets, 17 drawings, **shape macros** |

For chart/pivot work: find an Arnaud workbook that has them. Don't write parsers without test data.

---

## What not to do

- **Don't add config flags for "v1.0 vs v1.1 layout".** Hard cutover.
- **Don't write parsers for OOXML features without a test workbook that exercises them.** This bit us before.
- **Don't add a build system.** Still VBA + a single PowerShell script.
- **Don't try to round-trip in v1.1.** Explicitly v1.2+.
- **Don't start by removing things you don't understand.** v1.0 is the working baseline; treat it with respect.

---

## Provenance

This handover written 2026-04-30 by an AI agent at the end of the v1.0 implementation session. Earlier in the session: Phase A (per-sheet folder + data/formulas/styles), Phase B (tables under host), Phase C (validations/CF/comments/_meta), Phase D1 (drawings + macros). Phases D2 (charts) and D3 (pivots) were honestly deferred because the test workbooks don't exercise them.

Earlier in vba-sync history: PR #2 (UTF-8 re-encode attempt that broke too much, closed) and PR #3 (vscode-encoding-hint, the lightweight fix) — both useful cautionary tales about over-scoping. If you find yourself proposing 5 changes in one PR, you're doing it wrong.
