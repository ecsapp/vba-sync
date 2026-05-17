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
