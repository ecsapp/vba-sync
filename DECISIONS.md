# Decisions

Design decisions for vba-sync's Excel export. These were debated and
locked. Don't reopen unless you have new information.

## Output layout

- **Tables nest under their host worksheet folder**, not at the Excel/
  root. Same for pivots, charts, drawings.
- **Sheet folders are `<NN> - <SheetName>/`** — sheetId-prefix then
  name.
- **`data.tsv`** holds cell values, formula cached values included.
  Formulas themselves live in `formulas.json`.
- **Table data is excluded from its host sheet's `data.tsv`** (it lives
  in the table's own `data.tsv`). A marker row at the table's start
  in the sheet `data.tsv` preserves positional context.
- **Table calc-column formulas live in `tables/<Name>/definition.json`**
  — never per-row inside the table's `data.tsv`.
- **Formulas in A1 form, range-collapsed.** R1C1 was rejected because
  translating it requires a real formula parser.
- **Drawings are first-class.** Buttons + macros are essential UI
  context. Skipped: pixel positions, rotations, gradients, SmartArt.
- **Styles emit only non-default attributes.** Don't write 40 lines per
  range when only `fill` differs.
- **Range merging uses greedy rectangles.** Disjoint sets get split
  into multiple entries.

## Scope

- **Round-trip import is out of scope for v1.x.** Deferred — see
  BACKLOG.
- **The .xlsm itself stays committed alongside the export folder.**
  Excel binary in git is the contract.
- **Hard cutover at v1.0.** No "old layout vs new layout" config flag.
  Migration commits are the documentation.
- **Password hashes are silently stripped.** Not configurable.

## Encoding / files

- **UTF-8 export** (Fix #1, merged via PR #3). Auto-generates
  `.vscode/settings.json` with the system codepage hint so VS Code
  opens files at the right encoding.
- **One PR per concern.** Closed-PR history shows that 5-changes-in-one
  PRs stall and get reverted.
