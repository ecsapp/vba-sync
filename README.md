# VBA Sync

An Excel add-in that exports your VBA project and workbook structure to a clean folder hierarchy, so Excel applications can be version-controlled, code-reviewed, and worked on with the same tools as any other software project.

## Why

Excel files are binary blobs. That makes the VBA code, table schemas, formulas, and worksheet structure inside them opaque to Git, to code review, and to AI tools that read source. VBA Sync exports everything to plain-text files arranged by concern, so:

- **Git** sees real diffs instead of "binary file changed"
- **Code review** works the same as any other repo
- **AI assistants** can read both your VBA and your data model
- **Teams** can branch, review, and merge Excel applications like software

## Quick start

1. Download `VBA Sync.xlam`, copy to `%APPDATA%\Microsoft\AddIns\`, enable it in Excel.
2. Open your workbook from a local or synced folder (not a SharePoint URL).
3. Click **VBA Sync → Export**. A folder named after your workbook is created next to it.
4. Edit code in VS Code, commit to Git, get AI assistance.
5. Click **VBA Sync → Import** to load VBA changes back into Excel.

## What gets exported

```
YourWorkbook/
├── Modules/                            # Standard VBA modules (.bas)
├── ClassModules/                       # VBA class modules (.cls)
├── Forms/                              # UserForms (.frm + .frx)
├── Objects/                            # ThisWorkbook & Sheet modules (.cls)
└── Excel/
    ├── MANIFEST.json                   # Workbook structure: sheets, tables, defined names, lambdas, calc
    ├── lambdas/                        # LAMBDA defined names, one .lambda file per unique body
    │   └── AppendRange.lambda
    ├── worksheets/
    │   └── <NN> - <SheetName>/         # One folder per sheet (sheetId-prefixed for stable sort)
    │       ├── data.tsv                # Cell values, with shared strings resolved inline
    │       ├── formulas.json           # Cell formulas, range-collapsed (B2:B100 = "=...")
    │       ├── styles.json             # Resolved cell styles (fill/font/border/numFmt), range-merged
    │       ├── _meta.json              # Tab colour, frozen panes, column widths, hosted tables
    │       ├── validations.json        # Data validation rules (if any)
    │       ├── conditional_formats.json
    │       ├── comments.json           # Cell comments + author
    │       ├── tables/<TableName>/
    │       │   ├── definition.json     # Schema, columns, calc formulas, overrides
    │       │   └── data.tsv            # Clean dataset (headers row 1, data rows below)
    │       └── drawings/
    │           ├── shapes.json         # Shapes, pictures, with OnAction macro names
    │           └── _assets/            # Embedded images
    └── STRUCTURE_SUMMARY.md            # Human-readable overview
```

The `Excel/` folder is produced by reading the live Excel object model from VBA:

- Splits each worksheet into per-concern files so a data edit doesn't churn the styles file (and vice versa).
- Range-merges adjacent cells with identical styles or formulas (a calc column with one formula across 700 rows is one entry, not 700).
- Captures shape macros (`OnAction`) so AI agents can trace UI → VBA without opening Excel.
- Deduplicates LAMBDA copies (Excel writes one per sheet that uses it; we collapse to one file).
- Skips volatile metadata that doesn't survive a no-change save.

The result is small, readable diffs even after a "save with no changes". A column added to a table touches one or two files; the raw OOXML would have touched dozens.

The export also writes Git configuration:

- `.gitattributes` — line endings and encoding rules per file type
- `.gitignore` — excludes Excel temp files and system cruft
- `README.md` — created once, never overwritten

## Installation

1. Download `VBA Sync.xlam` from this repository.
2. **Unblock the file** (Windows mark-of-the-web):
   - Right-click `VBA Sync.xlam` → **Properties**
   - Tick **Unblock** at the bottom, click **OK**
3. Double-click the file. Excel will prompt to install it.
4. Click **Enable** when Excel asks about the add-in.
5. The **VBA Sync** tab appears in the ribbon.

If the double-click flow doesn't work:

- Copy `VBA Sync.xlam` to `%APPDATA%\Microsoft\AddIns\`
- Excel → **File** → **Options** → **Add-ins** → **Excel Add-ins** → **Browse** → select the file

Excel must be allowed to read its own VBA project:

- **File** → **Options** → **Trust Center** → **Trust Center Settings**
- **Macro Settings** → tick **Trust access to the VBA project object model**

## Use cases

- **Financial models** — version control formulas, table schemas, and VBA business logic
- **Reporting tools** — track changes to data pipelines and report generation
- **Dashboards** — collaborate on interactive Excel apps with PR-based workflows
- **Data integration** — manage API connections, queries, and ETL in Git
- **Automation** — full change history on Excel automation scripts

## Notes and limitations

- **Local files only.** Open workbooks from a local or synced folder, not from a SharePoint URL.
- **Import is VBA-only.** The `Excel/` structure export is for versioning and review. Import only updates VBA modules; it does not rebuild worksheets, tables, or formulas from the exported files.
- **Smart filtering.** Empty document modules (sheets/`ThisWorkbook`) are skipped; removed components are cleaned up on the next export.
- **Macro security.** Settings must allow the add-in to run.
- **Back up first.** Save and back up your workbook before the first export.

## License

MIT — see [LICENSE](LICENSE).

## Credits

Created by **Arnaud Lavignolle**
(with help from Claude and ChatGPT)

Built to bridge the gap between Excel development and modern software engineering practices.
