# Excel VBA Sync - Version Control for Excel Applications

🚀 **Bring modern software development practices to Excel VBA!**

This Excel add-in automatically exports your VBA project AND Excel file structure to a clean folder hierarchy, enabling Git version control, AI assistance, team collaboration, and professional development workflows for Excel-based applications.

## 🎯 Key Benefits

- 🤖 **AI Collaboration**: AI tools can understand both your VBA code AND Excel data models
- 📚 **Version Control**: Full Git history of code changes, table schemas, and worksheet structure
- 👥 **Team Development**: Review changes, manage pull requests, and collaborate like software teams
- 🔍 **Code Intelligence**: Syntax highlighting, linting, and IDE features in VS Code
- 📊 **Data Model Tracking**: Version control Excel table definitions and worksheet schemas
- 🛡️ **Backup & Recovery**: Never lose VBA code changes again

## ⚡ Quick Start

1. **Install**: Download `VBA Sync.xlam`, copy to your Excel add-ins folder and enable it
2. **Open**: Open your Excel workbook locally (avoid SharePoint direct links)
3. **Export**: Click **VBA Sync > Export** to create folder structure (named after your Excel file)
4. **Develop**: Edit code in VS Code, use Git for version control, get AI assistance
5. **Import**: Click **VBA Sync > Import** to load changes back into Excel

## 📁 What Gets Exported

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
    │       ├── validations.json        # Data validation rules (only if any)
    │       ├── conditional_formats.json
    │       ├── comments.json           # Cell comments + author
    │       ├── tables/<TableName>/
    │       │   ├── definition.json     # Schema, columns, calc formulas, overrides
    │       │   └── data.tsv            # Clean dataset (headers row 1, data rows below)
    │       └── drawings/
    │           ├── shapes.json         # Shapes, pictures, **with OnAction macro names**
    │           └── _assets/            # Embedded images
    └── STRUCTURE_SUMMARY.md            # Human-readable overview
```

The Excel/ folder is produced by reading the live Excel object model from VBA:

- Splits each worksheet into per-concern files so a data edit doesn't churn the styles file (and vice versa)
- Range-merges adjacent cells with identical styles or formulas (a calc column with one formula across 700 rows is one entry, not 700)
- Captures shape macros (`OnAction`) so AI agents can trace UI → VBA without opening Excel
- Deduplicates LAMBDA copies (Excel writes one per sheet that uses it; we collapse to one file)
- Skips volatile metadata that doesn't survive a no-change save

The result: small, readable diffs even after a "save with no changes". A column added to a table touches one or two files; today it would have touched 25.

Plus auto-generated Git configuration files:

- `.gitattributes` - Line endings & encoding rules for the various file types
- `.gitignore` - Excludes Excel temp files and system cruft
- `README.md` - Auto-generated documentation (created once, never overwritten)

## 🎬 Real-World Use Cases

- **Financial Models**: Version control formulas, table schemas, and VBA business logic
- **Reporting Tools**: Track changes to data processing pipelines and report generation
- **Dashboard Applications**: Collaborate on interactive Excel apps with professional workflows
- **Data Integration**: Manage API connections, database queries, and ETL processes
- **Automation Scripts**: Version control Excel automation with full change history

## 🔧 Installation

1. Download `VBA Sync.xlam` from this repository
2. **Unblock the file** (Windows security protection):
   - Right-click the downloaded `VBA Sync.xlam` file
   - Select "Properties"
   - Check "Unblock" at the bottom and click OK
3. **Double-click** the VBA Sync.xlam file - Excel will prompt to install it automatically
4. Click "Enable" when Excel asks about the add-in
5. The **VBA Sync** tab should appear in the ribbon

**If the simple method doesn't work:**

- Copy .xlam to Excel add-ins folder (`%APPDATA%\Microsoft\AddIns\`)
- Excel → File → Options → Add-ins → Excel Add-ins → Browse → Select file

**Important**: Enable "Trust access to the VBA project object model":

- File → Options → Trust Center → Trust Center Settings
- Macro Settings → Check "Trust access to the VBA project object model"

## 💡 Pro Tips

- **Git Integration**: Initialize a Git repository in your workbook folder for full version control
- **VS Code**: Install VBA language extensions for syntax highlighting and IntelliSense
- **AI Assistance**: Tools like GitHub Copilot can now understand your Excel data models
- **Team Workflow**: Use Git branches and pull requests for collaborative Excel development
- **Documentation**: The auto-generated `STRUCTURE_SUMMARY.md` helps onboard new team members

## ⚠️ Important Notes

- **Local Files Only**: Must open Excel files from local/synced folders (not SharePoint URLs)
- **VBA Only Import**: Excel structure export is for versioning; import only updates VBA code
- **Smart Filtering**: Empty document modules (sheets/ThisWorkbook) are skipped; removed components are cleaned up automatically
- **File Size**: Worksheet XML files are truncated to prevent huge files
- **Macro Security**: Ensure macro security settings allow the add-in to run
- **Backup Recommended**: Save/backup your workbook before first export

## ✨ What Makes This Special

This tool uniquely exports **both** VBA code and Excel file structure (tables, worksheets, workbook schema) to enable:

- 🧠 **AI Understanding**: AI tools can see your complete application architecture
- 🔄 **Full Version Control**: Track changes to both logic and data models
- 🏗️ **Professional Workflows**: Bring software engineering practices to Excel development
- 🤝 **Team Collaboration**: Review and merge changes like any software project

## 📄 License

This project is released under the MIT License - feel free to use, modify, and distribute!

## 👨‍💻 Credits

Created by **Arnaud Lavignolle** at **Axiom Project Services Pty Ltd**  
(with help from Claude and ChatGPT o3)  
Built to bridge the gap between Excel development and modern software engineering practices.

---

_Making Excel development as professional as any other software project._ 🎉
