Attribute VB_Name = "modSync"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
#Else
    Private Declare Function GetACP Lib "kernel32" () As Long
#End If

' Namespace of the embedded excel-control harness payload (a CustomXMLPart).
' See RefreshHarnessPayload (build) and WriteHarness (export).
Private Const HARNESS_NS As String = "urn:vba-sync:excel-control-harness:v1"

' MIT License
'
' Copyright (c) 2025 Arnaud Lavignolle
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' VBA Sync Module
'
' Bidirectional sync between Excel VBA projects and filesystem for version control,
' collaboration, and AI assistance.
'
' EXPORT: Extracts all VBA code (modules, classes, forms, sheets) and Excel structure
'         (tables, worksheets, workbook metadata) to organized folder structure.
'         Creates Git configuration files (.gitattributes, .gitignore, README.md).
'         Enables professional development workflows with version control, code review,
'         IDE editing, team collaboration, and AI assistance.
'
' IMPORT: Reads VBA code files from filesystem back into Excel VBA project.
'         Updates existing components or creates new ones as needed.
'         Preserves Excel structure (import is VBA code only).
'
' INSTALLATION:
' 1. Download VBA Sync.xlam add-in file
' 2. Copy to Excel add-ins folder (typically %APPDATA%\Microsoft\AddIns\)
'    Alternative: Double-click .xlam file and Excel will prompt to install
' 3. Open Excel > File > Options > Add-ins > Excel Add-ins > Browse
' 4. Select VBA Sync.xlam and check the box to enable it
' 5. Click OK - the VBA Sync tab should appear in the ribbon
' 6. If tab doesn't appear: restart Excel or check macro security settings
'
' REQUIREMENTS:
' - Workbook must be saved locally or on synced drive (not SharePoint URLs)
' - Enable "Trust access to the VBA project object model":
'   1. File > Options > Trust Center > Trust Center Settings
'   2. Macro Settings > Check "Trust access to the VBA project object model"
'   3. Click OK and restart Excel
' - RECOMMENDED: Save/backup your workbook before first export
'
' USAGE:
' Access via VBA Sync ribbon tab with Export and Import buttons.
' (Note: Ribbon tab appears when a workbook is open, works with .xlsx/.xlsm/.xls files)
'
' WORKFLOW:
' 1. Export: Click "Export" button - creates clean file structure for version control
' 2. Develop: Edit code in VS Code, use Git, get AI help
' 3. Import: Click "Import" button - loads changes back into Excel VBA project
'
' FOLDER STRUCTURE:
' Creates folder structure named after workbook:
'   MyWorkbook\
'     Objects\      - Sheet and ThisWorkbook modules
'     Modules\      - Standard modules
'     ClassModules\ - Class modules
'     Forms\        - UserForms
'     Excel\        - Workbook structure files
'
' ADDITIONAL FEATURES:
' - Creates .gitattributes, .gitignore, and README.md if they don't exist
' - Empty document modules (Sheets/ThisWorkbook) are skipped during export
' - Stale files are automatically cleaned up to keep folders tidy

Const GIT_ATTRIB As String = ".gitattributes"
Const GIT_IGNORE As String = ".gitignore"
Const README_FILE As String = "README.md"

'====================  Ribbon wrappers  ====================
Public Sub ExportProject(control As Object)
    DoExportProject
End Sub

Public Sub ImportProject(control As Object)
    DoImportProject
End Sub

'====================  MAIN ROUTINES  ======================
Private Sub DoExportAddin()
    Dim wb As Workbook: Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Sub

    ' Save workbook before export to ensure latest changes are included
    Debug.Print "VBA Sync: Saving workbook before export..."
    On Error Resume Next
    wb.Save
    If Err.Number <> 0 Then
      Debug.Print "VBA Sync: Error - Could not save workbook: " & Err.Description
      MsgBox "Error: Could not save workbook before export. " & _
              "Export cancelled to prevent data loss." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "VBA Sync"
      Err.Clear
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    
    Debug.Print "VBA Sync: Starting export of " & wb.Name
    
    Dim rootPath As String: rootPath = GetRootPath(wb)
    If rootPath = "" Then Exit Sub
    Dim repoPath As String: repoPath = GetRepoPath(wb)

    Debug.Print "VBA Sync: Export folder - " & rootPath
    Dim exported As Object: Set exported = CreateObject("Scripting.Dictionary")

    Debug.Print "VBA Sync: Exporting VBA components..."
    Dim comp As Object, subDir As String, fullPath As String
    For Each comp In wb.VBProject.VBComponents
        ' Skip empty document modules (sheets/ThisWorkbook auto-created by Excel)
        If comp.Type = vbext_ct_Document Then
            If IsDocModuleEmpty(comp) Then GoTo NextComponent
        End If

        subDir = rootPath & CompFolder(comp.Type) & "\"
        EnsureFolder subDir
        fullPath = subDir & comp.Name & GetExt(comp.Type)
        ExportComponent comp, fullPath, subDir, exported
NextComponent:
    Next

    PruneStaleFiles rootPath, exported

    Debug.Print "VBA Sync: Creating Git helper files..."
    WriteGitAttributes repoPath
    WriteGitIgnore repoPath
    WriteReadme repoPath, wb
    WriteVSCodeSettings repoPath
    WriteHarness repoPath
    
    Debug.Print "VBA Sync: Export completed successfully!"
    MsgBox "VBA Sync export completed successfully!" & vbCrLf & "Files exported to: " & rootPath, vbInformation, "VBA Sync"
End Sub

Private Sub DoExportProject()
    Dim wb As Workbook: Set wb = TargetWB()
    If wb Is Nothing Then Exit Sub

    ' Save workbook before export to ensure latest changes are included
    Debug.Print "VBA Sync: Saving workbook before export..."
    On Error Resume Next
    wb.Save
    If Err.Number <> 0 Then
      Debug.Print "VBA Sync: Error - Could not save workbook: " & Err.Description
      MsgBox "Error: Could not save workbook before export. " & _
              "Export cancelled to prevent data loss." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "VBA Sync"
      Err.Clear
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    
    Debug.Print "VBA Sync: Starting export of " & wb.Name
    
    Dim rootPath As String: rootPath = GetRootPath(wb)
    If rootPath = "" Then Exit Sub
    Dim repoPath As String: repoPath = GetRepoPath(wb)

    Debug.Print "VBA Sync: Export folder - " & rootPath
    Dim exported As Object: Set exported = CreateObject("Scripting.Dictionary")

    ' Export VBA components
    Debug.Print "VBA Sync: Exporting VBA components..."
    Dim comp As Object, subDir As String, fullPath As String
    For Each comp In wb.VBProject.VBComponents
        ' Skip empty document modules (sheets/ThisWorkbook auto-created by Excel)
        If comp.Type = vbext_ct_Document Then
            If IsDocModuleEmpty(comp) Then GoTo NextComponent
        End If

        subDir = rootPath & CompFolder(comp.Type) & "\"
        EnsureFolder subDir
        fullPath = subDir & comp.Name & GetExt(comp.Type)
        ExportComponent comp, fullPath, subDir, exported
NextComponent:
    Next

    'Export Excel file structure
    Debug.Print "VBA Sync: Extracting Excel file structure..."
    ExtractExcelStructure wb, rootPath, exported

    'Remove files on disk that weren't (re)exported this run
    Debug.Print "VBA Sync: Cleaning up stale files..."
    PruneStaleFiles rootPath, exported

    'Git helpers
    Debug.Print "VBA Sync: Creating Git helper files..."
    WriteGitAttributes repoPath
    WriteGitIgnore repoPath
    WriteReadme repoPath, wb
    WriteVSCodeSettings repoPath
    WriteHarness repoPath

    ' Check for per-sheet failures captured by modExcelExport. If any are
    ' present, the success MsgBox is replaced with a warning summary so the
    ' user can't miss that some sheets didn't fully export.
    Dim sheetErrLog As String: sheetErrLog = rootPath & "Excel\.export_sheet_errors.log"
    Dim sheetErrText As String: sheetErrText = ReadSheetErrors(sheetErrLog)
    If Len(sheetErrText) > 0 Then
        Debug.Print "VBA Sync: Export completed with sheet-level failures."
        MsgBox "VBA Sync export completed, but some sheets had issues:" & vbCrLf & vbCrLf & _
               sheetErrText & vbCrLf & _
               "Full details: " & sheetErrLog, vbExclamation, "VBA Sync"
    Else
        Debug.Print "VBA Sync: Export completed successfully!"
        MsgBox "VBA Sync export completed successfully!" & vbCrLf & "Files exported to: " & rootPath, vbInformation, "VBA Sync"
    End If
End Sub

' Parse .export_sheet_errors.log into a human-readable summary
' ("<sheet>: <phase> - <message>" lines). Returns "" if file missing or empty.
Private Function ReadSheetErrors(path As String) As String
    On Error GoTo Fail
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then ReadSheetErrors = "": Exit Function
    Dim ts As Object: Set ts = fso.OpenTextFile(path, 1)
    Dim contents As String
    If Not ts.AtEndOfStream Then contents = ts.ReadAll
    ts.Close
    If Len(contents) = 0 Then ReadSheetErrors = "": Exit Function

    Dim lines() As String: lines = Split(contents, vbCrLf)
    Dim summary As String, i As Long, count As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String: ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            ' Strip the timestamp + "VBA Sync: " prefix for brevity in the MsgBox
            Dim p As Long: p = InStr(ln, " VBA Sync: ")
            If p > 0 Then ln = Mid$(ln, p + 11)
            summary = summary & "  - " & ln & vbCrLf
            count = count + 1
        End If
    Next
    If count = 0 Then ReadSheetErrors = "": Exit Function
    ReadSheetErrors = summary
    Exit Function
Fail:
    ReadSheetErrors = ""
End Function

' Writes the Excel/ folder layout:
'   Excel/
'   |-- MANIFEST.json                          Workbook structure as JSON
'   |-- lambdas/<Name>.lambda                  Deduplicated LAMBDA defined names
'   |-- worksheets/<NN> - <SheetName>/
'   |   |-- data.tsv                           Cell values (resolved inline)
'   |   |-- formulas.json                      Range-collapsed cell formulas
'   |   |-- styles.json                        Range-merged resolved styles
'   |   |-- _meta.json                         Tab colour, panes, columns, etc.
'   |   |-- validations.json                   (only if any rules)
'   |   |-- conditional_formats.json           (only if any rules)
'   |   |-- comments.json                      (only if any comments)
'   |   |-- tables/<TableName>/
'   |   |   |-- definition.json                Schema, columns, calc formulas
'   |   |   `-- data.tsv                       Clean dataset
'   |   `-- drawings/shapes.json               Shapes, pictures, OnAction macros
'   `-- STRUCTURE_SUMMARY.md                   Human-readable overview
Private Sub ExtractExcelStructure(wb As Workbook, rootPath As String, exported As Object)
    On Error GoTo ExcelStructureError

    Dim excelDir As String: excelDir = rootPath & "Excel\"
    EnsureFolder excelDir

    Debug.Print "VBA Sync: Extracting Excel structure (live model)..."
    modExcelExport.DoExportExcelStructure wb, rootPath, exported

    ' Create Excel structure summary (live workbook -> readable .md)
    CreateExcelStructureSummary wb, excelDir, exported

    Exit Sub

ExcelStructureError:
    Debug.Print "VBA Sync: Excel structure extraction failed: " & Err.Description
    ' Continue without Excel structure rather than failing the whole VBA export
End Sub


Private Sub CreateExcelStructureSummary(wb As Workbook, excelDir As String, exported As Object)
    Dim summaryPath As String: summaryPath = excelDir & "STRUCTURE_SUMMARY.md"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim summary As String
    summary = "# Excel File Structure Summary" & vbCrLf & vbCrLf
    summary = summary & "Workbook: " & wb.Name & vbCrLf & vbCrLf
    
    ' Worksheet summary
    summary = summary & "## Worksheets" & vbCrLf
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        summary = summary & "- **" & ws.Name & "**"
        If ws.UsedRange.Rows.Count > 1 Then
            summary = summary & " (" & ws.UsedRange.Rows.Count & " rows, " & ws.UsedRange.Columns.Count & " cols)"
        End If
        summary = summary & vbCrLf
    Next
    summary = summary & vbCrLf
    
    ' Table summary
    summary = summary & "## Excel Tables" & vbCrLf
    Dim tableCount As Long: tableCount = 0
    For Each ws In wb.Worksheets
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            tableCount = tableCount + 1
            summary = summary & "- **" & tbl.Name & "** (" & ws.Name & ")"
            summary = summary & " - " & tbl.ListRows.Count & " rows, " & tbl.ListColumns.Count & " columns" & vbCrLf
        Next
    Next
    If tableCount = 0 Then summary = summary & "- No Excel tables found" & vbCrLf
    summary = summary & vbCrLf
    
    ' Named ranges summary
    summary = summary & "## Named Ranges" & vbCrLf
    If wb.Names.Count > 0 Then
        Dim nm As Name
        For Each nm In wb.Names
            On Error Resume Next
            summary = summary & "- **" & nm.Name & "**: " & nm.RefersTo & vbCrLf
            On Error GoTo 0
        Next
    Else
        summary = summary & "- No named ranges found" & vbCrLf
    End If
    summary = summary & vbCrLf
    
    summary = summary & "## Files Included" & vbCrLf
    summary = summary & "- `MANIFEST.json` - Workbook structure (sheets, tables, defined names, lambdas, calc settings)" & vbCrLf
    summary = summary & "- `lambdas/<Name>.lambda` - LAMBDA defined names (deduplicated across sheets)" & vbCrLf
    summary = summary & "- `worksheets/<NN> - <SheetName>/` - One folder per sheet:" & vbCrLf
    summary = summary & "    - `data.tsv` (cell values, sst-resolved)" & vbCrLf
    summary = summary & "    - `formulas.json` (range-collapsed)" & vbCrLf
    summary = summary & "    - `styles.json` (range-merged, dxf-resolved)" & vbCrLf
    summary = summary & "    - `_meta.json` (tab colour, panes, columns, hosted tables)" & vbCrLf
    summary = summary & "    - `validations.json`, `conditional_formats.json`, `comments.json` (only if any)" & vbCrLf
    summary = summary & "    - `tables/<TableName>/{definition.json, data.tsv}` - Excel tables" & vbCrLf
    summary = summary & "    - `drawings/{shapes.json, _assets/}` - shapes + pictures + macro refs" & vbCrLf
    summary = summary & vbCrLf

    modExcelExport.WriteIfChanged summaryPath, summary, exported, True
End Sub

Private Sub DoImportProject()
    Dim wb As Workbook: Set wb = TargetWB()
    If wb Is Nothing Then Exit Sub

    Debug.Print "VBA Sync: Starting import to " & wb.Name

    Dim rootPath As String: rootPath = GetRootPath(wb)
    If rootPath = "" Then Exit Sub
    If Dir(rootPath, vbDirectory) = "" Then
        MsgBox "Nothing to import - folder '" & rootPath & "' not found.", vbExclamation
        Debug.Print "VBA Sync: Import cancelled - source folder not found"
        Exit Sub
    End If
    
    Debug.Print "VBA Sync: Import folder - " & rootPath

    '-- remove all non-document components first
    ' Collect names into a list BEFORE removing — iterating Remove on the
    ' live VBComponents collection skips entries as indices shift, leaving
    ' stale components behind (notably UserForms, which then trip the
    ' "existing component" path below and get their .frm header text
    ' shoved into CodeModule).
    Debug.Print "VBA Sync: Removing existing VBA components..."
    Dim vc As Object
    Dim toRemove As Object: Set toRemove = CreateObject("Scripting.Dictionary")
    For Each vc In wb.VBProject.VBComponents
        If vc.Type <> vbext_ct_Document Then toRemove(vc.Name) = True
    Next
    Dim removeKey As Variant
    For Each removeKey In toRemove.Keys
        wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(CStr(removeKey))
    Next

    '-- iterate expected sub-folders
    Debug.Print "VBA Sync: Importing VBA components..."
    Dim subFolder As Variant, f As String, vbComp As Object, filePath As String
    For Each subFolder In Array("Modules", "ClassModules", "Forms", "Objects", "Misc")
        filePath = rootPath & subFolder & "\"
        If Dir(filePath, vbDirectory) <> "" Then
            f = Dir(filePath & "*.*")
            Do While Len(f) > 0
                ' Process only VBA source files. Skip:
                '   .frx                 — binary partner of .frm (auto-loaded by Import)
                '   .frx.fingerprint     — vba-sync sidecar for .frx churn suppression
                '   anything else        — future sidecars, stray files
                ' Without this guard, e.g. .frx.fingerprint matches the base
                ' name of an existing form and falls through to the "replace
                ' CodeModule" branch, stuffing the fingerprint text into the
                ' form's code section.
                Dim ext As String: ext = LCase$(Right$(f, 4))
                If ext <> ".bas" And ext <> ".cls" And ext <> ".frm" Then GoSub SkipFile

                Dim tgtName As String: tgtName = Split(f, ".")(0)
                Set vbComp = Nothing
                On Error Resume Next
                Set vbComp = wb.VBProject.VBComponents(tgtName)
                On Error GoTo 0

                If vbComp Is Nothing Then
                    wb.VBProject.VBComponents.Import filePath & f
                ElseIf LCase$(Right$(f, 4)) = ".frm" Then
                    ' .frm contains form-designer state + control declarations
                    ' + paired .frx pointer. Only VBComponents.Import can
                    ' rehydrate all that — overwriting CodeModule would drop
                    ' the form header into the code itself (the "fingerprint
                    ' in code" bug). Remove the stale shell, re-import fresh.
                    wb.VBProject.VBComponents.Remove vbComp
                    wb.VBProject.VBComponents.Import filePath & f
                Else
                    Dim txt As String
                    txt = CreateObject("Scripting.FileSystemObject") _
                          .OpenTextFile(filePath & f, 1).ReadAll
                    txt = CleanCode(txt)
                    With vbComp.CodeModule
                        .DeleteLines 1, .CountOfLines
                        .InsertLines 1, txt
                    End With
                End If
SkipFile:
                f = Dir
            Loop
        End If
    Next subFolder
    
    Debug.Print "VBA Sync: Import completed successfully!"
    MsgBox "VBA Sync import completed successfully!" & vbCrLf & "VBA components imported from: " & rootPath, vbInformation, "VBA Sync"
    ' Note: Excel structure import not implemented - files are for version control, collaboration, and AI assistance purposes only
End Sub

'====================  PATH / FILE HELPERS  ========================
Private Function GetRootPath(wb As Workbook) As String
    Dim p As String: p = wb.Path
    If p = "" Then
        MsgBox "Please save the workbook first.", vbExclamation
        Exit Function
    End If
    If LCase$(Left$(p, 4)) = "http" Then
        MsgBox "This workbook is open directly from SharePoint/Teams. " & _
               "Please open it from your local OneDrive sync folder or map it " & _
               "to a drive letter before running the export/import.", vbExclamation
        Exit Function
    End If
    If Right$(p, 1) <> "\" Then p = p & "\"
    
    ' Use workbook name as folder name, sanitizing invalid characters
    Dim folderName As String
    folderName = wb.Name
    If InStr(folderName, ".") > 0 Then
        folderName = Left(folderName, InStrRev(folderName, ".") - 1)
    End If
    folderName = Replace(folderName, "/", "_")
    folderName = Replace(folderName, "\", "_")
    folderName = Replace(folderName, ":", "_")
    folderName = Replace(folderName, "*", "_")
    folderName = Replace(folderName, "?", "_")
    folderName = Replace(folderName, """", "_")
    folderName = Replace(folderName, "<", "_")
    folderName = Replace(folderName, ">", "_")
    folderName = Replace(folderName, "|", "_")
    
    p = p & folderName & "\"
    EnsureFolder p
    GetRootPath = p
End Function

Private Function GetRepoPath(wb As Workbook) As String
    Dim p As String: p = wb.Path
    If Right$(p, 1) <> "\" Then p = p & "\"
    GetRepoPath = p
End Function

Private Sub EnsureFolder(fPath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fPath) Then fso.CreateFolder fPath
End Sub

'Delete any .bas/.cls/.frm/.frx/.xml file that wasn't exported this run
Private Sub PruneStaleFiles(rootPath As String, exported As Object)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    ' VBA component folders -- flat, fixed extensions
    Dim subFolder As Variant, folderPath As String
    For Each subFolder In Array("Modules", "ClassModules", "Forms", "Objects", "Misc")
        folderPath = rootPath & subFolder & "\"
        If fso.FolderExists(folderPath) Then
            Dim f As Object
            For Each f In fso.GetFolder(folderPath).Files
                Dim ext As String: ext = LCase$(fso.GetExtensionName(f.Path))
                If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "frx" Then
                    If Not exported.Exists(AddSlash(f.Path)) Then
                        On Error Resume Next
                        f.Delete True
                        On Error GoTo 0
                    End If
                End If
            Next
        End If
    Next

    ' Excel folder -- recursive walk (the v0.2 layout adds lambdas/ and may add
    ' more subfolders in v0.3). Anything under Excel/ that isn't in the export
    ' set gets pruned. The .normalize.log is preserved (we don't track it in
    ' exported, but it's harmless to leave).
    Dim excelDir As String: excelDir = rootPath & "Excel\"
    If fso.FolderExists(excelDir) Then
        PruneFolderRecursive fso.GetFolder(excelDir), exported
    End If
End Sub

' Recursively prune stale exported files from the Excel/ subtree.
' Walks bottom-up so we can detect newly-emptied folders and remove them.
' Preserves .normalize.log (it's overwritten by the next run anyway).
'
' Extensions covered: every output type the normaliser writes -- xml, json,
' lambda, tsv, md, plus the image assets png/jpg/jpeg/gif/emf/bmp. Adding
' a new output extension to the normaliser means adding it here too.
Private Sub PruneFolderRecursive(folder As Object, exported As Object)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object, sub_ As Object

    ' Recurse FIRST so child folders get pruned before we check if this one
    ' is empty. SubFolders enumeration mutates badly if we delete during
    ' iteration -- collect to a list first.
    Dim subFolders As Object: Set subFolders = CreateObject("Scripting.Dictionary")
    For Each sub_ In folder.SubFolders
        subFolders(sub_.Path) = True
    Next
    Dim subPath As Variant
    For Each subPath In subFolders.Keys
        On Error Resume Next
        PruneFolderRecursive fso.GetFolder(CStr(subPath)), exported
        On Error GoTo 0
    Next

    ' Files in THIS folder
    Dim toDelete As Object: Set toDelete = CreateObject("Scripting.Dictionary")
    For Each f In folder.Files
        Dim ext As String: ext = LCase$(fso.GetExtensionName(f.Path))
        ' Preserve the normaliser's log file
        If LCase$(f.Name) = ".normalize.log" Then
            ' skip
        ElseIf ext = "xml" Or ext = "md" Or ext = "json" Or ext = "lambda" Or _
               ext = "tsv" Or ext = "png" Or ext = "jpg" Or ext = "jpeg" Or _
               ext = "gif" Or ext = "emf" Or ext = "bmp" Then
            If Not exported.Exists(AddSlash(f.Path)) Then
                toDelete(f.Path) = True
            End If
        End If
    Next
    Dim p As Variant
    For Each p In toDelete.Keys
        On Error Resume Next
        fso.DeleteFile p, True
        On Error GoTo 0
    Next

    ' If this folder is now empty AND it's not the root Excel/ folder, remove it.
    ' (We never delete the root Excel/ -- export creates it fresh.)
    If LCase$(folder.Name) <> "excel" Then
        On Error Resume Next
        Dim refreshed As Object: Set refreshed = fso.GetFolder(folder.Path)
        If refreshed.Files.Count = 0 And refreshed.SubFolders.Count = 0 Then
            folder.Delete True
        End If
        On Error GoTo 0
    End If
End Sub

Private Function AddSlash(p As String) As String
    AddSlash = Replace$(p, "/", "\")
End Function

'====================  GIT FILE WRITERS  ======================
Private Sub WriteGitAttributes(basePath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim aPath As String: aPath = basePath & GIT_ATTRIB
    If fso.FileExists(aPath) Then
        Debug.Print "VBA Sync: .gitattributes already exists, skipping"
        Exit Sub
    End If
    Debug.Print "VBA Sync: Creating .gitattributes"

    Dim txt As String
    txt = "# Auto-generated by VBA Sync Add-in on " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
          "# Treat VBA text modules as CRLF-normalised text files (Windows style)" & vbCrLf & _
          "*.bas text eol=crlf" & vbCrLf & _
          "*.cls text eol=crlf" & vbCrLf & _
          "*.frm text eol=crlf" & vbCrLf & _
          vbCrLf & _
          "# Excel structure files (XML + Markdown + TSV use CRLF; JSON/lambda use LF)" & vbCrLf & _
          "*.xml text eol=crlf" & vbCrLf & _
          "*.md text eol=crlf" & vbCrLf & _
          "*.tsv text eol=crlf" & vbCrLf & _
          "*.json text eol=lf" & vbCrLf & _
          "*.lambda text eol=lf" & vbCrLf & _
          vbCrLf & _
          "# Binary partner of UserForms" & vbCrLf & _
          "*.frx binary" & vbCrLf & _
          vbCrLf & _
          "# Ignore diff for Excel workbooks" & vbCrLf & _
          "*.xls* binary" & vbCrLf

    Dim ts
    Set ts = fso.CreateTextFile(aPath, True)
    ts.Write txt
    ts.Close
End Sub

Private Sub WriteGitIgnore(basePath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim iPath As String: iPath = basePath & GIT_IGNORE

    ' Harness rules every exported repo must ignore: the per-session
    ' runtime dir, and the optional VBA-password file (the -VbaPassword
    ' file fallback). Bare names so they match wherever the harness puts
    ' them. Appended even to a .gitignore written before these existed.
    Const RUNTIME_DIR As String = ".excel-control/"
    Const SECRET_LINE As String = ".vba-password"

    If fso.FileExists(iPath) Then
        Dim existing As String: existing = ReadAllText(iPath)
        Dim addLines As String: addLines = ""
        If InStr(1, existing, RUNTIME_DIR, vbTextCompare) = 0 Then
            addLines = addLines & "# excel-control harness runtime/session data" & vbCrLf & _
                       RUNTIME_DIR & vbCrLf
        End If
        If InStr(1, existing, SECRET_LINE, vbTextCompare) = 0 Then
            addLines = addLines & "# excel-control harness VBA password (never commit)" & vbCrLf & _
                       SECRET_LINE & vbCrLf
        End If
        If Len(addLines) = 0 Then
            Debug.Print "VBA Sync: .gitignore already has harness rules, skipping"
        Else
            Debug.Print "VBA Sync: appending harness rules to .gitignore"
            Dim ap
            Set ap = fso.OpenTextFile(iPath, 8, True)   ' 8 = ForAppending
            ap.Write vbCrLf & addLines
            ap.Close
        End If
        Exit Sub
    End If
    Debug.Print "VBA Sync: Creating .gitignore"

    Dim txt As String
    txt = "# Auto-generated by VBA Sync Add-in on " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
          "# Ignore Excel/Office temp and cache files" & vbCrLf & _
          "~$*" & vbCrLf & _
          "*.tmp" & vbCrLf & _
          "*.bak" & vbCrLf & _
          "*.log" & vbCrLf & _
          "*.ldb" & vbCrLf & _
          "*.laccdb" & vbCrLf & _
          "*.asd" & vbCrLf & _
          "*.wbk" & vbCrLf & _
          vbCrLf & _
          "# Office autosave / lock files" & vbCrLf & _
          "*.owner" & vbCrLf & _
          vbCrLf & _
          "# VBA Sync temporary files" & vbCrLf & _
          "*.temp.zip" & vbCrLf & _
          vbCrLf & _
          "# OS cruft" & vbCrLf & _
          "Thumbs.db" & vbCrLf & _
          ".DS_Store" & vbCrLf & _
          vbCrLf & _
          "# IDE/project folders (optional)" & vbCrLf & _
          ".vs/" & vbCrLf & _
          ".idea/" & vbCrLf
    ' Second statement: VBA caps line-continuations at 25 per logical
    ' statement — keep the harness rules split off from the block above.
    txt = txt & vbCrLf & _
          "# excel-control harness runtime/session data" & vbCrLf & _
          RUNTIME_DIR & vbCrLf & vbCrLf & _
          "# excel-control harness VBA password (never commit)" & vbCrLf & _
          SECRET_LINE & vbCrLf

    Dim ts
    Set ts = fso.CreateTextFile(iPath, True)
    ts.Write txt
    ts.Close
End Sub

Private Sub WriteReadme(basePath As String, wb As Workbook)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim rPath As String: rPath = basePath & README_FILE
    If fso.FileExists(rPath) Then
        Debug.Print "VBA Sync: README.md already exists, skipping"
        Exit Sub
    End If
    Debug.Print "VBA Sync: Creating README.md"

    Dim txt As String
    Dim workbookName As String: workbookName = wb.Name
    If InStr(workbookName, ".") > 0 Then
        workbookName = Left(workbookName, InStrRev(workbookName, ".") - 1)
    End If
    
    txt = "# " & workbookName & " - VBA Project" & vbCrLf & vbCrLf & _
          "Your Excel VBA project has been exported for modern development with Git, VS Code, " & _
          "and AI assistance. Edit the .bas, .cls, and .frm files, then import back to Excel." & vbCrLf & vbCrLf & _
          "## QUICK START" & vbCrLf & _
          "```" & vbCrLf & _
          "git init" & vbCrLf & _
          "git add ." & vbCrLf & _
          "git commit -m ""Initial export of " & workbookName & """" & vbCrLf & _
          "```" & vbCrLf
    
    txt = txt & vbCrLf & _
          "## WORKFLOW" & vbCrLf & _
          "1. Edit .bas/.cls/.frm files in VS Code" & vbCrLf & _
          "2. Commit changes with Git" & vbCrLf & _
          "3. Use **VBA Sync > Import** to load back into Excel" & vbCrLf & vbCrLf & _
          "## PROJECT STRUCTURE" & vbCrLf & "```" & vbCrLf
    
    txt = txt & "This Folder/" & vbCrLf & _
          "|-- Modules/              # Standard VBA modules (.bas files)" & vbCrLf & _
          "|-- ClassModules/         # VBA class modules (.cls files)" & vbCrLf & _
          "|-- Forms/                # UserForms (.frm + .frx files)" & vbCrLf & _
          "|-- Objects/              # ThisWorkbook & Sheet code-behind (.cls files)" & vbCrLf & _
          "|-- Excel/                # Excel structure split into per-concern files" & vbCrLf & _
          "|   |-- MANIFEST.json         # Sheets, tables, defined names, lambdas" & vbCrLf & _
          "|   |-- lambdas/              # Deduplicated LAMBDA defined names" & vbCrLf & _
          "|   |-- worksheets/<NN>-<Name>/" & vbCrLf & _
          "|   |       data.tsv, formulas.json, styles.json, _meta.json" & vbCrLf & _
          "|   |       tables/<TableName>/{definition.json, data.tsv}" & vbCrLf & _
          "|   |       drawings/{shapes.json, _assets/}" & vbCrLf & _
          "|   `-- STRUCTURE_SUMMARY.md  # Human-readable data model overview" & vbCrLf
    txt = txt & _
          "|-- .claude/skills/excel-control/  # excel-control: AI agent harness" & vbCrLf & _
          "|   |-- SKILL.md              # Claude Code auto-discovers this skill" & vbCrLf & _
          "|   `-- tools/                # PowerShell session harness" & vbCrLf & _
          "|       |-- start-session.ps1     # Spawn a long-running Excel session" & vbCrLf & _
          "|       |-- INTERFACE.md          # Full command + event reference" & vbCrLf & _
          "|       `-- clsAssert.cls         # Optional assertion lib for run_tests" & vbCrLf & _
          "|-- .gitattributes        # Git configuration for VBA files" & vbCrLf & _
          "|-- .gitignore            # Excludes temp files from version control" & vbCrLf & _
          "`-- README.md             # This file" & vbCrLf
    
    txt = txt & "```" & vbCrLf & vbCrLf & _
          "## TOOLS" & vbCrLf & _
          "- **VS Code**: Install VBA extensions for syntax highlighting" & vbCrLf & _
          "- **AI Tools**: GitHub Copilot, Claude, ChatGPT work with your exported code" & vbCrLf & _
          "- **Git**: Use branches (`git checkout -b feature-name`) for development" & vbCrLf
    
    txt = txt & vbCrLf & _
          "## AGENT HARNESS" & vbCrLf & _
          "An AI agent (Claude Code, MCP client, bash script) can drive this" & vbCrLf & _
          "workbook end-to-end via the bundled PowerShell session harness:" & vbCrLf & _
          "run macros, read/write cells, respond to dialogs, capture runtime" & vbCrLf & _
          "errors, run a discovered Test_* suite. The harness skill lives at" & vbCrLf & _
          "``.claude/skills/excel-control/`` — see SKILL.md there for usage" & vbCrLf & _
          "patterns and tools/INTERFACE.md for the full reference." & vbCrLf & vbCrLf & _
          "## NOTES" & vbCrLf & _
          "- Edit code files directly, then import back to Excel" & vbCrLf & _
          "- Form design must be done in Excel (only code imports)" & vbCrLf & _
          "- Excel/ folder files are for AI context, not editing" & vbCrLf
    
    txt = txt & "- Check `Excel/STRUCTURE_SUMMARY.md` for data model overview" & vbCrLf & vbCrLf
    
    txt = txt & "---" & vbCrLf & _
          "*Exported on " & Format(Now, "yyyy-mm-dd") & " using VBA Sync by Arnaud Lavignolle*"

    Dim ts
    Set ts = fso.CreateTextFile(rPath, True)
    ts.Write txt
    ts.Close
End Sub

' VBA's comp.Export writes files in the system ANSI codepage, not UTF-8. By
' default VS Code reads files as UTF-8 and shows mojibake for any non-ASCII
' character (Polish ą, French é, the Euro sign, em-dashes, smart quotes, etc.).
' Drop a .vscode/settings.json next to the exported VBA folder so VS Code
' opens the files in the codepage that produced them. Created once per export
' folder, never overwritten — so a contributor on a different codepage can
' edit it manually if they need to.
Private Sub WriteVSCodeSettings(basePath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim vsDir As String: vsDir = basePath & ".vscode\"
    Dim sPath As String: sPath = vsDir & "settings.json"
    If fso.FileExists(sPath) Then
        Debug.Print "VBA Sync: .vscode/settings.json already exists, skipping"
        Exit Sub
    End If
    EnsureFolder vsDir
    Debug.Print "VBA Sync: Creating .vscode/settings.json (codepage " & GetACP() & ")"

    Dim enc As String: enc = VSCodeEncodingName(GetACP())
    Dim txt As String
    txt = "{" & vbCrLf & _
          "  // Auto-generated by VBA Sync. VBA exports are in the system ANSI" & vbCrLf & _
          "  // codepage; this tells VS Code to read VBA files in that codepage" & vbCrLf & _
          "  // so non-ASCII characters render correctly. files.associations" & vbCrLf & _
          "  // pins .bas/.cls/.frm to the 'vba' language so the per-language" & vbCrLf & _
          "  // encoding override applies even when no VBA extension is" & vbCrLf & _
          "  // installed. autoGuessEncoding is off because pure-ASCII files" & vbCrLf & _
          "  // parse cleanly as UTF-8 and the guesser would override us." & vbCrLf & _
          "  // Other file types (README.md, .json, .xml) keep VS Code defaults." & vbCrLf & _
          "  ""files.autoGuessEncoding"": false," & vbCrLf & _
          "  ""files.associations"": {" & vbCrLf & _
          "    ""*.bas"": ""vba""," & vbCrLf & _
          "    ""*.cls"": ""vba""," & vbCrLf & _
          "    ""*.frm"": ""vba""" & vbCrLf & _
          "  }," & vbCrLf & _
          "  ""[vba]"": { ""files.encoding"": """ & enc & """ }" & vbCrLf & _
          "}" & vbCrLf

    Dim ts As Object
    Set ts = fso.CreateTextFile(sPath, True)
    ts.Write txt
    ts.Close
End Sub

' --- excel-control harness bundling ------------------------------------
' The harness payload (PowerShell scripts + docs) is embedded in this
' .xlam as a single CustomXMLPart, so Export can write a complete tools/
' folder on ANY machine with no dependency on an external excel-control\
' folder sitting next to the .xlam.
'
'   RefreshHarnessPayload   build time, dev machine:
'                           excel-control\ + MANIFEST.json  ->  CustomXMLPart
'   WriteHarness            run time, any machine:
'                           CustomXMLPart  ->  exported repo
'
' Each payload entry carries its own destination, so WriteHarness needs
' nothing external. Re-run RefreshHarnessPayload whenever a harness file
' changes, before saving the rebuilt .xlam.

' Build-time. Reads excel-control\MANIFEST.json next to the .xlam,
' base64-encodes every file it lists, and stores them all in one
' CustomXMLPart. Public so the .xlam rebuild can invoke it via the
' excel-control harness (run_macro RefreshHarnessPayload) between
' sync_vba and save_workbook.
Public Sub RefreshHarnessPayload()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim xlamDir As String: xlamDir = ThisWorkbook.Path
    If Right$(xlamDir, 1) <> "\" Then xlamDir = xlamDir & "\"
    Dim srcRoot As String: srcRoot = xlamDir & "excel-control\"
    Dim manifestPath As String: manifestPath = srcRoot & "MANIFEST.json"
    If Not fso.FileExists(manifestPath) Then
        Err.Raise vbObjectError + 720, "RefreshHarnessPayload", _
            "excel-control\MANIFEST.json not found at: " & manifestPath
    End If

    Dim manifest As Object
    Set manifest = JsonConverter.ParseJson(ReadAllText(manifestPath))
    If Not manifest.Exists("files") Then
        Err.Raise vbObjectError + 721, "RefreshHarnessPayload", _
            "MANIFEST.json has no 'files' array."
    End If
    Dim files As Object: Set files = manifest("files")

    Dim xml As String
    xml = "<harness xmlns=""" & HARNESS_NS & """>"
    Dim entry As Variant
    For Each entry In files
        Dim srcRel As String: srcRel = CStr(entry("src"))
        Dim dstRel As String: dstRel = CStr(entry("dst"))
        Dim srcPath As String: srcPath = srcRoot & Replace(srcRel, "/", "\")
        If Not fso.FileExists(srcPath) Then
            Err.Raise vbObjectError + 722, "RefreshHarnessPayload", _
                "manifest lists a missing file: " & srcPath
        End If
        Dim raw() As Byte: raw = ReadFileBytes(srcPath)
        xml = xml & "<file path=""" & dstRel & """>" & _
              BytesToBase64(raw) & "</file>"
    Next
    xml = xml & "</harness>"

    DeleteHarnessPayload
    ThisWorkbook.CustomXMLParts.Add xml
    Debug.Print "VBA Sync: harness payload refreshed (" & _
                files.Count & " files embedded)"
End Sub

' Removes any previously embedded harness payload part.
Private Sub DeleteHarnessPayload()
    Dim parts As Object
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(HARNESS_NS)
    Dim i As Long
    For i = parts.Count To 1 Step -1
        parts(i).Delete
    Next
End Sub

' Run-time. Writes the embedded harness payload into the exported repo so
' an AI agent can drive the workbook via the excel-control session
' protocol. Each payload entry carries its own destination relative to
' basePath. If no payload is embedded (an older .xlam), this silently
' skips and the rest of the export is unaffected.
Private Sub WriteHarness(basePath As String)
    Dim parts As Object
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(HARNESS_NS)
    If parts.Count = 0 Then
        Debug.Print "VBA Sync: excel-control harness payload not embedded (skipping tools/)"
        Exit Sub
    End If
    Debug.Print "VBA Sync: Writing excel-control harness from embedded payload"

    Dim dom As Object: Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.LoadXML parts(1).xml
    dom.setProperty "SelectionNamespaces", "xmlns:h='" & HARNESS_NS & "'"

    Dim nodes As Object: Set nodes = dom.SelectNodes("//h:file")
    Dim node As Object
    For Each node In nodes
        Dim dst As String: dst = node.getAttribute("path")
        Dim outPath As String: outPath = basePath & Replace(dst, "/", "\")
        EnsureFolderTree ParentFolder(outPath)
        Dim raw() As Byte: raw = Base64ToBytes(node.Text)
        WriteFileBytes outPath, raw
    Next
    Debug.Print "VBA Sync: harness written (" & nodes.Length & " files)"
End Sub

' Create fPath and any missing parent folders.
Private Sub EnsureFolderTree(fPath As String)
    If Len(fPath) = 0 Then Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(fPath) Then Exit Sub
    Dim parent As String: parent = fso.GetParentFolderName(fPath)
    If Len(parent) > 0 Then EnsureFolderTree parent
    If Not fso.FolderExists(fPath) Then fso.CreateFolder fPath
End Sub

' Directory portion of a backslash path (no trailing separator).
Private Function ParentFolder(path As String) As String
    Dim p As Long: p = InStrRev(path, "\")
    If p > 0 Then ParentFolder = Left$(path, p - 1) Else ParentFolder = ""
End Function

' base64 round-trip via an MSXML node typed bin.base64 (ships with
' Windows — no extra reference). Byte arrays preserve encoding/BOM.
Private Function BytesToBase64(bytes() As Byte) As String
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    BytesToBase64 = node.Text
End Function

Private Function Base64ToBytes(b64 As String) As Byte()
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64
    Base64ToBytes = node.nodeTypedValue
End Function

' Raw byte file IO via ADODB.Stream — exact bytes, no codepage coercion.
Private Function ReadFileBytes(path As String) As Byte()
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    s.Type = 1                  ' adTypeBinary
    s.Open
    s.LoadFromFile path
    ReadFileBytes = s.Read
    s.Close
End Function

Private Sub WriteFileBytes(path As String, bytes() As Byte)
    Dim s As Object: Set s = CreateObject("ADODB.Stream")
    s.Type = 1                  ' adTypeBinary
    s.Open
    s.Write bytes
    s.SaveToFile path, 2        ' adSaveCreateOverWrite
    s.Close
End Sub

' Map a Windows ANSI codepage number to the VS Code encoding name.
' Covers the codepages that ship with Windows. Falls back to windows1252 for
' anything unknown — a sensible default for Western users and harmless if
' wrong (the file just keeps showing mojibake, no data loss).
Private Function VSCodeEncodingName(codepage As Long) As String
    Select Case codepage
        Case 1250: VSCodeEncodingName = "windows1250"   ' Central European
        Case 1251: VSCodeEncodingName = "windows1251"   ' Cyrillic
        Case 1252: VSCodeEncodingName = "windows1252"   ' Western European
        Case 1253: VSCodeEncodingName = "windows1253"   ' Greek
        Case 1254: VSCodeEncodingName = "windows1254"   ' Turkish
        Case 1255: VSCodeEncodingName = "windows1255"   ' Hebrew
        Case 1256: VSCodeEncodingName = "windows1256"   ' Arabic
        Case 1257: VSCodeEncodingName = "windows1257"   ' Baltic
        Case 1258: VSCodeEncodingName = "windows1258"   ' Vietnamese
        Case 932:  VSCodeEncodingName = "shiftjis"      ' Japanese
        Case 936:  VSCodeEncodingName = "gbk"           ' Simplified Chinese
        Case 949:  VSCodeEncodingName = "uhc"           ' Korean
        Case 950:  VSCodeEncodingName = "big5"          ' Traditional Chinese
        Case 874:  VSCodeEncodingName = "windows874"    ' Thai
        Case Else: VSCodeEncodingName = "windows1252"
    End Select
End Function

'====================  COMPONENT HELPERS  ===================
Private Function CompFolder(t As Long) As String
    Select Case t
        Case vbext_ct_StdModule:     CompFolder = "Modules"
        Case vbext_ct_ClassModule:   CompFolder = "ClassModules"
        Case vbext_ct_MSForm:        CompFolder = "Forms"
        Case vbext_ct_Document:      CompFolder = "Objects"
        Case Else:                   CompFolder = "Misc"
    End Select
End Function

Private Function GetExt(t As Long) As String
    Select Case t
        Case vbext_ct_StdModule:     GetExt = ".bas"
        Case vbext_ct_ClassModule:   GetExt = ".cls"
        Case vbext_ct_MSForm:        GetExt = ".frm"   'paired .frx auto-exported
        Case vbext_ct_Document:      GetExt = ".cls"   'sheet / ThisWorkbook code-behind
        Case Else:                   GetExt = ".bas"
    End Select
End Function

' Check if Document module is empty (only Option Explicit/whitespace)
Private Function IsDocModuleEmpty(vbComp As Object) As Boolean
    Dim cm As Object: Set cm = vbComp.CodeModule
    Dim txt As String
    
    If cm.CountOfLines = 0 Then IsDocModuleEmpty = True: Exit Function
    txt = cm.Lines(1, cm.CountOfLines)
    txt = CleanCode(txt)

    Dim ln As Variant, hasRealCode As Boolean
    For Each ln In Split(txt, vbCrLf)
        Dim t As String: t = Trim$(ln)
        If Len(t) > 0 And LCase$(t) <> "option explicit" Then
            hasRealCode = True
            Exit For
        End If
    Next
    IsDocModuleEmpty = Not hasRealCode
End Function

'====================  UTILITY HELPERS  =====================
Private Function TargetWB() As Workbook
    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No workbook is active.", vbExclamation
    ElseIf wb.IsAddin Then
        MsgBox "Switch to the workbook you want to export, not the add-in tab.", vbExclamation
        Set wb = Nothing
    End If
    Set TargetWB = wb
End Function

Private Function CleanCode(src As String) As String
    Dim ln As Variant, out$, inBegin As Boolean, t As String
    ' Normalize line endings: files hand-edited in VS Code / other modern
    ' editors default to LF, while Excel's own comp.Export writes CRLF.
    ' Splitting on vbCrLf alone collapses an LF-only file into a single
    ' "line" that starts with "VERSION 1.0 CLASS" and matches the wholesale
    ' header skip — silently emptying the target CodeModule on import.
    src = Replace(Replace(src, vbCrLf, vbLf), vbCr, vbLf)
    For Each ln In Split(src, vbLf)
        t = Trim$(ln)
        Select Case True
            Case t Like "VERSION *", t Like "Attribute VB_*":
            Case Left$(t, 5) = "BEGIN":                       inBegin = True
            Case inBegin And t = "END":                        inBegin = False
            Case inBegin:
            Case Else:                                          out = out & ln & vbCrLf
        End Select
    Next
    CleanCode = RTrim$(out)
End Function

' Export one VBA component (.bas/.cls/.frm/.frx) to disk.
'
' For MSForms: Excel rewrites a few metadata bytes inside .frx on every
' comp.Export even when the form wasn't touched (MS-OFORMS has padding
' bytes that leak uninitialised heap memory). To avoid noise diffs we
' build a deterministic fingerprint by introspecting comp.Designer.Controls,
' hash it, and store it in a `<FormName>.frx.fingerprint` sidecar. On
' next export, if the fingerprint matches the stored sidecar the new .frx
' is guaranteed equivalent -- restore the previous .frx so git sees no
' change. If the fingerprint differs, both .frx + new sidecar are written.
'
' Known gap: forms with embedded Picture controls fingerprint only
' Picture.Width/Height/Type, not the picture bytes, so a picture swap on
' an otherwise-unchanged form looks identical to the fingerprint.
Private Sub ExportComponent(comp As Object, fullPath As String, subDir As String, exported As Object)
    Dim oldFrxPath As String, hadOldFrx As Boolean, newFrx As String, fpPath As String
    Dim oldFingerprint As String

    If comp.Type = vbext_ct_MSForm Then
        newFrx = subDir & comp.Name & ".frx"
        fpPath = subDir & comp.Name & ".frx.fingerprint"
        If FileExistsLocal(newFrx) Then
            oldFrxPath = newFrx & ".vbasync_prev"
            CopyFileLocal newFrx, oldFrxPath
            hadOldFrx = True
        End If
        If FileExistsLocal(fpPath) Then oldFingerprint = ReadAllText(fpPath)
    End If

    comp.Export fullPath
    CleanExportedFile fullPath  ' Remove trailing empty lines + collapse .frm blanks

    If comp.Type = vbext_ct_MSForm Then
        Dim newFingerprint As String: newFingerprint = ComputeFormFingerprint(comp)
        If hadOldFrx And Len(oldFingerprint) > 0 And newFingerprint = oldFingerprint Then
            ' Form is semantically unchanged; restore the previous .frx so
            ' git sees no change. Fingerprint sidecar already correct.
            CopyFileLocal oldFrxPath, newFrx
        ElseIf Len(newFingerprint) > 0 Then
            ' Form changed (or first export). Write the new fingerprint
            ' alongside the new .frx; both committed together.
            WriteAllText fpPath, newFingerprint
            exported(AddSlash(fpPath)) = True
        End If
        If hadOldFrx Then
            On Error Resume Next: Kill oldFrxPath: On Error GoTo 0
        End If
        ' Always count the .frx and .fingerprint as "exported" so PruneStaleFiles
        ' doesn't sweep them when the fingerprint matched (no new write).
        exported(AddSlash(newFrx)) = True
        If FileExistsLocal(fpPath) Then exported(AddSlash(fpPath)) = True
    End If
    exported(AddSlash(fullPath)) = True
End Sub

' Build a deterministic fingerprint of a UserForm's semantic content from
' the in-memory Designer. Output is a multi-line text blob containing every
' control's identifying properties, sorted by control name for stability.
' Hash is the raw text (small forms) -- callers can SHA-256 it later if size
' becomes a concern. Returns "" if the Designer can't be introspected
' (typical for forms loaded but not designable in this VBE state).
Private Function ComputeFormFingerprint(comp As Object) As String
    On Error GoTo Fail
    Dim designer As Object
    Set designer = comp.designer
    If designer Is Nothing Then Exit Function

    Dim parts As New Collection
    parts.Add "FORM|" & comp.Name & _
              "|caption=" & SafeProp(designer, "Caption") & _
              "|width=" & SafeProp(designer, "Width") & _
              "|height=" & SafeProp(designer, "Height")

    ' Collect control fingerprints, then sort by Name for determinism.
    Dim ctlFps As New Collection
    Dim ctl As Object
    For Each ctl In designer.Controls
        ctlFps.Add ControlFingerprint(ctl)
    Next
    Dim sorted() As String: sorted = SortStringCollectionToArray(ctlFps)
    Dim i As Long
    For i = LBound(sorted) To UBound(sorted)
        parts.Add sorted(i)
    Next

    Dim out As String, p As Variant
    For Each p In parts
        out = out & CStr(p) & vbLf
    Next
    ComputeFormFingerprint = out
    Exit Function
Fail:
    ComputeFormFingerprint = ""
End Function

Private Function ControlFingerprint(ctl As Object) As String
    On Error Resume Next
    ControlFingerprint = "CTL|" & SafeProp(ctl, "Name") & _
        "|type=" & TypeName(ctl) & _
        "|top=" & SafeProp(ctl, "Top") & _
        "|left=" & SafeProp(ctl, "Left") & _
        "|width=" & SafeProp(ctl, "Width") & _
        "|height=" & SafeProp(ctl, "Height") & _
        "|tabIndex=" & SafeProp(ctl, "TabIndex") & _
        "|caption=" & SafeProp(ctl, "Caption") & _
        "|value=" & SafeProp(ctl, "Value") & _
        "|tag=" & SafeProp(ctl, "Tag") & _
        "|tooltip=" & SafeProp(ctl, "ControlTipText") & _
        "|enabled=" & SafeProp(ctl, "Enabled") & _
        "|visible=" & SafeProp(ctl, "Visible") & _
        "|fontName=" & SafeProp(ctl, "Font.Name") & _
        "|fontSize=" & SafeProp(ctl, "Font.Size") & _
        "|fontBold=" & SafeProp(ctl, "Font.Bold")
    On Error GoTo 0
End Function

' Read a property by name via late binding; tolerates missing properties.
' Supports dotted access ("Font.Name") with one level of nesting.
Private Function SafeProp(obj As Object, propName As String) As String
    On Error Resume Next
    Dim dotPos As Long: dotPos = InStr(propName, ".")
    Dim result As Variant
    If dotPos > 0 Then
        Dim outer As String: outer = Left$(propName, dotPos - 1)
        Dim inner As String: inner = Mid$(propName, dotPos + 1)
        Dim outerObj As Object
        Set outerObj = CallByName(obj, outer, VbGet)
        If Not outerObj Is Nothing Then result = CallByName(outerObj, inner, VbGet)
    Else
        result = CallByName(obj, propName, VbGet)
    End If
    If IsObject(result) Then
        SafeProp = "<obj>"
    ElseIf IsNull(result) Or IsEmpty(result) Then
        SafeProp = ""
    Else
        SafeProp = CStr(result)
    End If
    On Error GoTo 0
End Function

Private Function SortStringCollectionToArray(c As Collection) As String()
    Dim n As Long: n = c.Count
    Dim arr() As String
    If n = 0 Then ReDim arr(0 To 0): arr(0) = "": SortStringCollectionToArray = arr: Exit Function
    ReDim arr(0 To n - 1)
    Dim i As Long: i = 0
    Dim x As Variant
    For Each x In c: arr(i) = CStr(x): i = i + 1: Next
    ' Insertion sort -- small N (form control counts).
    Dim j As Long, key As String
    For i = 1 To n - 1
        key = arr(i)
        j = i - 1
        Do While j >= 0
            If arr(j) <= key Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = key
    Next
    SortStringCollectionToArray = arr
End Function

Private Function ReadAllText(path As String) As String
    On Error GoTo Fail
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Dim ts As Object: Set ts = fso.OpenTextFile(path, 1)
    If Not ts.AtEndOfStream Then ReadAllText = ts.ReadAll
    ts.Close
    Exit Function
Fail:
    ReadAllText = ""
End Function

Private Sub WriteAllText(path As String, content As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object: Set ts = fso.CreateTextFile(path, True)
    ts.Write content
    ts.Close
End Sub

Private Function FileExistsLocal(p As String) As Boolean
    On Error Resume Next
    FileExistsLocal = (Len(Dir$(p)) > 0)
    On Error GoTo 0
End Function

Private Sub CopyFileLocal(src As String, dst As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile src, dst, True
    On Error GoTo 0
End Sub

Private Sub CleanExportedFile(filePath As String)
    On Error GoTo CleanError

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Sub

    ' Read the file content
    Dim content As String
    content = fso.OpenTextFile(filePath, 1).ReadAll

    ' For .frm files only: collapse runs of 3+ blank lines down to 1. VBA's
    ' comp.Export writes the form code section with a chunk of blank lines
    ' between the Attribute block and Option Explicit; comp.Import preserves
    ' them, and on the next export Excel adds another. The result is one
    ' extra blank line per round-trip, growing forever. Collapsing 3+ to 1
    ' is convergent and leaves intentional 1- or 2-blank spacing alone.
    If LCase$(Right$(filePath, 4)) = ".frm" Then
        Dim re As Object: Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.MultiLine = False
        ' Match 3 or more consecutive vbCrLf -> collapse to 2 (one blank line).
        re.Pattern = "(\r\n){3,}"
        content = re.Replace(content, vbCrLf & vbCrLf)
    End If

    ' Remove trailing empty lines (but preserve one final line break)
    Do While Right$(content, 4) = vbCrLf & vbCrLf
        content = Left$(content, Len(content) - 2)
    Loop

    ' Ensure file ends with exactly one line break
    If Right$(content, 2) <> vbCrLf And Len(content) > 0 Then
        content = content & vbCrLf
    End If

    ' Write back the cleaned content
    Dim ts As Object: Set ts = fso.CreateTextFile(filePath, True)
    ts.Write content
    ts.Close

    Exit Sub
CleanError:
    ' Continue silently if cleanup fails - don't break the export process
End Sub
