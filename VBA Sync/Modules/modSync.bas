Attribute VB_Name = "modSync"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
#Else
    Private Declare Function GetACP Lib "kernel32" () As Long
#End If

' MIT License
'
' Copyright (c) 2025 Arnaud Lavignolle, Axiom Project Services Pty Ltd
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
Const WORKSHEET_LINE_LIMIT As Long = 200

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
        comp.Export fullPath
        CleanExportedFile fullPath  ' Remove trailing empty lines
        exported(AddSlash(fullPath)) = True
        If comp.Type = vbext_ct_MSForm Then
            exported(AddSlash(subDir & comp.Name & ".frx")) = True
        End If
NextComponent:
    Next

    PruneStaleFiles rootPath, exported

    Debug.Print "VBA Sync: Creating Git helper files..."
    WriteGitAttributes repoPath
    WriteGitIgnore repoPath
    WriteReadme repoPath, wb
    WriteVSCodeSettings repoPath
    
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
        comp.Export fullPath
        CleanExportedFile fullPath  ' Remove trailing empty lines
        exported(AddSlash(fullPath)) = True
        If comp.Type = vbext_ct_MSForm Then
            exported(AddSlash(subDir & comp.Name & ".frx")) = True
        End If
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
    
    Debug.Print "VBA Sync: Export completed successfully!"
    MsgBox "VBA Sync export completed successfully!" & vbCrLf & "Files exported to: " & rootPath, vbInformation, "VBA Sync"
End Sub

'Extract Excel file structure for version control and collaboration
Private Sub ExtractExcelStructure(wb As Workbook, rootPath As String, exported As Object)
    On Error GoTo ExcelStructureError
    
    Dim excelDir As String: excelDir = rootPath & "Excel\"
    EnsureFolder excelDir
    
    ' Create temporary copy of workbook as ZIP
    Dim tempZip As String: tempZip = wb.Path & "\" & wb.Name & ".temp.zip"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile wb.FullName, tempZip
    Debug.Print "VBA Sync: Creating temporary ZIP copy..."
    
    ' Extract Excel structure using Shell
    Dim tempExtract As String: tempExtract = wb.Path & "\temp_excel_extract"
    EnsureFolder tempExtract
    
    ' Use PowerShell to extract ZIP (more reliable than Shell.Application)
    Debug.Print "VBA Sync: Extracting Excel XML structure..."
    Dim psCmd As String
    psCmd = "powershell -Command ""Expand-Archive -Path '" & Replace(tempZip, "'", "''") & "' -DestinationPath '" & Replace(tempExtract, "'", "''") & "' -Force"""
    CreateObject("WScript.Shell").Run psCmd, 0, True
    
    ' Copy key Excel files to src/Excel/
    Debug.Print "VBA Sync: Copying workbook structure..."
    CopyExcelFile tempExtract & "\xl\workbook.xml", excelDir, "workbook.xml", exported
    
    ' Copy table definitions
    Debug.Print "VBA Sync: Copying table definitions..."
    Dim tablesDir As String: tablesDir = excelDir & "tables\"
    If fso.FolderExists(tempExtract & "\xl\tables") Then
        EnsureFolder tablesDir
        Dim tableFile As Object
        For Each tableFile In fso.GetFolder(tempExtract & "\xl\tables").Files
            If LCase(fso.GetExtensionName(tableFile.Name)) = "xml" Then
                CopyExcelFile tableFile.Path, tablesDir, tableFile.Name, exported
            End If
        Next
    End If
    
    ' Copy worksheet structure (limited lines to avoid huge data files)
    Debug.Print "VBA Sync: Copying worksheet schemas..."
    Dim worksheetsDir As String: worksheetsDir = excelDir & "worksheets\"
    If fso.FolderExists(tempExtract & "\xl\worksheets") Then
        EnsureFolder worksheetsDir
        Dim wsFile As Object
        For Each wsFile In fso.GetFolder(tempExtract & "\xl\worksheets").Files
            If LCase(fso.GetExtensionName(wsFile.Name)) = "xml" And wsFile.Name <> "_rels" Then
                CopyExcelFileWithLimit wsFile.Path, worksheetsDir, wsFile.Name, exported, WORKSHEET_LINE_LIMIT
            End If
        Next
    End If
    
    ' Create Excel structure summary
    CreateExcelStructureSummary wb, excelDir, exported
    
    ' Cleanup temporary files
    On Error Resume Next
    fso.DeleteFile tempZip, True
    fso.DeleteFolder tempExtract, True
    On Error GoTo 0
    
    Exit Sub
    
ExcelStructureError:
    ' Cleanup on error
    On Error Resume Next
    If fso.FileExists(tempZip) Then fso.DeleteFile tempZip, True
    If fso.FolderExists(tempExtract) Then fso.DeleteFolder tempExtract, True
    On Error GoTo 0
    ' Continue without Excel structure if extraction fails
End Sub

Private Sub CopyExcelFile(sourcePath As String, destDir As String, fileName As String, exported As Object)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sourcePath) Then
        Dim destPath As String: destPath = destDir & fileName
        fso.CopyFile sourcePath, destPath, True
        exported(AddSlash(destPath)) = True
    End If
End Sub

Private Sub CopyExcelFileWithLimit(sourcePath As String, destDir As String, fileName As String, exported As Object, maxLines As Long)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sourcePath) Then
        Dim sourceText As String
        sourceText = fso.OpenTextFile(sourcePath, 1).ReadAll
        
        ' Limit to first N lines to avoid huge worksheet data files
        Dim Lines As Variant: Lines = Split(sourceText, vbCrLf)
        If UBound(Lines) > maxLines Then
            ReDim Preserve Lines(0 To maxLines)
            sourceText = Join(Lines, vbCrLf) & vbCrLf & _
                        "<!-- Truncated at " & maxLines & " lines by VBA Sync to avoid large files -->"
        End If
        
        Dim destPath As String: destPath = destDir & fileName
        Dim ts As Object: Set ts = fso.CreateTextFile(destPath, True)
        ts.Write sourceText
        ts.Close
        exported(AddSlash(destPath)) = True
    End If
End Sub

Private Sub CreateExcelStructureSummary(wb As Workbook, excelDir As String, exported As Object)
    Dim summaryPath As String: summaryPath = excelDir & "STRUCTURE_SUMMARY.md"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim summary As String
    summary = "# Excel File Structure Summary" & vbCrLf & vbCrLf
    summary = summary & "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
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
    summary = summary & "- `workbook.xml` - Overall workbook structure" & vbCrLf
    summary = summary & "- `tables/*.xml` - Excel table definitions" & vbCrLf
    summary = summary & "- `worksheets/*.xml` - Worksheet schemas (first " & WORKSHEET_LINE_LIMIT & " lines)" & vbCrLf
    
    Dim ts As Object: Set ts = fso.CreateTextFile(summaryPath, True)
    ts.Write summary
    ts.Close
    exported(AddSlash(summaryPath)) = True
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
    Debug.Print "VBA Sync: Removing existing VBA components..."
    Dim vc As Object
    For Each vc In wb.VBProject.VBComponents
        If vc.Type <> vbext_ct_Document Then
            wb.VBProject.VBComponents.Remove vc
        End If
    Next

    '-- iterate expected sub-folders
    Debug.Print "VBA Sync: Importing VBA components..."
    Dim subFolder As Variant, f As String, vbComp As Object, filePath As String
    For Each subFolder In Array("Modules", "ClassModules", "Forms", "Objects", "Misc")
        filePath = rootPath & subFolder & "\"
        If Dir(filePath, vbDirectory) <> "" Then
            f = Dir(filePath & "*.*")
            Do While Len(f) > 0
                If LCase$(Right$(f, 4)) = ".frx" Then GoSub SkipFile 'ignore binary partner

                Dim tgtName As String: tgtName = Split(f, ".")(0)
                Set vbComp = Nothing
                On Error Resume Next
                Set vbComp = wb.VBProject.VBComponents(tgtName)
                On Error GoTo 0

                If vbComp Is Nothing Then
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
    Dim subFolder As Variant, folderPath As String
    For Each subFolder In Array("Modules", "ClassModules", "Forms", "Objects", "Misc", "Excel", "Excel/tables", "Excel/worksheets")
        folderPath = rootPath & subFolder & "\"
        If fso.FolderExists(folderPath) Then
            Dim f As Object
            For Each f In fso.GetFolder(folderPath).Files
                Dim ext As String: ext = LCase$(fso.GetExtensionName(f.Path))
                If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "frx" Or ext = "xml" Or ext = "md" Then
                    If Not exported.Exists(AddSlash(f.Path)) Then
                        On Error Resume Next
                        f.Delete True
                        On Error GoTo 0
                    End If
                End If
            Next
        End If
    Next
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
          "# Excel structure files" & vbCrLf & _
          "*.xml text eol=crlf" & vbCrLf & _
          "*.md text eol=crlf" & vbCrLf & _
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
    If fso.FileExists(iPath) Then
        Debug.Print "VBA Sync: .gitignore already exists, skipping"
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
          "temp_excel_extract/" & vbCrLf & _
          vbCrLf & _
          "# OS cruft" & vbCrLf & _
          "Thumbs.db" & vbCrLf & _
          ".DS_Store" & vbCrLf & _
          vbCrLf & _
          "# IDE/project folders (optional)" & vbCrLf & _
          ".vs/" & vbCrLf & _
          ".idea/" & vbCrLf

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
          "|-- Excel/                # Excel structure (for AI & documentation)" & vbCrLf & _
          "|   |-- workbook.xml      # Workbook metadata & named ranges" & vbCrLf & _
          "|   |-- tables/           # Excel table definitions" & vbCrLf & _
          "|   |-- worksheets/       # Worksheet schemas (truncated for size)" & vbCrLf & _
          "|   `-- STRUCTURE_SUMMARY.md  # Human-readable data model overview" & vbCrLf & _
          "|-- .gitattributes        # Git configuration for VBA files" & vbCrLf & _
          "|-- .gitignore            # Excludes temp files from version control" & vbCrLf & _
          "`-- README.md             # This file" & vbCrLf
    
    txt = txt & "```" & vbCrLf & vbCrLf & _
          "## TOOLS" & vbCrLf & _
          "- **VS Code**: Install VBA extensions for syntax highlighting" & vbCrLf & _
          "- **AI Tools**: GitHub Copilot, Claude, ChatGPT work with your exported code" & vbCrLf & _
          "- **Git**: Use branches (`git checkout -b feature-name`) for development" & vbCrLf
    
    txt = txt & vbCrLf & _
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
    For Each ln In Split(src, vbCrLf)
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

' Remove trailing empty lines from exported VBA files
Private Sub CleanExportedFile(filePath As String)
    On Error GoTo CleanError
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Sub
    
    ' Read the file content
    Dim content As String
    content = fso.OpenTextFile(filePath, 1).ReadAll
    
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
