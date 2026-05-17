Attribute VB_Name = "modExcelExport"
Option Explicit

' MIT License
'
' Copyright (c) 2025 Arnaud Lavignolle

' modExcelExport
' --------------
' Writes a Git-friendly text export of a workbook's structure to the
' Excel/ folder next to the workbook. Read directly from the live Excel
' object model -- no .xlsx/.xlsb unzip, no OOXML parsing (with one
' minor exception: see "Drawing _assets/" below).
'
' Entry point: DoExportExcelStructure(wb, rootPath, exported)
'   called by modSync.ExtractExcelStructure
'
' Output layout:
'   Excel/
'     MANIFEST.json                       sheets, tables, defined names, lambdas
'     lambdas/<Name>.lambda               LAMBDA defined names, deduplicated
'     worksheets/<NN> - <Name>/           NN = sheetId, zero-padded
'       data.tsv                          cell values (resolved, not formulas)
'       formulas.json                     cell formulas, range-collapsed
'       styles.json                       resolved styles, range-merged
'       _meta.json                        tab colour, panes, columns, etc.
'       validations.json                  data validation rules (if any)
'       conditional_formats.json          CF rules (if any)
'       comments.json                     cell comments (if any)
'       tables/<TableName>/
'         definition.json                 schema, columns, calc formulas
'         data.tsv                        clean dataset, headers row 1
'       drawings/
'         shapes.json                     shapes + pictures + OnAction macros
'         _assets/<file>.png|jpg|...      embedded image binaries
'       charts/
'         <ChartName>.png                 PNG snapshot via Chart.Export
'         <ChartName>.json                type, title, anchor, series formulas
'
' Charts:
'   Embedded ChartObjects only. The PNG is the visual source of truth;
'   the JSON answers "what data does this chart use and what kind of
'   chart is it". Visual properties (3D rotation, fill effects, smooth
'   lines, marker styles) are not captured -- look at the PNG for those.
'   Standalone chart sheets (wb.Charts) are not exported.
'
' Drawing _assets/:
'   Picture binaries are copied from xl/media/ inside the .xlsm/.xlsb
'   package. The workbook is unzipped once at export start into a temp
'   folder via Shell.Application COM, the per-sheet drawing rels (small
'   XML files) are parsed to map shape -> media file, and the binaries
'   get copied into each sheet's drawings/_assets/. This is the only
'   place in the export that touches the package contents directly.
'
' Performance:
'   Reads each sheet's UsedRange as a single 2D Variant array (1 COM
'   call) for both .Value and .Formula -- never iterates cells one at
'   a time. Disables ScreenUpdating + manual calc during the run.
'   StatusBar shows per-sheet progress because chart PNG export is the
'   slowest single step.
'
' WriteIfChanged:
'   Every output file is content-compared before writing. Unchanged
'   files are skipped to keep git diff noise to zero on "no-op" exports
'   (e.g. when the user just opens the workbook and saves without edits).

' Per-sheet phase tracker. Updated inside ProcessWorksheet, read by the
' per-sheet error handler in DoExportExcelStructure so log messages can say
' which sub-step raised, not just the sheet name.
Private g_sheetPhase As String

' Cumulative per-phase timings (label -> seconds). Reset at start of each
' DoExportExcelStructure call. Dumped to .export_timing.log on success.
Private g_Timings As Object
Private g_TimingsT0 As Double

Private Sub TimingsInit()
    Set g_Timings = CreateObject("Scripting.Dictionary")
    g_TimingsT0 = Timer
End Sub
Private Sub TimingsAdd(label As String, t0 As Double)
    If g_Timings Is Nothing Then Exit Sub
    Dim dur As Double: dur = Timer - t0
    If dur < 0 Then dur = dur + 86400
    If g_Timings.Exists(label) Then
        g_Timings(label) = CDbl(g_Timings(label)) + dur
    Else
        g_Timings(label) = dur
    End If
End Sub
Private Sub TimingsDump(path As String)
    If g_Timings Is Nothing Then Exit Sub
    On Error Resume Next
    Dim total As Double: total = Timer - g_TimingsT0
    If total < 0 Then total = total + 86400
    Dim keys As Variant: keys = g_Timings.Keys
    Dim vals As Variant: vals = g_Timings.Items
    Dim i As Long, j As Long, n As Long: n = g_Timings.Count
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If CDbl(vals(j)) < CDbl(vals(j + 1)) Then
                Dim tk As Variant: tk = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tk
                Dim tv As Variant: tv = vals(j): vals(j) = vals(j + 1): vals(j + 1) = tv
            End If
        Next
    Next
    Dim fnum As Integer: fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, "label,seconds"
    For i = 0 To n - 1: Print #fnum, keys(i) & "," & Format(CDbl(vals(i)), "0.000"): Next
    Print #fnum, "TOTAL," & Format(total, "0.000")
    Close #fnum
End Sub

'====================  PUBLIC ENTRY  =======================
Public Sub DoExportExcelStructure(wb As Workbook, rootPath As String, exported As Object)
    On Error GoTo Fail

    Dim excelDir As String: excelDir = rootPath & "Excel\"
    EnsureFolder excelDir
    TimingsInit

    ' Clear stale logs from any previous run so the end-of-export summary
    ' reflects only this invocation's failures.
    On Error Resume Next
    Kill excelDir & ".export_sheet_errors.log"
    Kill excelDir & ".export_error.log"
    On Error GoTo 0

    ' Phase tracker (updated at each major step) — surfaced in the failure log
    ' so we know which step blew up when Erl returns 0.
    Dim phase As String: phase = "init"

    ' Restore-on-exit state
    Dim oldStatus As Variant: oldStatus = Application.DisplayStatusBar
    Dim oldStatusText As Variant: oldStatusText = Application.StatusBar
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Dim oldEvents As Boolean: oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    On Error GoTo Fail
    Application.DisplayStatusBar = True

    ' --- Pad width for sheet folder prefix (max(2, len(maxSheetId))) ---
    Dim maxSheetId As Long: maxSheetId = 0
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Index > maxSheetId Then maxSheetId = ws.Index  ' fallback
    Next
    ' Use CodeName-derived sheetIds when possible. Excel exposes the persistent
    ' sheetId via Worksheet.Index only by tab position, not stable. The
    ' OOXML sheetId is what we want — and that lives only in the OOXML.
    ' Workaround: use the sheet's position-stable codeName-derived id via the
    ' VBProject (Sheet1 -> 1, Sheet13 -> 13). For sheets without a codeName
    ' index suffix, fall back to the tab Index.
    Dim sheetIds As Object: Set sheetIds = CreateObject("Scripting.Dictionary")
    Dim wsi As Long: wsi = 0
    Dim assignedIds As Object: Set assignedIds = CreateObject("Scripting.Dictionary")
    maxSheetId = 0
    For Each ws In wb.Worksheets
        wsi = wsi + 1
        Dim sid As Long: sid = SheetIdFromCodeName(ws.CodeName, wsi)
        ' Ensure unique (corrupt workbooks can have dupes)
        Do While assignedIds.Exists(sid)
            sid = sid + 1000
        Loop
        assignedIds(sid) = True
        sheetIds(ws.Name) = sid
        If sid > maxSheetId Then maxSheetId = sid
    Next
    Dim padWidth As Long: padWidth = Len(CStr(maxSheetId))
    If padWidth < 2 Then padWidth = 2

    ' --- LAMBDA extraction (wb.Names + per-sheet wb.Worksheets(i).Names) ---
    phase = "extract-lambdas"
    Dim lambdas As Object: Set lambdas = CreateObject("Scripting.Dictionary")  ' canonical name -> body
    ExtractLambdas wb, lambdas
    phase = "write-lambdas-folder"
    WriteLambdasFolder excelDir & "lambdas\", lambdas, exported

    ' --- Build manifest skeleton ---
    phase = "manifest-skeleton"
    Dim manifest As Object: Set manifest = CreateObject("Scripting.Dictionary")
    manifest("schemaVersion") = 1
    manifest("generatorVersion") = "1.0"

    Dim wbProps As Object: Set wbProps = CreateObject("Scripting.Dictionary")
    wbProps("codeName") = "ThisWorkbook"   ' Excel's stable codeName for the workbook
    On Error Resume Next
    wbProps("date1904") = wb.Date1904
    On Error GoTo Fail
    If Not wbProps.Exists("date1904") Then wbProps("date1904") = False
    wbProps("calcMode") = CalcModeString(Application.Calculation)
    Set manifest("workbook") = wbProps

    ' --- Sheets list (in tab order, with file path strings) ---
    Dim sheetsArr As Object: Set sheetsArr = New Collection
    For Each ws In wb.Worksheets
        Dim sheetMeta As Object: Set sheetMeta = CreateObject("Scripting.Dictionary")
        sheetMeta("sheetId") = sheetIds(ws.Name)
        sheetMeta("name") = ws.Name
        sheetMeta("codeName") = ws.CodeName
        sheetMeta("state") = SheetStateString(ws.Visible)
        sheetMeta("tabColor") = TabColorHex(ws)
        sheetMeta("file") = "worksheets/" & PadLeft(CStr(sheetIds(ws.Name)), padWidth, "0") & " - " & SafeFileName(ws.Name) & ".xml"
        sheetsArr.Add sheetMeta
    Next
    Set manifest("sheets") = sheetsArr

    ' --- Defined names (non-lambda) ---
    phase = "build-defined-names"
    Dim definedNames As Object: Set definedNames = BuildDefinedNames(wb, lambdas)
    Set manifest("definedNames") = definedNames

    ' --- Lambdas list (alphabetical) ---
    Set manifest("lambdas") = SortedKeys(lambdas)

    ' --- Process each worksheet ---
    phase = "prepare-worksheets-dir"
    Dim worksheetsDir As String: worksheetsDir = excelDir & "worksheets\"
    EnsureFolder worksheetsDir

    ' --- Build drawing-image map (sheet name -> { rId -> mediaPath, etc. }) ---
    ' One Shell.Application unzip up front; per-sheet rels parsed into memory.
    ' Returns a dict with keys "tempRoot" (path to extracted xl/ tree) and
    ' "perSheet" (Dictionary keyed by ws.Name -> sheet-level image info).
    phase = "build-drawing-map"
    Application.StatusBar = "VBA Sync: extracting drawing image map..."
    Dim drawingMap As Object: Set drawingMap = Nothing
    Dim tBDM As Double: tBDM = Timer
    On Error Resume Next
    Set drawingMap = BuildDrawingImageMap(wb)
    If Err.Number <> 0 Then
        Debug.Print "VBA Sync: drawing image map build failed: " & Err.Description & " -- continuing without _assets/"
        Err.Clear
        Set drawingMap = Nothing
    End If
    On Error GoTo Fail
    TimingsAdd "build-drawing-map", tBDM

    ' Side-load calc-column formulas from xl/tables/*.xml. The same temp unzip
    ' BuildDrawingImageMap created above is reused. Map is empty if no tables
    ' or if the unzip failed; callers must check.
    phase = "load-ooxml-side-data"
    Dim tableCalcFormulas As Object: Set tableCalcFormulas = CreateObject("Scripting.Dictionary")
    Dim sheetValidations As Object: Set sheetValidations = CreateObject("Scripting.Dictionary")
    Dim sheetViews As Object: Set sheetViews = CreateObject("Scripting.Dictionary")
    If Not drawingMap Is Nothing Then
        If drawingMap.Exists("tempDir") Then
            On Error Resume Next
            Dim tempDirStr As String: tempDirStr = CStr(drawingMap("tempDir"))
            Dim tCalc As Double: tCalc = Timer
            Set tableCalcFormulas = LoadTableCalcFormulas(tempDirStr)
            TimingsAdd "load-table-calc-formulas", tCalc
            Dim sheetMap As Object
            If drawingMap.Exists("sheetNameToTarget") Then
                Set sheetMap = drawingMap("sheetNameToTarget")
            End If
            If sheetMap Is Nothing Then Set sheetMap = LoadWorkbookSheetMap(tempDirStr)
            If sheetMap.Count > 0 Then
                Dim tVal As Double: tVal = Timer
                Set sheetValidations = LoadSheetValidations(tempDirStr, sheetMap)
                TimingsAdd "load-sheet-validations", tVal
                Dim tSV As Double: tSV = Timer
                Set sheetViews = LoadSheetViews(tempDirStr, sheetMap)
                TimingsAdd "load-sheet-views", tSV
            End If
            On Error GoTo Fail
        End If
    End If

    Dim allTableEntries As Object: Set allTableEntries = New Collection
    Dim sheetCounter As Long: sheetCounter = 0
    Dim totalSheets As Long: totalSheets = wb.Worksheets.Count

    For Each ws In wb.Worksheets
        sheetCounter = sheetCounter + 1
        phase = "sheet " & sheetCounter & "/" & totalSheets & " '" & ws.Name & "'"
        Application.StatusBar = "VBA Sync: processing sheet " & sheetCounter & " of " & totalSheets & " (" & ws.Name & ")..."
        On Error Resume Next
        Dim sheetEntries As Object: Set sheetEntries = Nothing
        Dim sheetDrawing As Object: Set sheetDrawing = Nothing
        If Not drawingMap Is Nothing Then
            If drawingMap.Exists("perSheet") Then
                If drawingMap("perSheet").Exists(ws.Name) Then
                    Set sheetDrawing = drawingMap("perSheet")(ws.Name)
                End If
            End If
        End If
        Dim wsValidations As Object: Set wsValidations = Nothing
        If sheetValidations.Exists(ws.Name) Then Set wsValidations = sheetValidations(ws.Name)
        Dim wsView As Object: Set wsView = Nothing
        If sheetViews.Exists(ws.Name) Then Set wsView = sheetViews(ws.Name)
        g_sheetPhase = "init"
        Dim tPS As Double: tPS = Timer
        Set sheetEntries = ProcessWorksheet(ws, sheetIds(ws.Name), padWidth, worksheetsDir, exported, tableCalcFormulas, sheetDrawing, wsValidations, wsView)
        TimingsAdd "per-sheet-loop", tPS
        If Err.Number <> 0 Then
            Dim sErr As String
            sErr = "VBA Sync: sheet '" & ws.Name & "' failed in '" & g_sheetPhase & "': " & _
                   Err.Description & " (Err.Number=" & Err.Number & ")"
            Debug.Print sErr
            Dim sLogPath As String: sLogPath = excelDir & ".export_sheet_errors.log"
            Dim sFnum As Integer: sFnum = FreeFile
            Open sLogPath For Append As #sFnum
            Print #sFnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & sErr
            Close #sFnum
            Err.Clear
        End If
        On Error GoTo Fail
        If Not sheetEntries Is Nothing Then
            Dim te As Variant
            For Each te In sheetEntries
                allTableEntries.Add te
            Next
        End If
    Next

    ' --- Cleanup temp drawing-extract folder ---
    If Not drawingMap Is Nothing Then
        If drawingMap.Exists("tempDir") Then
            On Error Resume Next
            CleanupTempDir CStr(drawingMap("tempDir"))
            On Error GoTo Fail
        End If
    End If

    ' Aggregate tables (sorted by sheet, then name)
    phase = "aggregate-tables"
    Set manifest("tables") = SortTableEntries(allTableEntries)

    ' Write MANIFEST.json LAST (tables array now populated)
    phase = "write-manifest"
    WriteIfChanged excelDir & "MANIFEST.json", JsonPretty(manifest), exported, False
    TimingsDump excelDir & ".export_timing.log"

    ' Post-export sanity check: every per-sheet folder must have _meta.json.
    ' Empty folders mean the sheet was processed but failed to write — either
    ' the per-sheet error handler already logged it, or it slipped through
    ' (rare but possible). Append any unaccounted-for empties to the same log
    ' so the user's failure summary catches them.
    phase = "sanity-check"
    SanityCheckSheetFolders worksheetsDir, excelDir
    phase = "done"

Done:
    On Error Resume Next
    Application.StatusBar = oldStatusText
    Application.DisplayStatusBar = oldStatus
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Exit Sub

Fail:
    Dim failMsg As String
    failMsg = "VBA Sync: ExcelExport failed in phase '" & phase & "': " & _
              Err.Description & " (Err.Number=" & Err.Number & ", Erl=" & Erl & ")"
    Debug.Print failMsg
    On Error Resume Next
    Dim logPath As String: logPath = excelDir & ".export_error.log"
    Dim fnum As Integer: fnum = FreeFile
    Open logPath For Output As #fnum
    Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & failMsg
    Close #fnum
    On Error GoTo 0
    Resume Done
End Sub

'====================  WORKSHEET PROCESSING  =================
' Processes a single worksheet end-to-end. Returns an ArrayList of table-manifest
' entries for the caller to aggregate into MANIFEST.tables.
Private Function ProcessWorksheet(ws As Worksheet, sheetId As Long, padWidth As Long, _
                                  worksheetsDir As String, exported As Object, _
                                  tableCalcFormulas As Object, _
                                  Optional sheetDrawing As Object, _
                                  Optional ooxmlValidations As Object, _
                                  Optional ooxmlSheetView As Object) As Object
    Dim safeName As String: safeName = SafeFileName(ws.Name)
    Dim padded As String: padded = PadLeft(CStr(sheetId), padWidth, "0")
    Dim sheetFolder As String: sheetFolder = worksheetsDir & padded & " - " & safeName & "\"
    EnsureFolder sheetFolder

    Dim tableEntries As Object: Set tableEntries = New Collection
    Dim tableRanges As Object: Set tableRanges = New Collection
    ' Each tableRange: dict { name, safeName, startCol, startRow, endCol, endRow, folder, listObject, ref }

    g_sheetPhase = "list-objects-enum"
    Dim lo As ListObject
    Dim tablesDir As String: tablesDir = sheetFolder & "tables\"
    For Each lo In ws.ListObjects
        Dim rng As Range: Set rng = lo.Range
        Dim tName As String: tName = lo.Name
        Dim tSafe As String: tSafe = SafeFileName(tName)
        Dim tFolder As String: tFolder = tablesDir & tSafe & "\"

        Dim tInfo As Object: Set tInfo = CreateObject("Scripting.Dictionary")
        tInfo("name") = tName
        tInfo("safeName") = tSafe
        tInfo("startCol") = rng.Column
        tInfo("startRow") = rng.Row
        tInfo("endCol") = rng.Column + rng.Columns.Count - 1
        tInfo("endRow") = rng.Row + rng.Rows.Count - 1
        tInfo("folder") = tFolder
        tInfo("ref") = ColLetters(rng.Column) & CStr(rng.Row) & ":" & _
                       ColLetters(rng.Column + rng.Columns.Count - 1) & CStr(rng.Row + rng.Rows.Count - 1)
        Set tInfo("listObject") = lo
        tableRanges.Add tInfo

        Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
        entry("name") = tName
        On Error Resume Next
        entry("displayName") = lo.DisplayName
        If entry("displayName") = "" Then entry("displayName") = tName
        On Error GoTo 0
        entry("sheet") = ws.Name
        entry("ref") = tInfo("ref")
        entry("folder") = "worksheets/" & padded & " - " & safeName & "/tables/" & tSafe & "/"
        tableEntries.Add entry
    Next

    ' --- Walk cells once via 2D arrays from UsedRange ---
    g_sheetPhase = "used-range"
    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0

    Dim cellValues As Object: Set cellValues = CreateObject("Scripting.Dictionary")     ' "col,row" -> value
    Dim cellFormulas As Object: Set cellFormulas = CreateObject("Scripting.Dictionary") ' "col,row" -> "=..."
    Dim cellStyles As Object: Set cellStyles = CreateObject("Scripting.Dictionary")     ' "col,row" -> style dict

    If Not ur Is Nothing Then
        g_sheetPhase = "walk-used-range"
        WalkUsedRange ur, cellValues, cellFormulas, cellStyles
    End If

    ' --- Carve out table cells (so they don't appear in sheet data.tsv) ---
    g_sheetPhase = "tables-carve"
    Dim tableMarkers As Object: Set tableMarkers = CreateObject("Scripting.Dictionary")
    Dim ti As Variant
    For Each ti In tableRanges
        Dim t As Object: Set t = ti
        Dim r As Long, c As Long
        Dim tableValues As Object: Set tableValues = CreateObject("Scripting.Dictionary")
        Dim tableFormulas As Object: Set tableFormulas = CreateObject("Scripting.Dictionary")
        For r = t("startRow") To t("endRow")
            For c = t("startCol") To t("endCol")
                Dim k As String: k = CStr(c) & "," & CStr(r)
                If cellValues.Exists(k) Then
                    tableValues(k) = cellValues(k)
                    cellValues.Remove k
                End If
                If cellFormulas.Exists(k) Then
                    tableFormulas(k) = cellFormulas(k)
                    cellFormulas.Remove k
                End If
            Next
        Next
        Dim markerKey As String: markerKey = CStr(t("startCol")) & "," & CStr(t("startRow"))
        tableMarkers(markerKey) = "[table:" & t("name") & " ref=" & t("ref") & " -> tables/" & t("safeName") & "/]"

        EnsureFolder t("folder")
        SaveTableDataTsv t("folder") & "data.tsv", tableValues, t, exported
        SaveTableDefinitionJson t("folder") & "definition.json", t, tableFormulas, tableValues, tableCalcFormulas, exported
    Next

    ' Merge marker rows back into sheet's values
    Dim mk As Variant
    For Each mk In tableMarkers.Keys
        cellValues(mk) = tableMarkers(mk)
    Next

    ' --- Write sheet-level files ---
    Dim tW As Double
    g_sheetPhase = "save-meta": tW = Timer:        SaveMetaJson sheetFolder & "_meta.json", ws, sheetId, tableRanges, exported, ooxmlSheetView: TimingsAdd "ws:save-meta", tW
    g_sheetPhase = "save-data-tsv": tW = Timer:    SaveDataTsv sheetFolder & "data.tsv", cellValues, exported: TimingsAdd "ws:save-data-tsv", tW
    g_sheetPhase = "save-formulas": tW = Timer:    SaveFormulasJson sheetFolder & "formulas.json", cellFormulas, exported: TimingsAdd "ws:save-formulas", tW
    g_sheetPhase = "save-styles": tW = Timer:      SaveStylesJson sheetFolder & "styles.json", cellStyles, exported: TimingsAdd "ws:save-styles", tW
    g_sheetPhase = "save-validations": tW = Timer: SaveValidationsJson sheetFolder & "validations.json", ws, exported, ooxmlValidations: TimingsAdd "ws:save-validations", tW
    g_sheetPhase = "save-cf": tW = Timer:          SaveConditionalFormatsJson sheetFolder & "conditional_formats.json", ws, exported: TimingsAdd "ws:save-cf", tW
    g_sheetPhase = "save-comments": tW = Timer:    SaveCommentsJson sheetFolder & "comments.json", ws, exported: TimingsAdd "ws:save-comments", tW
    g_sheetPhase = "save-drawings": tW = Timer:    SaveDrawingsJson sheetFolder & "drawings\", ws, exported, sheetDrawing: TimingsAdd "ws:save-drawings", tW
    g_sheetPhase = "export-charts": tW = Timer:    ExportCharts sheetFolder & "charts\", ws, exported: TimingsAdd "ws:export-charts", tW

    g_sheetPhase = "done"
    Set ProcessWorksheet = tableEntries
End Function

'====================  SANITY CHECK  =========================
' Walks worksheetsDir, flags any per-sheet folder missing _meta.json.
' Appends entries to .export_sheet_errors.log so the same MsgBox summary
' downstream picks them up alongside in-process failures.
Private Sub SanityCheckSheetFolders(worksheetsDir As String, excelDir As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(worksheetsDir) Then Exit Sub
    Dim folder As Object: Set folder = fso.GetFolder(worksheetsDir)
    Dim sub_ As Object
    Dim logPath As String: logPath = excelDir & ".export_sheet_errors.log"
    For Each sub_ In folder.SubFolders
        If Not fso.FileExists(sub_.path & "\_meta.json") Then
            Dim fnum As Integer: fnum = FreeFile
            Open logPath For Append As #fnum
            Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & _
                " VBA Sync: sheet folder '" & sub_.Name & _
                "' has no _meta.json (sanity check)"
            Close #fnum
        End If
    Next
End Sub

'====================  USEDRANGE WALKER  =====================
' Reads UsedRange.Value2 and UsedRange.Formula in two big COM calls.
' Then walks the 2D arrays to populate cellValues, cellFormulas, cellStyles.
' Styles are read cell-by-cell only for non-default cells (still need to touch
' COM per cell for Font/Interior/Border, but only where the formula or value
' indicates a non-empty cell — empty cells skip the style probe).
Private Sub WalkUsedRange(ur As Range, ByRef cellValues As Object, _
                          ByRef cellFormulas As Object, ByRef cellStyles As Object)
    Dim baseCol As Long: baseCol = ur.Column
    Dim baseRow As Long: baseRow = ur.Row
    Dim nRows As Long: nRows = ur.Rows.Count
    Dim nCols As Long: nCols = ur.Columns.Count

    Dim values As Variant
    Dim formulas As Variant
    Dim ws As Worksheet: Set ws = ur.Worksheet

    If nRows = 1 And nCols = 1 Then
        ' Single-cell UsedRange: VBA returns a scalar, not a 2D array.
        ReDim values(1 To 1, 1 To 1)
        ReDim formulas(1 To 1, 1 To 1)
        values(1, 1) = ur.Value2
        On Error Resume Next
        formulas(1, 1) = ur.Formula
        On Error GoTo 0
    Else
        ' Bulk-read both arrays in one COM call each. Range.Formula raises
        ' 1004 when the range contains a multi-cell array formula (CSE) --
        ' Excel forbids reading .Formula for those at range level. Fall back
        ' to a per-cell loop with FormulaArray when the bulk read fails.
        values = ur.Value2
        On Error Resume Next
        formulas = ur.Formula
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            ReDim formulas(1 To nRows, 1 To nCols)
            Dim ri As Long, ci As Long
            For ri = 1 To nRows
                For ci = 1 To nCols
                    Dim cellRef As Range: Set cellRef = ur.Cells(ri, ci)
                    Dim fStr As String: fStr = ""
                    On Error Resume Next
                    If cellRef.HasArray Then
                        fStr = CStr(cellRef.FormulaArray)
                    Else
                        fStr = CStr(cellRef.Formula)
                    End If
                    On Error GoTo 0
                    formulas(ri, ci) = fStr
                Next
            Next
        End If
        On Error GoTo 0
    End If

    Dim i As Long, j As Long
    Dim absCol As Long, absRow As Long
    Dim v As Variant, f As Variant
    Dim k As String

    For i = 1 To nRows
        absRow = baseRow + i - 1
        For j = 1 To nCols
            absCol = baseCol + j - 1
            v = values(i, j)
            f = formulas(i, j)

            ' Skip wholly empty cells
            Dim hasValue As Boolean: hasValue = False
            Dim hasFormula As Boolean: hasFormula = False
            If Not IsEmpty(v) And Not IsNull(v) Then
                If VarType(v) <> vbString Then
                    hasValue = True
                ElseIf CStr(v) <> "" Then
                    hasValue = True
                End If
            End If
            If Not IsEmpty(f) And Not IsNull(f) Then
                If VarType(f) = vbString Then
                    If Len(CStr(f)) > 0 Then
                        If Left$(CStr(f), 1) = "=" Then
                            hasFormula = True
                        End If
                    End If
                End If
            End If

            If hasValue Or hasFormula Then
                k = CStr(absCol) & "," & CStr(absRow)
                If hasValue Then
                    cellValues(k) = FormatCellValue(v)
                End If
                If hasFormula Then
                    cellFormulas(k) = CStr(f)
                End If

                ' cellStyles is populated via OOXML cellXfs lookup (TODO);
                ' per-cell COM probing is intentionally not done here.
            End If
        Next
    Next
End Sub

' Coerce a VBA cell value (Variant from .Value2) to its TSV-emit form.
Private Function FormatCellValue(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then FormatCellValue = "": Exit Function
    Select Case VarType(v)
        Case vbBoolean
            FormatCellValue = IIf(CBool(v), "TRUE", "FALSE")
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbInteger, vbLong
            ' Emit as the textual rep used by Excel sharedStrings/values:
            ' integers without decimals, floats with full precision.
            ' R-conversion gives ~15 sig figs (matches Excel internal).
            Dim s As String: s = CStr(v)
            ' VBA sometimes returns "1.5" with locale-specific decimal; force "."
            If InStr(s, ",") > 0 Then s = Replace(s, ",", ".")
            FormatCellValue = s
        Case vbDate
            ' Numbers in sharedStrings/values are stored as serials; emit serial
            FormatCellValue = CStr(CDbl(CDate(v)))
        Case vbError
            ' Error cell: emit the error string
            FormatCellValue = CStr(v)
        Case Else
            FormatCellValue = CStr(v)
    End Select
End Function

'====================  STYLE RESOLUTION  =====================
' Reads a cell's resolved formatting. Emits only non-default attrs.
' Defaults: name=Calibri, size=11, bold=false, italic=false, etc.
Private Function ResolveCellStyle(cell As Range) As Object
    Dim r As Object: Set r = CreateObject("Scripting.Dictionary")

    ' --- Number format (skip "General") ---
    Dim nf As String
    On Error Resume Next
    nf = cell.NumberFormat
    On Error GoTo 0
    If Len(nf) > 0 And nf <> "General" Then
        r("numberFormat") = nf
    End If

    ' --- Font ---
    Dim fontDict As Object: Set fontDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim fName As String: fName = cell.Font.Name
    If Len(fName) > 0 Then fontDict("name") = fName
    Dim fSize As Variant: fSize = cell.Font.Size
    If Not IsNull(fSize) And Not IsEmpty(fSize) Then
        ' Only emit size if differs from 11 — but we ALWAYS emit when emitting font
        ' (matches PS which emits {name, size} on any non-default cellXf reference).
        fontDict("size") = CDbl(fSize)
    End If
    If cell.Font.Bold = True Then fontDict("bold") = True
    If cell.Font.Italic = True Then fontDict("italic") = True
    If cell.Font.Underline <> xlUnderlineStyleNone And cell.Font.Underline <> Empty Then
        fontDict("underline") = True
    End If
    If cell.Font.Strikethrough = True Then fontDict("strike") = True
    ' Font color (only if not automatic / default black)
    Dim fColor As Variant: fColor = cell.Font.Color
    If Not IsNull(fColor) And Not IsEmpty(fColor) Then
        If CLng(fColor) <> 0 And cell.Font.ColorIndex <> xlColorIndexAutomatic Then
            fontDict("color") = OleColorToHex(CLng(fColor))
        End If
    End If
    On Error GoTo 0
    ' Heuristic: emit font only if a *meaningful* override exists.
    ' "Calibri size 11" is the workbook default; emit only if anything else is set.
    Dim emitFont As Boolean: emitFont = False
    If fontDict.Exists("bold") Or fontDict.Exists("italic") Or fontDict.Exists("underline") Or _
       fontDict.Exists("strike") Or fontDict.Exists("color") Then
        emitFont = True
    End If
    ' Match PS behaviour: PS emits font on EVERY cellXf that references a non-zero
    ' fontId (which means anything other than the workbook's default font 0). We
    ' can't see Excel's fontId from VBA, but a reasonable proxy is: any cell whose
    ' rendered font has size != 11 OR name != Calibri also gets emit. Otherwise
    ' we'd over-suppress. Be slightly conservative.
    If Not emitFont Then
        If fontDict.Exists("name") Then
            If LCase(fontDict("name")) <> "calibri" Then emitFont = True
        End If
        If fontDict.Exists("size") Then
            If fontDict("size") <> 11 Then emitFont = True
        End If
    End If
    If emitFont And fontDict.Count > 0 Then
        Set r("font") = fontDict
    End If

    ' --- Fill / Interior ---
    On Error Resume Next
    Dim iColor As Variant: iColor = cell.Interior.Color
    Dim iPattern As Variant: iPattern = cell.Interior.Pattern
    On Error GoTo 0
    If Not IsNull(iColor) And Not IsEmpty(iColor) Then
        If CLng(iColor) <> 16777215 And cell.Interior.ColorIndex <> xlColorIndexNone Then
            ' Not pure white (= no fill)
            r("fill") = OleColorToHex(CLng(iColor))
        End If
    End If

    ' --- Border ---
    Dim borderDict As Object: Set borderDict = CreateObject("Scripting.Dictionary")
    Dim sides As Variant: sides = Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
    Dim sideNames As Variant: sideNames = Array("left", "right", "top", "bottom")
    Dim si As Long
    For si = LBound(sides) To UBound(sides)
        On Error Resume Next
        Dim b As Border: Set b = cell.Borders(sides(si))
        If Not b Is Nothing Then
            If b.LineStyle <> xlLineStyleNone And b.LineStyle <> Empty Then
                Dim bd As Object: Set bd = CreateObject("Scripting.Dictionary")
                bd("style") = BorderStyleString(b.LineStyle, b.Weight)
                If b.Color <> 0 And b.ColorIndex <> xlColorIndexAutomatic Then
                    bd("color") = OleColorToHex(CLng(b.Color))
                End If
                Set borderDict(sideNames(si)) = bd
            End If
        End If
        On Error GoTo 0
    Next
    If borderDict.Count > 0 Then
        Set r("border") = borderDict
    End If

    Set ResolveCellStyle = r
End Function

Private Function BorderStyleString(ls As Variant, w As Variant) As String
    Select Case ls
        Case 1    ' xlContinuous
            Select Case w
                Case 1: BorderStyleString = "hair"    ' xlHairline
                Case 2: BorderStyleString = "thin"    ' xlThin
                Case -4138: BorderStyleString = "medium"    ' xlMedium
                Case 4: BorderStyleString = "thick"    ' xlThick
                Case Else: BorderStyleString = "thin"
            End Select
        Case -4115: BorderStyleString = "dashed"    ' xlDash
        Case -4118: BorderStyleString = "dotted"    ' xlDot
        Case -4119: BorderStyleString = "double"    ' xlDouble
        Case 13: BorderStyleString = "slantDashDot"    ' xlSlantDashDot
        Case Else: BorderStyleString = "thin"
    End Select
End Function

Private Function OleColorToHex(c As Long) As String
    ' OLE color is &H00BBGGRR; convert to "#RRGGBB"
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&
    g = (c \ 256) And &HFF&
    b = (c \ 65536) And &HFF&
    OleColorToHex = "#" & UCase(Right$("00" & Hex(r), 2) & Right$("00" & Hex(g), 2) & Right$("00" & Hex(b), 2))
End Function

'====================  RANGE MERGING  ========================
' Greedy rectangle merging. Cells keyed "col,row". Returns ArrayList of
' { range: "A1:B5", value: <original value> } ordered top-left to bottom-right.
Private Function CompressCellsToRanges(cells As Object) As Object
    ' Compress a sparse "col,row" -> value Dictionary into a Collection of
    ' { range:"A1:B3", value:... } entries, merging adjacent cells with
    ' identical values into rectangles.
    Dim out As Object: Set out = New Collection
    If cells.Count = 0 Then Set CompressCellsToRanges = out: Exit Function

    Dim keysArr As Variant: keysArr = cells.Keys
    Dim itemsArr As Variant: itemsArr = cells.Items
    Dim nKeys As Long: nKeys = cells.Count
    Dim ccs() As Long: ReDim ccs(0 To nKeys - 1)
    Dim rrs() As Long: ReDim rrs(0 To nKeys - 1)
    Dim minRow As Long: minRow = 2147483647
    Dim maxRow As Long: maxRow = 0
    Dim minCol As Long: minCol = 2147483647
    Dim maxCol As Long: maxCol = 0
    Dim ki As Long
    For ki = 0 To nKeys - 1
        Dim parts() As String: parts = Split(CStr(keysArr(ki)), ",")
        Dim cc As Long: cc = CLng(parts(0))
        Dim rr As Long: rr = CLng(parts(1))
        ccs(ki) = cc
        rrs(ki) = rr
        If rr < minRow Then minRow = rr
        If rr > maxRow Then maxRow = rr
        If cc < minCol Then minCol = cc
        If cc > maxCol Then maxCol = cc
    Next

    Dim nRows As Long: nRows = maxRow - minRow + 1
    Dim nCols As Long: nCols = maxCol - minCol + 1

    Dim sv() As String: ReDim sv(0 To nRows - 1, 0 To nCols - 1)
    Dim vals() As Variant: ReDim vals(0 To nRows - 1, 0 To nCols - 1)
    For ki = 0 To nKeys - 1
        Dim ri As Long: ri = rrs(ki) - minRow
        Dim ci As Long: ci = ccs(ki) - minCol
        If IsObject(itemsArr(ki)) Then
            sv(ri, ci) = JsonConverter.ConvertToJson(itemsArr(ki))
            Set vals(ri, ci) = itemsArr(ki)
        Else
            sv(ri, ci) = "S:" & CStr(itemsArr(ki))
            vals(ri, ci) = itemsArr(ki)
        End If
    Next

    ' Pass 3: greedy rectangle merge. visited mask is a 2D Boolean array
    ' (much cheaper than Dictionary).
    Dim vis() As Boolean: ReDim vis(0 To nRows - 1, 0 To nCols - 1)
    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        For c = 0 To nCols - 1
            If vis(r, c) Then GoTo NextCell
            If Len(sv(r, c)) = 0 Then GoTo NextCell

            Dim curVal As String: curVal = sv(r, c)

            ' Extend right
            Dim endCol As Long: endCol = c
            Do While endCol + 1 <= nCols - 1
                If vis(r, endCol + 1) Then Exit Do
                If sv(r, endCol + 1) <> curVal Then Exit Do
                endCol = endCol + 1
            Loop

            ' Extend down (entire row from c..endCol must match)
            Dim endRow As Long: endRow = r
            Dim canExtend As Boolean: canExtend = True
            Do While endRow + 1 <= nRows - 1 And canExtend
                Dim cc3 As Long
                For cc3 = c To endCol
                    If vis(endRow + 1, cc3) Then canExtend = False: Exit For
                    If sv(endRow + 1, cc3) <> curVal Then canExtend = False: Exit For
                Next
                If canExtend Then endRow = endRow + 1
            Loop

            ' Mark visited
            Dim rr3 As Long, cc4 As Long
            For rr3 = r To endRow
                For cc4 = c To endCol
                    vis(rr3, cc4) = True
                Next
            Next

            ' Emit (translate back to absolute col,row). Each entry is a
            ' 2-element Variant array (range, value) — cheaper than a
            ' Scripting.Dictionary when there can be tens of thousands.
            Dim absC As Long: absC = c + minCol
            Dim absR As Long: absR = r + minRow
            Dim absEC As Long: absEC = endCol + minCol
            Dim absER As Long: absER = endRow + minRow
            Dim startRef As String: startRef = ColLetters(absC) & CStr(absR)
            Dim rangeStr As String
            If endCol = c And endRow = r Then
                rangeStr = startRef
            Else
                rangeStr = startRef & ":" & ColLetters(absEC) & CStr(absER)
            End If

            ' Top-left value lives in the parallel `vals` grid at (r, c).
            Dim entryArr(0 To 1) As Variant
            entryArr(0) = rangeStr
            If IsObject(vals(r, c)) Then
                Set entryArr(1) = vals(r, c)
            Else
                entryArr(1) = vals(r, c)
            End If
            out.Add entryArr
NextCell:
        Next
    Next

    Set CompressCellsToRanges = out
End Function

'====================  FILE WRITERS  =========================

Private Sub SaveDataTsv(path As String, cells As Object, exported As Object)
    Dim sb As Object: Set sb = NewStringBuilder()

    If cells.Count = 0 Then
        sbAppend sb, "Row" & vbCrLf
        WriteIfChanged path, sbToString(sb), exported, True
        Exit Sub
    End If

    ' Determine extents
    Dim maxCol As Long: maxCol = 0
    Dim k As Variant
    Dim byRow As Object: Set byRow = CreateObject("Scripting.Dictionary")
    For Each k In cells.Keys
        Dim parts() As String: parts = Split(CStr(k), ",")
        Dim c As Long: c = CLng(parts(0))
        Dim r As Long: r = CLng(parts(1))
        If c > maxCol Then maxCol = c
        If Not byRow.Exists(r) Then
            Dim rowDict As Object: Set rowDict = CreateObject("Scripting.Dictionary")
            Set byRow(r) = rowDict
        End If
        byRow(r)(c) = cells(CStr(k))
    Next

    ' Header: "Row\tA\tB\t..."
    sbAppend sb, "Row"
    Dim ci As Long
    For ci = 1 To maxCol
        sbAppend sb, vbTab & ColLetters(ci)
    Next
    sbAppend sb, vbCrLf

    ' Sort rows ascending
    Dim sortedRows As Object: Set sortedRows = SortedNumericKeys(byRow)
    Dim ri As Variant
    For Each ri In sortedRows
        sbAppend sb, CStr(ri)
        Dim rowCells As Object: Set rowCells = byRow(CLng(ri))
        For ci = 1 To maxCol
            sbAppend sb, vbTab
            If rowCells.Exists(ci) Then
                sbAppend sb, TsvEscape(CStr(rowCells(ci)))
            End If
        Next
        sbAppend sb, vbCrLf
    Next

    WriteIfChanged path, sbToString(sb), exported, True
End Sub

Private Sub SaveFormulasJson(path As String, cells As Object, exported As Object)
    If cells.Count = 0 Then Exit Sub

    ' CompressCellsToRanges returns entries in row-major order; no sort needed.
    Dim sortedMerged As Object: Set sortedMerged = CompressCellsToRanges(cells)

    ' Each entry is a 2-element Variant array (range:String, value:Variant).
    Dim ranges As Object: Set ranges = CreateObject("Scripting.Dictionary")
    Dim cellsOnly As Object: Set cellsOnly = CreateObject("Scripting.Dictionary")
    Dim entry As Variant
    For Each entry In sortedMerged
        Dim rng As String: rng = CStr(entry(0))
        If InStr(rng, ":") > 0 Then
            ranges(rng) = entry(1)
        Else
            cellsOnly(rng) = entry(1)
        End If
    Next

    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    If ranges.Count > 0 Then Set obj("ranges") = ranges
    If cellsOnly.Count > 0 Then Set obj("cells") = cellsOnly
    ' (No sharedFormulas section: VBA returns the resolved formula per cell;
    ' shared/si tracking is an OOXML storage detail we don't surface.)

    If obj.Count = 0 Then Exit Sub
    WriteIfChanged path, JsonPretty(obj), exported, False
End Sub

Private Sub SaveStylesJson(path As String, cells As Object, exported As Object)
    If cells.Count = 0 Then Exit Sub

    ' CompressCellsToRanges returns entries in row-major order; no sort needed.
    Dim sortedMerged As Object: Set sortedMerged = CompressCellsToRanges(cells)

    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    Dim entry As Variant
    For Each entry In sortedMerged
        Set obj(CStr(entry(0))) = entry(1)
    Next
    If obj.Count = 0 Then Exit Sub
    WriteIfChanged path, JsonPretty(obj), exported, False
End Sub

Private Sub SaveMetaJson(path As String, ws As Worksheet, sheetId As Long, _
                         tableRanges As Object, exported As Object, _
                         Optional ooxmlSheetView As Object)
    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    obj("sheetId") = sheetId
    obj("name") = ws.Name
    obj("codeName") = ws.CodeName
    obj("state") = SheetStateString(ws.Visible)
    obj("tabColor") = TabColorHex(ws)
    If IsNull(obj("tabColor")) Then obj("tabColor") = Null   ' explicit JSON null

    ' showGridLines + frozenPanes come from OOXML <sheetView> in the
    ' extracted temp dir. Window.DisplayGridlines / Window.FreezePanes
    ' only reflect the active sheet, so OOXML is the only reliable
    ' per-sheet source.
    If Not ooxmlSheetView Is Nothing Then
        If ooxmlSheetView.Exists("showGridLines") Then obj("showGridLines") = ooxmlSheetView("showGridLines")
        If ooxmlSheetView.Exists("frozenPanes") Then Set obj("frozenPanes") = ooxmlSheetView("frozenPanes")
    End If

    ' Columns (custom widths)
    Dim cols As Object: Set cols = New Collection
    Dim colRng As Range
    Dim lastCol As Long
    On Error Resume Next
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1

    ' Build column entries by scanning for non-default widths
    Dim colIdx As Long
    Dim defaultWidth As Double: defaultWidth = ws.StandardWidth
    Dim runStart As Long: runStart = 0
    Dim runWidth As Double: runWidth = 0
    Dim runHidden As Boolean: runHidden = False
    Dim runCustom As Boolean: runCustom = False

    For colIdx = 1 To lastCol
        Dim curWidth As Double, curHidden As Boolean
        curWidth = ws.Columns(colIdx).ColumnWidth
        curHidden = ws.Columns(colIdx).Hidden
        Dim curCustom As Boolean: curCustom = (Abs(curWidth - defaultWidth) > 0.0001)

        If runStart = 0 Then
            ' Start a run if non-default
            If curCustom Or curHidden Then
                runStart = colIdx
                runWidth = curWidth
                runHidden = curHidden
                runCustom = curCustom
            End If
        Else
            If Abs(curWidth - runWidth) > 0.0001 Or curHidden <> runHidden Then
                ' Close current run
                cols.Add MakeColEntry(runStart, colIdx - 1, runWidth, runCustom, runHidden)
                ' Start new run if still non-default
                If curCustom Or curHidden Then
                    runStart = colIdx
                    runWidth = curWidth
                    runHidden = curHidden
                    runCustom = curCustom
                Else
                    runStart = 0
                End If
            End If
        End If
    Next
    If runStart > 0 Then
        cols.Add MakeColEntry(runStart, lastCol, runWidth, runCustom, runHidden)
    End If

    If cols.Count > 0 Then Set obj("columns") = cols

    ' Merged cells. Sheet-level fast-path: Range.MergeCells returns False if
    ' nothing on the sheet is merged (one COM call). Only fall through to the
    ' per-cell scan when the result is Null (mixed) or True (rare: whole sheet).
    Dim mergedCells As Object: Set mergedCells = New Collection
    On Error Resume Next
    Dim usedR As Range: Set usedR = ws.UsedRange
    If Not usedR Is Nothing Then
        Dim mergeFlag As Variant: mergeFlag = usedR.MergeCells
        If IsNull(mergeFlag) Or mergeFlag = True Then
            Dim mergeArea As Range
            Dim seenMerges As Object: Set seenMerges = CreateObject("Scripting.Dictionary")
            For Each mergeArea In usedR.Cells
                If mergeArea.MergeCells Then
                    Dim mAddr As String: mAddr = mergeArea.MergeArea.Address(False, False)
                    If Not seenMerges.Exists(mAddr) Then
                        seenMerges(mAddr) = True
                        mergedCells.Add mAddr
                    End If
                End If
            Next
        End If
    End If
    On Error GoTo 0
    If mergedCells.Count > 0 Then
        SortStringCollection mergedCells
        Set obj("mergedCells") = mergedCells
    End If

    ' Hosted tables
    If tableRanges.Count > 0 Then
        Dim hostedTables As Object: Set hostedTables = New Collection
        Dim ti As Variant
        For Each ti In tableRanges
            Dim ht As Object: Set ht = CreateObject("Scripting.Dictionary")
            ht("name") = ti("name")
            ht("ref") = ti("ref")
            hostedTables.Add ht
        Next
        Set obj("tables") = hostedTables
    End If

    WriteIfChanged path, JsonPretty(obj), exported, False
End Sub

Private Function MakeColEntry(minCol As Long, maxCol As Long, width As Double, custom As Boolean, hidden As Boolean) As Object
    Dim e As Object: Set e = CreateObject("Scripting.Dictionary")
    e("min") = minCol
    e("max") = maxCol
    e("width") = width
    If custom Then e("customWidth") = True
    If hidden Then e("hidden") = True
    Set MakeColEntry = e
End Function

Private Sub SaveValidationsJson(path As String, ws As Worksheet, exported As Object, _
                                Optional ooxmlValidations As Object)
    ' Validations come from OOXML (LoadSheetValidations parses
    ' xl/worksheets/sheetN.xml in the synchronously-extracted temp dir).
    ' No VBA SpecialCells fallback -- it returns 1004 in some COM contexts
    ' even when validations exist, which silently dropped data.
    On Error GoTo Done
    If ooxmlValidations Is Nothing Then Exit Sub
    If ooxmlValidations.Count = 0 Then Exit Sub

    Dim ooxObj As Object: Set ooxObj = CreateObject("Scripting.Dictionary")
    Set ooxObj("validations") = ooxmlValidations
    WriteIfChanged path, JsonPretty(ooxObj), exported, False
Done:
End Sub

Private Function ValidationTypeString(t As Long) As String
    Select Case t
        Case 0: ValidationTypeString = ""    ' xlValidateInputOnly
        Case 1: ValidationTypeString = "whole"    ' xlValidateWholeNumber
        Case 2: ValidationTypeString = "decimal"    ' xlValidateDecimal
        Case 3: ValidationTypeString = "list"    ' xlValidateList
        Case 4: ValidationTypeString = "date"    ' xlValidateDate
        Case 5: ValidationTypeString = "time"    ' xlValidateTime
        Case 6: ValidationTypeString = "textLength"    ' xlValidateTextLength
        Case 7: ValidationTypeString = "custom"    ' xlValidateCustom
        Case Else: ValidationTypeString = ""
    End Select
End Function

Private Function ValidationOperatorString(o As Long) As String
    Select Case o
        Case 1: ValidationOperatorString = "between"    ' xlBetween
        Case 2: ValidationOperatorString = "notBetween"    ' xlNotBetween
        Case 3: ValidationOperatorString = "equal"    ' xlEqual
        Case 4: ValidationOperatorString = "notEqual"    ' xlNotEqual
        Case 5: ValidationOperatorString = "greaterThan"    ' xlGreater
        Case 6: ValidationOperatorString = "lessThan"    ' xlLess
        Case 7: ValidationOperatorString = "greaterThanOrEqual"    ' xlGreaterEqual
        Case 8: ValidationOperatorString = "lessThanOrEqual"    ' xlLessEqual
        Case Else: ValidationOperatorString = ""
    End Select
End Function

Private Sub SaveConditionalFormatsJson(path As String, ws As Worksheet, exported As Object)
    ' Range.FormatConditions. Walk each FC, capture sqref + type + priority + formula.
    On Error GoTo Done
    Dim list As Object: Set list = New Collection

    Dim fcRange As Range
    On Error Resume Next
    Set fcRange = ws.UsedRange
    On Error GoTo 0
    If fcRange Is Nothing Then Exit Sub

    ' Track seen FCs by appliesTo + type + priority to avoid duplicate emit
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    Dim ur As Range: Set ur = ws.UsedRange
    Dim allFC As Object: Set allFC = CreateObject("Scripting.Dictionary")

    ' Iterate FormatConditions via ws.Cells.FormatConditions (workbook-wide on the sheet)
    Dim fcc As Long
    For fcc = 1 To ws.Cells.FormatConditions.Count
        On Error Resume Next
        Dim fc As Object: Set fc = ws.Cells.FormatConditions(fcc)
        Dim aRef As String: aRef = fc.AppliesTo.Address(False, False)
        Dim fcType As Long: fcType = fc.Type
        Dim fcPriority As Long: fcPriority = fc.Priority

        Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
        entry("sqref") = aRef
        entry("type") = ConditionalFormatTypeString(fcType)
        entry("priority") = fcPriority

        Dim fcOp As Long: fcOp = fc.Operator
        If fcOp <> 0 Then entry("operator") = ValidationOperatorString(fcOp)

        Dim formulae As Object: Set formulae = New Collection
        Dim f1 As String: f1 = ""
        f1 = fc.Formula1
        If Len(f1) > 0 Then formulae.Add f1
        Dim f2 As String: f2 = ""
        f2 = fc.Formula2
        If Len(f2) > 0 Then formulae.Add f2
        If formulae.Count > 0 Then Set entry("formulae") = formulae

        If Err.Number = 0 Then list.Add entry
        Err.Clear
        On Error GoTo 0
    Next

    If list.Count = 0 Then Exit Sub

    ' Sort by sqref then priority (bubble — small lists)
    SortDictListBySqrefPriority list

    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    Set obj("conditionalFormats") = list
    WriteIfChanged path, JsonPretty(obj), exported, False
Done:
End Sub

' Numeric values from XlFormatConditionType. We avoid the enum-name form
' because `xlIconSet` clashes with the IconSet class in some Excel versions
' and produces a compile error "Expected variable or procedure, not module".
Private Function ConditionalFormatTypeString(t As Long) As String
    Select Case t
        Case 1:  ConditionalFormatTypeString = "cellIs"            ' xlCellValue
        Case 2:  ConditionalFormatTypeString = "expression"        ' xlExpression
        Case 3:  ConditionalFormatTypeString = "colorScale"        ' xlColorScale
        Case 4:  ConditionalFormatTypeString = "dataBar"           ' xlDatabar
        Case 5:  ConditionalFormatTypeString = "top10"             ' xlTop10
        Case 6:  ConditionalFormatTypeString = "iconSet"           ' xlIconSet
        Case 8:  ConditionalFormatTypeString = "uniqueValues"      ' xlUniqueValues
        Case 9:  ConditionalFormatTypeString = "containsText"      ' xlTextString
        Case 10: ConditionalFormatTypeString = "containsBlanks"    ' xlBlanksCondition
        Case 11: ConditionalFormatTypeString = "timePeriod"        ' xlTimePeriod
        Case 12: ConditionalFormatTypeString = "aboveAverage"      ' xlAboveAverageCondition
        Case 13: ConditionalFormatTypeString = "notContainsBlanks" ' xlNoBlanksCondition
        Case 16: ConditionalFormatTypeString = "containsErrors"    ' xlErrorsCondition
        Case 17: ConditionalFormatTypeString = "notContainsErrors" ' xlNoErrorsCondition
        Case Else: ConditionalFormatTypeString = "unknown(" & t & ")"
    End Select
End Function

Private Sub SaveCommentsJson(path As String, ws As Worksheet, exported As Object)
    Dim list As Object: Set list = New Collection
    Dim cmt As Comment
    For Each cmt In ws.Comments
        Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
        entry("ref") = cmt.Parent.Address(False, False)
        entry("text") = cmt.Text
        On Error Resume Next
        Dim au As String: au = cmt.Author
        If Len(au) > 0 Then entry("author") = au
        On Error GoTo 0
        list.Add entry
    Next

    If list.Count = 0 Then Exit Sub

    ' Sort by row-major (parse address back to col/row)
    SortCommentsByRowMajor list

    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    Set obj("comments") = list
    WriteIfChanged path, JsonPretty(obj), exported, False
End Sub

Private Sub SortCommentsByRowMajor(ByRef list As Object)
    ' Bubble sort by row*16384 + col (small lists). Each entry has "ref" like "A1".
    ' Collection has no indexed assignment, so we copy out to a Variant array,
    ' sort in place, and rebuild a fresh Collection for the caller.
    Dim n As Long: n = list.Count
    If n <= 1 Then Exit Sub
    Dim arr() As Variant: ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1: Set arr(i) = list(i + 1): Next
    Dim j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            Dim aKey As Long: aKey = AddressSortKey(CStr(arr(j)("ref")))
            Dim bKey As Long: bKey = AddressSortKey(CStr(arr(j + 1)("ref")))
            If aKey > bKey Then
                Dim tmp As Object: Set tmp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = tmp
            End If
        Next
    Next
    Set list = New Collection
    For i = 0 To n - 1: list.Add arr(i): Next
End Sub

Private Function AddressSortKey(addr As String) As Long
    ' Parse "A1" -> 1*16384+1. Best-effort.
    Dim i As Long
    Dim letters As String, digits As String
    For i = 1 To Len(addr)
        Dim ch As String: ch = Mid$(addr, i, 1)
        If ch >= "A" And ch <= "Z" Then
            letters = letters & ch
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If
    Next
    Dim col As Long: col = 0
    For i = 1 To Len(letters)
        col = col * 26 + (Asc(Mid$(letters, i, 1)) - 64)
    Next
    Dim row As Long: row = 0
    If Len(digits) > 0 Then row = CLng(digits)
    AddressSortKey = row * 16384 + col
End Function

Private Sub SaveDrawingsJson(drawingsDir As String, ws As Worksheet, exported As Object, _
                             Optional sheetDrawing As Object)
    ' Walk ws.Shapes. Build entries; create dir + shapes.json only if any.
    ' sheetDrawing (if non-Nothing) is a Dictionary with:
    '   "picsInOrder"  -> ArrayList of media filenames (basename) in drawing.xml
    '                     anchor order (first <pic>, second <pic>, ...).
    '   "picsByName"   -> Dictionary keyed by <xdr:cNvPr name=".."> -> filename.
    '   "mediaPaths"   -> Dictionary keyed by basename -> absolute extracted path.
    Dim shapes As Object: Set shapes = New Collection

    ' For picture ordering: enumerate VBA Shapes in collection order, count
    ' the Nth picture as we go, and match against picsInOrder when name lookup
    ' fails. VBA's ws.Shapes collection order generally matches OOXML anchor
    ' order, but cNvPr-name lookup is the authoritative path.
    Dim picIdx As Long: picIdx = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim entry As Object: Set entry = ShapeToEntry(shp)
        If Not entry Is Nothing Then
            ' If this is a picture entry and we have a drawing map, rewrite
            ' the placeholder asset filename with the real OOXML media name.
            If entry.Exists("type") Then
                If CStr(entry("type")) = "picture" Then
                    Dim realName As String: realName = ""
                    If Not sheetDrawing Is Nothing Then
                        If sheetDrawing.Exists("picsByName") Then
                            If sheetDrawing("picsByName").Exists(shp.Name) Then
                                realName = CStr(sheetDrawing("picsByName")(shp.Name))
                            End If
                        End If
                        If Len(realName) = 0 Then
                            ' Fall back to positional lookup
                            If sheetDrawing.Exists("picsInOrder") Then
                                If picIdx < sheetDrawing("picsInOrder").Count Then
                                    realName = CStr(sheetDrawing("picsInOrder")(picIdx))
                                End If
                            End If
                        End If
                    End If
                    If Len(realName) > 0 Then
                        entry("asset") = "_assets/" & realName
                    End If
                    picIdx = picIdx + 1
                End If
            End If
            shapes.Add entry
        End If
    Next

    If shapes.Count = 0 Then Exit Sub

    EnsureFolder drawingsDir
    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    Set obj("shapes") = shapes
    WriteIfChanged drawingsDir & "shapes.json", JsonPretty(obj), exported, False

    ' Copy media binaries into _assets/. Copy ALL media files referenced by
    ' this sheet's drawing rels, even if we couldn't match every shape -- this
    ' keeps the asset directory complete and lets shapes.json be reviewed
    ' against the actual binaries.
    If Not sheetDrawing Is Nothing Then
        If sheetDrawing.Exists("mediaPaths") Then
            Dim mp As Object: Set mp = sheetDrawing("mediaPaths")
            If mp.Count > 0 Then
                Dim assetsDir As String: assetsDir = drawingsDir & "_assets\"
                EnsureFolder assetsDir
                Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
                Dim mk As Variant
                For Each mk In mp.Keys
                    Dim src As String: src = CStr(mp(mk))
                    Dim dst As String: dst = assetsDir & CStr(mk)
                    On Error Resume Next
                    If fso.FileExists(src) Then
                        Dim doCopy As Boolean: doCopy = True
                        If fso.FileExists(dst) Then
                            If fso.GetFile(dst).Size = fso.GetFile(src).Size Then
                                doCopy = False
                            End If
                        End If
                        If doCopy Then fso.CopyFile src, dst, True
                        If Not exported Is Nothing Then exported(dst) = True
                    End If
                    On Error GoTo 0
                Next
            End If
        End If
    End If
End Sub

Private Function ShapeToEntry(shp As Shape) As Object
    On Error GoTo Skip
    Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")

    ' Anchor
    Dim tlc As Range, brc As Range
    On Error Resume Next
    Set tlc = shp.TopLeftCell
    Set brc = shp.BottomRightCell
    On Error GoTo Skip
    If Not tlc Is Nothing And Not brc Is Nothing Then
        Dim anchor As Object: Set anchor = CreateObject("Scripting.Dictionary")
        anchor("from") = tlc.Address(False, False)
        anchor("to") = brc.Address(False, False)
        Set entry("anchor") = anchor
    End If

    Dim st As Long: st = shp.Type

    ' Numeric MsoShapeType values — `msoConnector` collides with the
    ' Connector class on some Excel installs (same family as xlIconSet).
    ' 13=Picture, 11=LinkedPicture, 6=Group, 9=Line, 18=Connector
    ' (deprecated, treated like Line), 5=Freeform.
    Select Case st
        Case 13, 11   ' msoPicture, msoLinkedPicture
            entry("type") = "picture"
            entry("name") = shp.Name
            ' Asset path; binary is copied to _assets/ by SaveDrawingsJson.
            entry("asset") = "_assets/" & SafePictureBaseName(shp.Name) & ".png"
        Case 6        ' msoGroup
            entry("type") = "group"
            entry("name") = shp.Name
            ' Recurse children (one level)
            Dim children As Object: Set children = New Collection
            On Error Resume Next
            Dim child As Shape
            For Each child In shp.GroupItems
                Dim cEntry As Object: Set cEntry = ShapeToEntry(child)
                If Not cEntry Is Nothing Then children.Add cEntry
            Next
            On Error GoTo Skip
            If children.Count > 0 Then Set entry("children") = children
        Case 9, 18, 5 ' msoLine, msoConnector, msoFreeform
            entry("type") = "connector"
            entry("name") = shp.Name
        Case Else
            entry("type") = "shape"
            entry("name") = shp.Name
            ' OnAction (assigned macro)
            On Error Resume Next
            Dim oa As String: oa = shp.OnAction
            On Error GoTo Skip
            If Len(oa) > 0 Then
                ' Strip "[N]!" workbook-id prefix
                Dim re As Object: Set re = CreateObject("VBScript.RegExp")
                re.Pattern = "^\[\d+\]!"
                If re.Test(oa) Then oa = re.Replace(oa, "")
                entry("macro") = oa
            End If
            ' Preset geometry
            On Error Resume Next
            Dim preset As String: preset = AutoShapePresetName(shp)
            On Error GoTo Skip
            If Len(preset) > 0 Then entry("preset") = preset
            ' Text content
            On Error Resume Next
            Dim text As String
            If shp.HasTextFrame Then
                If shp.TextFrame2.HasText Then
                    text = shp.TextFrame2.TextRange.Text
                End If
            End If
            On Error GoTo Skip
            If Len(text) > 0 Then entry("text") = text
    End Select

    Set ShapeToEntry = entry
    Exit Function
Skip:
    Set ShapeToEntry = Nothing
End Function

Private Function AutoShapePresetName(shp As Shape) As String
    On Error Resume Next
    Dim t As Long: t = shp.AutoShapeType
    On Error GoTo 0
    Select Case t
        Case 1: AutoShapePresetName = "rect"    ' msoShapeRectangle
        Case 5: AutoShapePresetName = "roundRect"    ' msoShapeRoundedRectangle
        Case 9: AutoShapePresetName = "ellipse"    ' msoShapeOval
        Case 8: AutoShapePresetName = "rtTriangle"    ' msoShapeRightTriangle
        Case 7: AutoShapePresetName = "triangle"    ' msoShapeIsoscelesTriangle
        Case 4: AutoShapePresetName = "diamond"    ' msoShapeDiamond
        Case 10: AutoShapePresetName = "hexagon"    ' msoShapeHexagon
        Case 11: AutoShapePresetName = "octagon"    ' msoShapeOctagon
        Case 33: AutoShapePresetName = "rightArrow"    ' msoShapeRightArrow
        Case 34: AutoShapePresetName = "leftArrow"    ' msoShapeLeftArrow
        Case 35: AutoShapePresetName = "upArrow"    ' msoShapeUpArrow
        Case 36: AutoShapePresetName = "downArrow"    ' msoShapeDownArrow
        Case Else: AutoShapePresetName = ""
    End Select
End Function

Private Function SafePictureBaseName(shpName As String) As String
    Dim s As String: s = shpName
    Dim bad As Variant: bad = Array("/", "\", ":", "*", "?", """", "<", ">", "|", " ")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next
    SafePictureBaseName = s
End Function

'====================  CHART EXPORT  =========================
' For each embedded ChartObject on the sheet, write <ChartName>.png (via
' chartObj.Chart.Export) and <ChartName>.json (structural manifest).
'
' JSON schema (PNG is the visual source of truth):
'   { name, type, title, anchor: {from, to},
'     series: [{name, formula}], axes: {categoryAxis: {title}, valueAxis: {title}} }
Private Sub ExportCharts(chartsDir As String, ws As Worksheet, exported As Object)
    On Error GoTo Done
    If ws.ChartObjects.Count = 0 Then Exit Sub

    EnsureFolder chartsDir

    Dim chartIdx As Long: chartIdx = 0
    Dim totalCharts As Long: totalCharts = ws.ChartObjects.Count
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        chartIdx = chartIdx + 1
        On Error Resume Next
        Application.StatusBar = "VBA Sync: exporting chart " & chartIdx & " of " & totalCharts & " on '" & ws.Name & "'..."
        On Error GoTo 0

        Dim safeName As String: safeName = SafeFileName(co.Name)
        Dim pngPath As String: pngPath = chartsDir & safeName & ".png"
        Dim jsonPath As String: jsonPath = chartsDir & safeName & ".json"

        ' --- PNG snapshot ---
        On Error Resume Next
        ' Delete any existing PNG first; Chart.Export sometimes refuses to overwrite.
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(pngPath) Then fso.DeleteFile pngPath, True
        co.Chart.Export Filename:=pngPath, FilterName:="PNG"
        If Err.Number <> 0 Then
            Debug.Print "VBA Sync: chart PNG export failed for '" & co.Name & "' on '" & ws.Name & "': " & Err.Description
            Err.Clear
        Else
            If Not exported Is Nothing Then exported(pngPath) = True
        End If
        On Error GoTo 0

        ' --- JSON manifest ---
        Dim obj As Object: Set obj = BuildChartJson(co)
        If Not obj Is Nothing Then
            WriteIfChanged jsonPath, JsonPretty(obj), exported, False
        End If
    Next
Done:
    If Err.Number <> 0 Then
        Debug.Print "VBA Sync: ExportCharts on '" & ws.Name & "' failed: " & Err.Description
        Err.Clear
    End If
End Sub

Private Function BuildChartJson(co As ChartObject) As Object
    On Error GoTo Skip
    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    obj("name") = co.Name

    ' --- Type (XlChartType enum) ---
    Dim ct As Variant: ct = Empty
    On Error Resume Next
    ct = co.Chart.ChartType
    On Error GoTo 0
    If Not IsEmpty(ct) Then obj("type") = ChartTypeString(CLng(ct))

    ' --- Title ---
    Dim title As String: title = ""
    On Error Resume Next
    If co.Chart.HasTitle Then
        title = co.Chart.ChartTitle.Text
    End If
    On Error GoTo 0
    If Len(title) > 0 Then obj("title") = title

    ' --- Anchor (TopLeftCell / BottomRightCell of the host shape) ---
    On Error Resume Next
    Dim tlc As Range: Set tlc = co.TopLeftCell
    Dim brc As Range: Set brc = co.BottomRightCell
    On Error GoTo 0
    If Not tlc Is Nothing And Not brc Is Nothing Then
        Dim anchor As Object: Set anchor = CreateObject("Scripting.Dictionary")
        anchor("from") = tlc.Address(False, False)
        anchor("to") = brc.Address(False, False)
        Set obj("anchor") = anchor
    End If

    ' --- Series (Formula stored as-is; raw SERIES(...) string) ---
    Dim seriesList As Object: Set seriesList = New Collection
    On Error Resume Next
    Dim sc As Object: Set sc = co.Chart.SeriesCollection
    On Error GoTo 0
    If Not sc Is Nothing Then
        Dim si As Long
        Dim seriesCount As Long: seriesCount = 0
        On Error Resume Next
        seriesCount = sc.Count
        On Error GoTo 0
        For si = 1 To seriesCount
            On Error Resume Next
            Dim s As Object: Set s = sc.Item(si)
            If Not s Is Nothing Then
                Dim sEntry As Object: Set sEntry = CreateObject("Scripting.Dictionary")
                Dim sName As String: sName = ""
                sName = s.Name
                If Len(sName) > 0 Then sEntry("name") = sName
                Dim sFormula As String: sFormula = ""
                sFormula = s.Formula
                If Len(sFormula) > 0 Then sEntry("formula") = sFormula
                If sEntry.Count > 0 Then seriesList.Add sEntry
            End If
            On Error GoTo 0
        Next
    End If
    If seriesList.Count > 0 Then Set obj("series") = seriesList

    ' --- Axes (category + value titles only) ---
    Dim axes As Object: Set axes = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim catAx As Object: Set catAx = co.Chart.Axes(xlCategory)
    If Not catAx Is Nothing Then
        Dim catTitle As String: catTitle = ""
        If catAx.HasTitle Then catTitle = catAx.AxisTitle.Text
        If Len(catTitle) > 0 Then
            Dim catObj As Object: Set catObj = CreateObject("Scripting.Dictionary")
            catObj("title") = catTitle
            Set axes("categoryAxis") = catObj
        End If
    End If
    Dim valAx As Object: Set valAx = co.Chart.Axes(xlValue)
    If Not valAx Is Nothing Then
        Dim valTitle As String: valTitle = ""
        If valAx.HasTitle Then valTitle = valAx.AxisTitle.Text
        If Len(valTitle) > 0 Then
            Dim valObj As Object: Set valObj = CreateObject("Scripting.Dictionary")
            valObj("title") = valTitle
            Set axes("valueAxis") = valObj
        End If
    End If
    On Error GoTo 0
    If axes.Count > 0 Then Set obj("axes") = axes

    Set BuildChartJson = obj
    Exit Function
Skip:
    Set BuildChartJson = Nothing
End Function

' Map XlChartType enum -> readable name. Falls back to integer for unknown.
' Common values; not exhaustive. Don't crash on unmapped values.
Private Function ChartTypeString(t As Long) As String
    Select Case t
        Case 51: ChartTypeString = "xlColumnClustered"    ' xlColumnClustered
        Case 52: ChartTypeString = "xlColumnStacked"    ' xlColumnStacked
        Case 53: ChartTypeString = "xlColumnStacked100"    ' xlColumnStacked100
        Case -4100: ChartTypeString = "xl3DColumn"    ' xl3DColumn
        Case 54: ChartTypeString = "xl3DColumnClustered"    ' xl3DColumnClustered
        Case 55: ChartTypeString = "xl3DColumnStacked"    ' xl3DColumnStacked
        Case 56: ChartTypeString = "xl3DColumnStacked100"    ' xl3DColumnStacked100
        Case 57: ChartTypeString = "xlBarClustered"    ' xlBarClustered
        Case 58: ChartTypeString = "xlBarStacked"    ' xlBarStacked
        Case 59: ChartTypeString = "xlBarStacked100"    ' xlBarStacked100
        Case 60: ChartTypeString = "xl3DBarClustered"    ' xl3DBarClustered
        Case 61: ChartTypeString = "xl3DBarStacked"    ' xl3DBarStacked
        Case 62: ChartTypeString = "xl3DBarStacked100"    ' xl3DBarStacked100
        Case 4: ChartTypeString = "xlLine"    ' xlLine
        Case 65: ChartTypeString = "xlLineMarkers"    ' xlLineMarkers
        Case 66: ChartTypeString = "xlLineMarkersStacked"    ' xlLineMarkersStacked
        Case 67: ChartTypeString = "xlLineMarkersStacked100"    ' xlLineMarkersStacked100
        Case 63: ChartTypeString = "xlLineStacked"    ' xlLineStacked
        Case 64: ChartTypeString = "xlLineStacked100"    ' xlLineStacked100
        Case -4101: ChartTypeString = "xl3DLine"    ' xl3DLine
        Case 5: ChartTypeString = "xlPie"    ' xlPie
        Case -4102: ChartTypeString = "xl3DPie"    ' xl3DPie
        Case 70: ChartTypeString = "xl3DPieExploded"    ' xl3DPieExploded
        Case 69: ChartTypeString = "xlPieExploded"    ' xlPieExploded
        Case 68: ChartTypeString = "xlPieOfPie"    ' xlPieOfPie
        Case 71: ChartTypeString = "xlBarOfPie"    ' xlBarOfPie
        Case -4120: ChartTypeString = "xlDoughnut"    ' xlDoughnut
        Case 80: ChartTypeString = "xlDoughnutExploded"    ' xlDoughnutExploded
        Case -4169: ChartTypeString = "xlXYScatter"    ' xlXYScatter
        Case 74: ChartTypeString = "xlXYScatterLines"    ' xlXYScatterLines
        Case 75: ChartTypeString = "xlXYScatterLinesNoMarkers"    ' xlXYScatterLinesNoMarkers
        Case 72: ChartTypeString = "xlXYScatterSmooth"    ' xlXYScatterSmooth
        Case 73: ChartTypeString = "xlXYScatterSmoothNoMarkers"    ' xlXYScatterSmoothNoMarkers
        Case 15: ChartTypeString = "xlBubble"    ' xlBubble
        Case 87: ChartTypeString = "xlBubble3DEffect"    ' xlBubble3DEffect
        Case 1: ChartTypeString = "xlArea"    ' xlArea
        Case 76: ChartTypeString = "xlAreaStacked"    ' xlAreaStacked
        Case 77: ChartTypeString = "xlAreaStacked100"    ' xlAreaStacked100
        Case -4098: ChartTypeString = "xl3DArea"    ' xl3DArea
        Case 78: ChartTypeString = "xl3DAreaStacked"    ' xl3DAreaStacked
        Case 79: ChartTypeString = "xl3DAreaStacked100"    ' xl3DAreaStacked100
        Case -4151: ChartTypeString = "xlRadar"    ' xlRadar
        Case 82: ChartTypeString = "xlRadarFilled"    ' xlRadarFilled
        Case 81: ChartTypeString = "xlRadarMarkers"    ' xlRadarMarkers
        Case 88: ChartTypeString = "xlStockHLC"    ' xlStockHLC
        Case 89: ChartTypeString = "xlStockOHLC"    ' xlStockOHLC
        Case 90: ChartTypeString = "xlStockVHLC"    ' xlStockVHLC
        Case 91: ChartTypeString = "xlStockVOHLC"    ' xlStockVOHLC
        Case 83: ChartTypeString = "xlSurface"    ' xlSurface
        Case 85: ChartTypeString = "xlSurfaceTopView"    ' xlSurfaceTopView
        Case 86: ChartTypeString = "xlSurfaceTopViewWireframe"    ' xlSurfaceTopViewWireframe
        Case 84: ChartTypeString = "xlSurfaceWireframe"    ' xlSurfaceWireframe
        Case Else: ChartTypeString = CStr(t)
    End Select
End Function

'====================  TABLE CALC-COLUMN FORMULA LOOKUP  =====
' Reads xl/tables/*.xml from the workbook's already-unzipped temp folder
' (created by BuildDrawingImageMap for image binaries) and returns a map:
'   tableNameLowercase -> Dictionary("columnName" -> "=formula")
'
' Why this exists: VBA's ListColumn object has NO calc-formula property.
' For tables with at least one data row we can read the formula off
' lc.DataBodyRange.Cells(...).Formula, but for tables with zero data rows
' (DataBodyRange Is Nothing) there's no cell to read from. The formula
' lives ONLY in the OOXML <calculatedColumnFormula> element on the column
' definition. Since we already have the workbook unzipped for image
' extraction, reading these tiny XML files is essentially free.
'
' Returns an empty Dictionary if tempDir is missing or has no tables.
Private Function LoadTableCalcFormulas(tempDir As String) As Object
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    On Error GoTo Done
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tablesDir As String: tablesDir = tempDir & "\xl\tables"
    If Not fso.FolderExists(tablesDir) Then GoTo Done

    Dim folder As Object: Set folder = fso.GetFolder(tablesDir)
    Dim f As Object
    For Each f In folder.Files
        If LCase$(fso.GetExtensionName(f.Path)) = "xml" Then
            On Error Resume Next
            Dim doc As Object: Set doc = CreateObject("MSXML2.DOMDocument.6.0")
            doc.async = False
            doc.SetProperty "SelectionLanguage", "XPath"
            doc.Load f.Path
            On Error GoTo Done
            If doc Is Nothing Then GoTo NextFile
            If doc.parseError.errorCode <> 0 Then GoTo NextFile

            ' xl/tables/tableN.xml has <table name="..." ...> at root with
            ' <tableColumns><tableColumn name="..."><calculatedColumnFormula>
            Dim ns As String
            ns = "xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main'"
            doc.SetProperty "SelectionNamespaces", ns
            Dim tableNode As Object: Set tableNode = doc.SelectSingleNode("/s:table")
            If tableNode Is Nothing Then GoTo NextFile
            Dim tableName As String
            tableName = ""
            If Not tableNode.Attributes.getNamedItem("name") Is Nothing Then
                tableName = tableNode.Attributes.getNamedItem("name").Text
            End If
            If Len(tableName) = 0 Then GoTo NextFile

            Dim colFormulas As Object: Set colFormulas = CreateObject("Scripting.Dictionary")
            Dim colNodes As Object: Set colNodes = doc.SelectNodes("/s:table/s:tableColumns/s:tableColumn")
            Dim i As Long
            For i = 0 To colNodes.Length - 1
                Dim colNode As Object: Set colNode = colNodes.Item(i)
                Dim colName As String: colName = ""
                If Not colNode.Attributes.getNamedItem("name") Is Nothing Then
                    colName = colNode.Attributes.getNamedItem("name").Text
                End If
                If Len(colName) > 0 Then
                    Dim ccfNode As Object: Set ccfNode = colNode.SelectSingleNode("s:calculatedColumnFormula")
                    If Not ccfNode Is Nothing Then
                        Dim ccf As String: ccf = ccfNode.Text
                        If Len(ccf) > 0 Then
                            ' OOXML stores formulas WITHOUT leading "=". Add it
                            ' so downstream code can treat it like a cell formula.
                            If Left$(ccf, 1) <> "=" Then ccf = "=" & ccf
                            colFormulas(colName) = ccf
                        End If
                    End If
                End If
            Next i

            If colFormulas.Count > 0 Then
                Set out(LCase$(tableName)) = colFormulas
            End If
NextFile:
        End If
    Next
Done:
    Set LoadTableCalcFormulas = out
End Function

' Parse xl/workbook.xml + xl/_rels/workbook.xml.rels to get a {sheetName ->
' "worksheets/sheetN.xml"} mapping. Standalone helper so the OOXML readers
' (LoadSheetValidations, LoadSheetViews) don't depend on BuildDrawingImageMap
' completing -- BuildDrawingImageMap's Shell.Application extraction wait is
' fragile under COM and the function often returns early even though the
' tempDir's xl/ tree finishes extracting in the background.
Private Function LoadWorkbookSheetMap(tempDir As String) As Object
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    On Error GoTo Done
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim wbXmlPath As String: wbXmlPath = tempDir & "\xl\workbook.xml"
    Dim wbRelsPath As String: wbRelsPath = tempDir & "\xl\_rels\workbook.xml.rels"
    ' Extraction is synchronous (ExtractZipSynchronously) so files are on disk
    ' by the time we get here; no polling needed.
    If Not fso.FileExists(wbXmlPath) Or Not fso.FileExists(wbRelsPath) Then GoTo Done

    Dim wbXml As Object: Set wbXml = LoadXmlFile(wbXmlPath)
    Dim wbRels As Object: Set wbRels = LoadXmlFile(wbRelsPath)
    If wbXml Is Nothing Or wbRels Is Nothing Then GoTo Done

    Dim wbRelMap As Object: Set wbRelMap = CreateObject("Scripting.Dictionary")
    Dim relNode As Object
    For Each relNode In wbRels.documentElement.childNodes
        Dim relId As String, relTgt As String
        On Error Resume Next
        relId = relNode.getAttribute("Id")
        relTgt = relNode.getAttribute("Target")
        On Error GoTo Done
        If Len(relId) > 0 And Len(relTgt) > 0 Then wbRelMap(relId) = relTgt
    Next

    Dim sheetsNode As Object, sheetNode As Object
    For Each sheetsNode In wbXml.documentElement.childNodes
        If LCase$(localName(sheetsNode)) = "sheets" Then
            For Each sheetNode In sheetsNode.childNodes
                Dim sName As String, sRid As String
                On Error Resume Next
                sName = sheetNode.getAttribute("name")
                sRid = sheetNode.getAttribute("r:id")
                If Len(sRid) = 0 Then
                    Dim att As Object
                    For Each att In sheetNode.Attributes
                        If LCase$(att.baseName) = "id" Then sRid = att.Value: Exit For
                    Next
                End If
                On Error GoTo Done
                If Len(sName) > 0 And Len(sRid) > 0 Then
                    If wbRelMap.Exists(sRid) Then out(sName) = CStr(wbRelMap(sRid))
                End If
            Next
            Exit For
        End If
    Next
Done:
    Set LoadWorkbookSheetMap = out
End Function

'====================  OOXML SHEET DATA-VALIDATIONS  =========
' Reads xl/worksheets/sheetN.xml for each sheet, parses <dataValidations>,
' and returns Dictionary keyed by Worksheet.Name -> Collection of
' Scripting.Dictionary entries with the same shape SaveValidationsJson
' emits (sqref, type, operator, allowBlank, showInputMessage,
' showErrorMessage, errorTitle, error, promptTitle, prompt, formula1, formula2).
'
' OOXML rather than Range.SpecialCells(xlCellTypeAllValidation): the
' VBA call returns 1004 "no cells found" in some COM contexts even when
' validations exist. OOXML is authoritative.
Private Function LoadSheetValidations(tempDir As String, sheetNameToTarget As Object) As Object
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    On Error GoTo Done
    If sheetNameToTarget Is Nothing Then GoTo Done

    Dim ns As String
    ns = "xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main'"

    Dim shName As Variant
    For Each shName In sheetNameToTarget.Keys
        Dim relPath As String: relPath = Replace(CStr(sheetNameToTarget(shName)), "/", "\")
        Dim sheetPath As String: sheetPath = tempDir & "\xl\" & relPath
        On Error Resume Next
        Dim doc As Object: Set doc = CreateObject("MSXML2.DOMDocument.6.0")
        doc.async = False
        doc.SetProperty "SelectionLanguage", "XPath"
        doc.SetProperty "SelectionNamespaces", ns
        doc.Load sheetPath
        On Error GoTo Done
        If doc Is Nothing Then GoTo NextSheet
        If doc.parseError.errorCode <> 0 Then GoTo NextSheet

        Dim dvNodes As Object: Set dvNodes = doc.SelectNodes("/s:worksheet/s:dataValidations/s:dataValidation")
        If dvNodes Is Nothing Then GoTo NextSheet
        If dvNodes.Length = 0 Then GoTo NextSheet

        Dim entries As Object: Set entries = New Collection
        Dim i As Long
        For i = 0 To dvNodes.Length - 1
            Dim dvNode As Object: Set dvNode = dvNodes.Item(i)
            Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")

            Dim sqref As String: sqref = GetAttr(dvNode, "sqref")
            ' Translate OOXML's space-separated absolute refs to the same
            ' space-separated form SaveValidationsJson emits.
            entry("sqref") = sqref

            Dim t As String: t = GetAttr(dvNode, "type")
            If Len(t) > 0 Then entry("type") = t
            Dim op As String: op = GetAttr(dvNode, "operator")
            If Len(op) > 0 Then entry("operator") = op
            If GetAttr(dvNode, "allowBlank") = "1" Then entry("allowBlank") = True
            If GetAttr(dvNode, "showInputMessage") = "1" Then entry("showInputMessage") = True
            If GetAttr(dvNode, "showErrorMessage") = "1" Then entry("showErrorMessage") = True
            Dim et As String: et = GetAttr(dvNode, "errorTitle"): If Len(et) > 0 Then entry("errorTitle") = et
            Dim em As String: em = GetAttr(dvNode, "error"): If Len(em) > 0 Then entry("error") = em
            Dim pt As String: pt = GetAttr(dvNode, "promptTitle"): If Len(pt) > 0 Then entry("promptTitle") = pt
            Dim pm As String: pm = GetAttr(dvNode, "prompt"): If Len(pm) > 0 Then entry("prompt") = pm

            Dim f1Node As Object: Set f1Node = dvNode.SelectSingleNode("s:formula1")
            If Not f1Node Is Nothing Then
                Dim f1 As String: f1 = f1Node.Text
                If Len(f1) > 0 Then entry("formula1") = f1
            End If
            Dim f2Node As Object: Set f2Node = dvNode.SelectSingleNode("s:formula2")
            If Not f2Node Is Nothing Then
                Dim f2 As String: f2 = f2Node.Text
                If Len(f2) > 0 Then entry("formula2") = f2
            End If

            entries.Add entry
        Next
        If entries.Count > 0 Then Set out(CStr(shName)) = entries
NextSheet:
        Set doc = Nothing
    Next
Done:
    Set LoadSheetValidations = out
End Function

'====================  OOXML SHEET VIEW SETTINGS  ============
' Reads xl/worksheets/sheetN.xml, parses the first <sheetView> + any
' <pane>, returns Dictionary keyed by Worksheet.Name ->
'   { "showGridLines": Boolean (omitted unless explicitly false),
'     "frozenPanes": { "xSplit": N, "ySplit": N } (if any) }
'
' OOXML rather than Window.DisplayGridlines / Window.FreezePanes because
' those Window properties only reflect the active sheet.
Private Function LoadSheetViews(tempDir As String, sheetNameToTarget As Object) As Object
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    On Error GoTo Done
    If sheetNameToTarget Is Nothing Then GoTo Done

    Dim ns As String
    ns = "xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main'"

    Dim shName As Variant
    For Each shName In sheetNameToTarget.Keys
        Dim relPath As String: relPath = Replace(CStr(sheetNameToTarget(shName)), "/", "\")
        Dim sheetPath As String: sheetPath = tempDir & "\xl\" & relPath
        On Error Resume Next
        Dim doc As Object: Set doc = CreateObject("MSXML2.DOMDocument.6.0")
        doc.async = False
        doc.SetProperty "SelectionLanguage", "XPath"
        doc.SetProperty "SelectionNamespaces", ns
        doc.Load sheetPath
        On Error GoTo Done
        If doc Is Nothing Then GoTo NextSheet
        If doc.parseError.errorCode <> 0 Then GoTo NextSheet

        ' Per-sheet error tolerance: any failure inside the per-sheet block
        ' should skip that sheet, not abort the whole loader. (Important
        ' because VBA's And operator doesn't short-circuit, so naive
        ' Len(s) > 0 And CDbl(s) > 0 raises type-mismatch on empty strings.)
        On Error Resume Next
        Dim svNode As Object: Set svNode = doc.SelectSingleNode("/s:worksheet/s:sheetViews/s:sheetView")
        If svNode Is Nothing Then GoTo NextSheet

        Dim view As Object: Set view = CreateObject("Scripting.Dictionary")

        ' showGridLines defaults to true; only present when explicitly "0"
        If GetAttr(svNode, "showGridLines") = "0" Then view("showGridLines") = False

        Dim paneNode As Object: Set paneNode = svNode.SelectSingleNode("s:pane")
        If Not paneNode Is Nothing Then
            ' OOXML emits state="frozen" for true freezes (state="split" is
            ' the draggable-split case, which we don't surface).
            If GetAttr(paneNode, "state") = "frozen" Then
                Dim fz As Object: Set fz = CreateObject("Scripting.Dictionary")
                Dim xs As String: xs = GetAttr(paneNode, "xSplit")
                Dim ys As String: ys = GetAttr(paneNode, "ySplit")
                If Len(xs) > 0 Then If CDbl(xs) > 0 Then fz("xSplit") = CLng(CDbl(xs))
                If Len(ys) > 0 Then If CDbl(ys) > 0 Then fz("ySplit") = CLng(CDbl(ys))
                If fz.Count > 0 Then Set view("frozenPanes") = fz
            End If
        End If

        If view.Count > 0 Then Set out(CStr(shName)) = view
NextSheet:
        Set doc = Nothing
    Next
    On Error GoTo Done
Done:
    Set LoadSheetViews = out
End Function

Private Function GetAttr(node As Object, name As String) As String
    On Error Resume Next
    Dim a As Object: Set a = node.Attributes.getNamedItem(name)
    If Not a Is Nothing Then GetAttr = a.Text Else GetAttr = ""
    On Error GoTo 0
End Function

'====================  DRAWING IMAGE EXTRACTION  =============
' Unzip the workbook once via Shell.Application; parse rels chains to build
' a per-sheet map of (cNvPr name -> media filename), (positional pic list),
' and (basename -> absolute extracted media path).
'
' Returns Dictionary with:
'   "tempDir"   -> temp folder path (caller deletes when done)
'   "perSheet"  -> Dictionary keyed by Worksheet.Name (matches wb.Worksheets()
'                  via document-order index) -> sheet drawing info dict
'
' Why ws.Name as the key: VBA's wb.Worksheets enumerates in tab order, and
' xl/workbook.xml's <sheets><sheet> elements appear in tab order too -- we
' iterate both together and key by the Excel-side ws.Name (which equals the
' OOXML <sheet name="..">).
Private Function BuildDrawingImageMap(wb As Workbook) As Object
    On Error GoTo Fail
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim perSheet As Object: Set perSheet = CreateObject("Scripting.Dictionary")
    Set result("perSheet") = perSheet

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- Stage workbook into a temp .zip ---
    Dim tempBase As String: tempBase = fso.GetSpecialFolder(2) ' TemporaryFolder
    Dim stamp As String: stamp = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd * 100000))
    Dim tempDir As String: tempDir = tempBase & "\vbasync_xl_" & stamp
    Dim zipPath As String: zipPath = tempDir & ".zip"
    fso.CreateFolder tempDir
    result("tempDir") = tempDir

    ' Copy the workbook to .zip extension. Works for .xlsm and .xlsb -- both
    ' are ZIP packages. The workbook can be open in Excel; FSO can still copy
    ' read-only.
    On Error Resume Next
    fso.CopyFile wb.FullName, zipPath, True
    If Err.Number <> 0 Then
        Debug.Print "VBA Sync: BuildDrawingImageMap copy failed: " & Err.Description
        Err.Clear
        ' Fall back: ask Excel to save a copy. (Edge case; usually CopyFile works.)
        On Error GoTo Fail
        wb.SaveCopyAs zipPath
    End If
    On Error GoTo Fail

    ExtractZipSynchronously zipPath, tempDir

    Dim wbXmlPath As String: wbXmlPath = tempDir & "\xl\workbook.xml"
    Dim wbRelsPath As String: wbRelsPath = tempDir & "\xl\_rels\workbook.xml.rels"
    If Not fso.FileExists(wbXmlPath) Or Not fso.FileExists(wbRelsPath) Then GoTo Fail

    Dim wbXml As Object: Set wbXml = LoadXmlFile(wbXmlPath)
    Dim wbRels As Object: Set wbRels = LoadXmlFile(wbRelsPath)
    If wbXml Is Nothing Or wbRels Is Nothing Then GoTo Fail

    ' rId -> target (e.g. "worksheets/sheet1.xml")
    Dim wbRelMap As Object: Set wbRelMap = CreateObject("Scripting.Dictionary")
    Dim relNode As Object
    For Each relNode In wbRels.documentElement.childNodes
        Dim relId As String: relId = ""
        Dim relTgt As String: relTgt = ""
        On Error Resume Next
        relId = relNode.getAttribute("Id")
        relTgt = relNode.getAttribute("Target")
        On Error GoTo Fail
        If Len(relId) > 0 And Len(relTgt) > 0 Then wbRelMap(relId) = relTgt
    Next

    ' Walk <sheets><sheet name=".." r:id="rIdN"/></sheets> in DOCUMENT ORDER.
    ' This order matches wb.Worksheets enumeration (tab order).
    Dim sheetsNode As Object
    Dim sheetNode As Object
    Dim sheetNameToTarget As Object: Set sheetNameToTarget = CreateObject("Scripting.Dictionary")
    For Each sheetsNode In wbXml.documentElement.childNodes
        If LCase$(localName(sheetsNode)) = "sheets" Then
            For Each sheetNode In sheetsNode.childNodes
                Dim sName As String: sName = ""
                Dim sRid As String: sRid = ""
                On Error Resume Next
                sName = sheetNode.getAttribute("name")
                ' r:id is namespaced; getAttribute("r:id") works against MSXML2.DOMDocument
                ' when we don't enable namespace-aware mode. Fall back to scanning attrs.
                sRid = sheetNode.getAttribute("r:id")
                If Len(sRid) = 0 Then
                    Dim att As Object
                    For Each att In sheetNode.Attributes
                        If LCase$(att.baseName) = "id" Then
                            sRid = att.Value
                            Exit For
                        End If
                    Next
                End If
                On Error GoTo Fail
                If Len(sName) > 0 And Len(sRid) > 0 Then
                    If wbRelMap.Exists(sRid) Then
                        sheetNameToTarget(sName) = CStr(wbRelMap(sRid))
                    End If
                End If
            Next
            Exit For
        End If
    Next
    ' Expose the sheet-name -> OOXML-file map on the result so the OOXML
    ' validations / sheetView readers can reuse it without re-parsing
    ' workbook.xml. May be empty if BuildDrawingImageMap fails late in
    ' the per-sheet rels loop -- callers must check.
    Set result("sheetNameToTarget") = sheetNameToTarget

    ' --- For each sheet, find its drawing and copy its image rels ---
    Dim mediaSrcDir As String: mediaSrcDir = tempDir & "\xl\media"
    Dim shName As Variant
    For Each shName In sheetNameToTarget.Keys
        Dim sheetTarget As String: sheetTarget = CStr(sheetNameToTarget(shName))  ' e.g. "worksheets/sheet1.xml"
        Dim sheetFile As String: sheetFile = fso.GetFileName(sheetTarget)         ' "sheet1.xml"
        Dim sheetRelsPath As String
        sheetRelsPath = tempDir & "\xl\worksheets\_rels\" & sheetFile & ".rels"
        If Not fso.FileExists(sheetRelsPath) Then GoTo NextSheet

        Dim sheetRels As Object: Set sheetRels = LoadXmlFile(sheetRelsPath)
        If sheetRels Is Nothing Then GoTo NextSheet

        ' Find the drawing rel ("Type" ending in "/drawing")
        Dim drawingTarget As String: drawingTarget = ""
        Dim sr As Object
        For Each sr In sheetRels.documentElement.childNodes
            Dim sType As String: sType = ""
            Dim sTarget As String: sTarget = ""
            On Error Resume Next
            sType = sr.getAttribute("Type")
            sTarget = sr.getAttribute("Target")
            On Error GoTo Fail
            If Len(sType) > 0 And InStr(LCase$(sType), "/drawing") > 0 And InStr(LCase$(sType), "vmldrawing") = 0 Then
                drawingTarget = sTarget
                Exit For
            End If
        Next
        If Len(drawingTarget) = 0 Then GoTo NextSheet

        ' Resolve drawing path: "../drawings/drawing1.xml" relative to xl/worksheets/.
        Dim drawingPath As String
        drawingPath = ResolveRelative(tempDir & "\xl\worksheets", drawingTarget)
        If Not fso.FileExists(drawingPath) Then GoTo NextSheet
        Dim drawingFileName As String: drawingFileName = fso.GetFileName(drawingPath)
        Dim drawingDir As String: drawingDir = fso.GetParentFolderName(drawingPath)
        Dim drawingRelsPath As String
        drawingRelsPath = drawingDir & "\_rels\" & drawingFileName & ".rels"

        Dim rIdToFilename As Object: Set rIdToFilename = CreateObject("Scripting.Dictionary")
        If fso.FileExists(drawingRelsPath) Then
            Dim drDoc As Object: Set drDoc = LoadXmlFile(drawingRelsPath)
            If Not drDoc Is Nothing Then
                Dim dn As Object
                For Each dn In drDoc.documentElement.childNodes
                    Dim dId As String: dId = ""
                    Dim dTgt As String: dTgt = ""
                    On Error Resume Next
                    dId = dn.getAttribute("Id")
                    dTgt = dn.getAttribute("Target")
                    On Error GoTo Fail
                    If Len(dId) > 0 And Len(dTgt) > 0 Then
                        rIdToFilename(dId) = fso.GetFileName(dTgt)
                    End If
                Next
            End If
        End If

        ' Parse drawing.xml to extract <xdr:pic> elements in order with cNvPr name + blip embed
        Dim picsInOrder As Object: Set picsInOrder = New Collection
        Dim picsByName As Object: Set picsByName = CreateObject("Scripting.Dictionary")
        Dim mediaPaths As Object: Set mediaPaths = CreateObject("Scripting.Dictionary")

        Dim drawDoc As Object: Set drawDoc = LoadXmlFile(drawingPath)
        If Not drawDoc Is Nothing Then
            ' Walk all descendants; collect <pic> elements in document order.
            CollectPicsRecursive drawDoc.documentElement, picsInOrder, picsByName, rIdToFilename
        End If

        ' Build mediaPaths for everything referenced by rIdToFilename (we copy ALL
        ' media for the sheet, picture-shape-matched or not).
        Dim rIdKey As Variant
        For Each rIdKey In rIdToFilename.Keys
            Dim fn As String: fn = CStr(rIdToFilename(rIdKey))
            Dim mediaSrc As String: mediaSrc = mediaSrcDir & "\" & fn
            If fso.FileExists(mediaSrc) And Not mediaPaths.Exists(fn) Then
                mediaPaths(fn) = mediaSrc
            End If
        Next

        If mediaPaths.Count > 0 Or picsByName.Count > 0 Or picsInOrder.Count > 0 Then
            Dim sheetInfo As Object: Set sheetInfo = CreateObject("Scripting.Dictionary")
            Set sheetInfo("picsInOrder") = picsInOrder
            Set sheetInfo("picsByName") = picsByName
            Set sheetInfo("mediaPaths") = mediaPaths
            Set perSheet(CStr(shName)) = sheetInfo
        End If
NextSheet:
    Next

    Set BuildDrawingImageMap = result
    Exit Function
Fail:
    Debug.Print "VBA Sync: BuildDrawingImageMap failed: " & Err.Description
    Set BuildDrawingImageMap = result   ' may be partially populated; tempDir is still set for cleanup
End Function

' Recursively walk a drawing.xml element tree, appending each <pic> (xdr:pic)
' encountered (document order). For each pic, extract cNvPr name and the
' embedded media filename via the blip's r:embed rId -> rIdToFilename map.
Private Sub CollectPicsRecursive(node As Object, picsInOrder As Object, _
                                  picsByName As Object, rIdToFilename As Object)
    If node Is Nothing Then Exit Sub
    On Error Resume Next
    Dim ln As String: ln = LCase$(localName(node))
    On Error GoTo 0
    If ln = "pic" Then
        ' Find blip (a:blip) and cNvPr (xdr:cNvPr) descendants
        Dim picName As String: picName = ""
        Dim embedRid As String: embedRid = ""
        FindPicAttributes node, picName, embedRid
        If Len(embedRid) > 0 Then
            If rIdToFilename.Exists(embedRid) Then
                Dim fn As String: fn = CStr(rIdToFilename(embedRid))
                picsInOrder.Add fn
                If Len(picName) > 0 And Not picsByName.Exists(picName) Then
                    picsByName(picName) = fn
                End If
            End If
        End If
        ' Don't recurse into <pic> children -- treat as leaf for our purposes.
        Exit Sub
    End If
    Dim child As Object
    For Each child In node.childNodes
        ' Skip non-element nodes
        If child.nodeType = 1 Then  ' NODE_ELEMENT
            CollectPicsRecursive child, picsInOrder, picsByName, rIdToFilename
        End If
    Next
End Sub

' Walk descendants of a <pic> element to find cNvPr name + blip embed rId.
Private Sub FindPicAttributes(picNode As Object, ByRef picName As String, ByRef embedRid As String)
    Dim child As Object
    For Each child In picNode.childNodes
        If child.nodeType = 1 Then
            Dim ln As String: ln = LCase$(localName(child))
            If ln = "nvpicpr" Then
                Dim cn As Object
                For Each cn In child.childNodes
                    If cn.nodeType = 1 Then
                        If LCase$(localName(cn)) = "cnvpr" Then
                            On Error Resume Next
                            picName = cn.getAttribute("name")
                            On Error GoTo 0
                        End If
                    End If
                Next
            ElseIf ln = "blipfill" Then
                Dim bn As Object
                For Each bn In child.childNodes
                    If bn.nodeType = 1 Then
                        If LCase$(localName(bn)) = "blip" Then
                            Dim att As Object
                            For Each att In bn.Attributes
                                If LCase$(att.baseName) = "embed" Then
                                    embedRid = att.Value
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub

' MSXML2 nodes expose .baseName as the local-name. Use a helper that handles
' both DOMDocument node styles uniformly.
Private Function localName(node As Object) As String
    On Error Resume Next
    Dim bn As String: bn = node.baseName
    If Len(bn) = 0 Then bn = node.nodeName
    ' Strip "ns:" prefix if baseName wasn't available
    Dim p As Long: p = InStr(bn, ":")
    If p > 0 Then bn = Mid$(bn, p + 1)
    localName = bn
End Function

' Synchronously extract a .zip into destDir via tar.exe (bsdtar, ships
' with Windows 10 1803+). WshShell.Run with WaitOnReturn=True blocks
' until extraction completes -- files are on disk when this returns.
' Raises if tar.exe is missing or extract fails.
Private Sub ExtractZipSynchronously(zipPath As String, destDir As String)
    Dim tEx As Double: tEx = Timer
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = "cmd /c tar.exe -xf """ & zipPath & """ -C """ & destDir & """"
    Dim exitCode As Long
    exitCode = wsh.Run(cmd, 0, True)
    TimingsAdd "extract-zip", tEx
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 9001, "ExtractZipSynchronously", _
                  "tar.exe extract failed (exit " & exitCode & _
                  "). Source=" & zipPath & " Dest=" & destDir & _
                  ". Requires Windows 10 1803+."
    End If
End Sub

' Load a small XML file via MSXML2.DOMDocument. Returns Nothing on parse failure.
Private Function LoadXmlFile(path As String) As Object
    On Error Resume Next
    Dim doc As Object: Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    If doc Is Nothing Then Set doc = CreateObject("MSXML2.DOMDocument")
    If doc Is Nothing Then Exit Function
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    doc.Load path
    If doc.parseError.errorCode <> 0 Then
        Set LoadXmlFile = Nothing
        Exit Function
    End If
    Set LoadXmlFile = doc
End Function

' Resolve "../drawings/drawing1.xml" relative to a base directory. Returns
' an absolute Windows path.
Private Function ResolveRelative(baseDir As String, relTarget As String) As String
    Dim t As String: t = Replace(relTarget, "/", "\")
    ' Strip leading "\" if present
    If Left$(t, 1) = "\" Then t = Mid$(t, 2)
    Dim b As String: b = baseDir
    If Right$(b, 1) = "\" Then b = Left$(b, Len(b) - 1)
    Do While Left$(t, 3) = "..\"
        b = Left$(b, InStrRev(b, "\") - 1)
        t = Mid$(t, 4)
    Loop
    ResolveRelative = b & "\" & t
End Function

' Delete the temp extract folder and the staging .zip. Best-effort.
Private Sub CleanupTempDir(tempDir As String)
    If Len(tempDir) = 0 Then Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If fso.FolderExists(tempDir) Then fso.DeleteFolder tempDir, True
    Dim zipPath As String: zipPath = tempDir & ".zip"
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
    On Error GoTo 0
End Sub

'====================  TABLE WRITERS  ========================
Private Sub SaveTableDataTsv(path As String, cells As Object, t As Object, exported As Object)
    Dim startCol As Long: startCol = t("startCol")
    Dim endCol As Long: endCol = t("endCol")
    Dim startRow As Long: startRow = t("startRow")
    Dim endRow As Long: endRow = t("endRow")

    Dim byRow As Object: Set byRow = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In cells.Keys
        Dim parts() As String: parts = Split(CStr(k), ",")
        Dim c As Long: c = CLng(parts(0))
        Dim r As Long: r = CLng(parts(1))
        If Not byRow.Exists(r) Then
            Dim rowDict As Object: Set rowDict = CreateObject("Scripting.Dictionary")
            Set byRow(r) = rowDict
        End If
        byRow(r)(c) = cells(CStr(k))
    Next

    Dim sb As Object: Set sb = NewStringBuilder()
    Dim headers As Object: Set headers = New Collection
    Dim c2 As Long
    If byRow.Exists(startRow) Then
        For c2 = startCol To endCol
            Dim h As String
            If byRow(startRow).Exists(c2) Then
                h = CStr(byRow(startRow)(c2))
            Else
                h = ""
            End If
            If Len(h) = 0 Then h = ColLetters(c2)
            headers.Add h
        Next
    Else
        For c2 = startCol To endCol
            headers.Add ColLetters(c2)
        Next
    End If

    Dim hi As Long, hFirst As Boolean: hFirst = True
    Dim hVal As Variant
    For Each hVal In headers
        If hFirst Then hFirst = False Else sbAppend sb, vbTab
        sbAppend sb, TsvEscape(CStr(hVal))
    Next
    sbAppend sb, vbCrLf

    Dim r2 As Long
    For r2 = startRow + 1 To endRow
        Dim first As Boolean: first = True
        For c2 = startCol To endCol
            If Not first Then sbAppend sb, vbTab Else first = False
            If byRow.Exists(r2) Then
                If byRow(r2).Exists(c2) Then
                    sbAppend sb, TsvEscape(CStr(byRow(r2)(c2)))
                End If
            End If
        Next
        sbAppend sb, vbCrLf
    Next

    WriteIfChanged path, sbToString(sb), exported, True
End Sub

Private Sub SaveTableDefinitionJson(path As String, t As Object, cellFormulas As Object, cellValues As Object, tableCalcFormulas As Object, exported As Object)
    Dim lo As ListObject: Set lo = t("listObject")
    Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
    obj("name") = lo.Name
    On Error Resume Next
    obj("displayName") = lo.DisplayName
    If obj("displayName") = "" Then obj("displayName") = lo.Name
    On Error GoTo 0
    obj("totalsRowShown") = lo.ShowTotals

    ' Look up this table's column formulas in the OOXML map (authoritative
    ' source for calc-column formulas, especially on zero-row tables where
    ' DataBodyRange Is Nothing and we have no cell to read from).
    Dim ooxmlColMap As Object: Set ooxmlColMap = Nothing
    If Not tableCalcFormulas Is Nothing Then
        If tableCalcFormulas.Exists(LCase$(lo.Name)) Then
            Set ooxmlColMap = tableCalcFormulas(LCase$(lo.Name))
        End If
    End If

    Dim columns As Object: Set columns = New Collection
    Dim colIdx As Long: colIdx = 0
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        Dim col As Object: Set col = CreateObject("Scripting.Dictionary")
        col("name") = lc.Name

        ' Calculated column formula? Excel auto-fills a calc-column formula on
        ' row-add; users can break out on individual cells, which we track as
        ' `overrides`. Strategy:
        '   1. Scan data cells for the first formulated cell; use as default.
        '   2. If no formulated cell found (zero-row table, or every cell is
        '      a literal override), fall back to the OOXML
        '      <calculatedColumnFormula> from xl/tables/<table>.xml.
        Dim calcFormula As String: calcFormula = ""
        On Error Resume Next
        If Not lc.DataBodyRange Is Nothing Then
            Dim scanRow As Long, dbrRows As Long
            dbrRows = lc.DataBodyRange.Rows.Count
            For scanRow = 1 To dbrRows
                Dim candF As String: candF = lc.DataBodyRange.Cells(scanRow, 1).Formula
                If Len(candF) > 1 And Left$(candF, 1) = "=" Then
                    calcFormula = candF
                    Exit For
                End If
            Next
        End If
        On Error GoTo 0

        ' OOXML fallback
        If Len(calcFormula) = 0 And Not ooxmlColMap Is Nothing Then
            If ooxmlColMap.Exists(lc.Name) Then
                calcFormula = CStr(ooxmlColMap(lc.Name))
            End If
        End If

        If Len(calcFormula) > 0 Then
            col("formula") = calcFormula
            ' Find overrides: data cells whose formula differs from the column
            ' default OR cells that have a literal value where the default has
            ' a formula. Both cases get tracked.
            Dim sheetColIdx As Long: sheetColIdx = t("startCol") + colIdx
            Dim overrides As Object: Set overrides = CreateObject("Scripting.Dictionary")
            Dim r As Long
            For r = t("startRow") + 1 To t("endRow")
                Dim k As String: k = CStr(sheetColIdx) & "," & CStr(r)
                Dim tableRowIdx As Long: tableRowIdx = r - t("startRow")
                If cellFormulas.Exists(k) Then
                    Dim cellF As String: cellF = CStr(cellFormulas(k))
                    If cellF <> calcFormula Then
                        overrides(CStr(tableRowIdx)) = cellF
                    End If
                Else
                    ' Cell has no formula at all -- it's a literal value
                    ' overriding the calc column. Record the literal so the
                    ' override is reversible and reviewable in the diff.
                    Dim litVal As Variant: litVal = Empty
                    If cellValues.Exists(k) Then litVal = cellValues(k)
                    If IsEmpty(litVal) Then
                        overrides(CStr(tableRowIdx)) = ""   ' explicit blank override
                    Else
                        overrides(CStr(tableRowIdx)) = litVal
                    End If
                End If
            Next
            If overrides.Count > 0 Then Set col("overrides") = overrides
        End If

        columns.Add col
        colIdx = colIdx + 1
    Next
    Set obj("columns") = columns
    obj("ref") = t("ref")

    WriteIfChanged path, JsonPretty(obj), exported, False
End Sub

'====================  DEFINED NAMES + LAMBDAS  ==============
Private Sub ExtractLambdas(wb As Workbook, ByRef lambdas As Object)
    ' Walk wb.Names + per-sheet Names. Find ones whose RefersTo starts with
    ' "=_xlfn.LAMBDA(". Dedup by Name. RefersTo includes leading "=", strip.
    Dim n As Name
    For Each n In wb.Names
        AddLambdaIfPresent n, lambdas
    Next
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim sn As Name
        For Each sn In ws.Names
            AddLambdaIfPresent sn, lambdas
        Next
    Next
End Sub

Private Sub AddLambdaIfPresent(n As Name, ByRef lambdas As Object)
    On Error Resume Next
    Dim refersTo As String: refersTo = n.refersTo
    On Error GoTo 0
    If Len(refersTo) = 0 Then Exit Sub
    ' Strip leading "="
    Dim body As String: body = refersTo
    If Left$(body, 1) = "=" Then body = Mid$(body, 2)

    ' Heuristic: starts with _xlfn.LAMBDA( (optionally with whitespace)
    Dim trimmed As String: trimmed = LTrim$(body)
    If Left$(trimmed, 13) <> "_xlfn.LAMBDA(" Then Exit Sub

    ' Strip "name@SheetName" — use plain Name. Sheet-local copies dedup.
    Dim plainName As String: plainName = n.Name
    Dim eqPos As Long: eqPos = InStr(plainName, "!")
    If eqPos > 0 Then plainName = Mid$(plainName, eqPos + 1)

    If Not lambdas.Exists(plainName) Then
        lambdas(plainName) = body
    End If
End Sub

Private Sub WriteLambdasFolder(dir As String, lambdas As Object, exported As Object)
    If lambdas.Count = 0 Then Exit Sub
    EnsureFolder dir
    Dim k As Variant
    For Each k In lambdas.Keys
        Dim safe As String: safe = SafeFileName(CStr(k))
        WriteIfChanged dir & safe & ".lambda", CStr(lambdas(k)), exported, False
    Next
End Sub

Private Function BuildDefinedNames(wb As Workbook, lambdas As Object) As Object
    ' Returns Dictionary keyed by name (or name@SheetName for sheet-local).
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary")
    Dim n As Name
    For Each n In wb.Names
        On Error Resume Next
        Dim rt As String: rt = n.refersTo
        On Error GoTo 0
        If Len(rt) = 0 Then GoTo NextG
        ' Skip lambdas (already extracted)
        Dim bodyG As String: bodyG = rt
        If Left$(bodyG, 1) = "=" Then bodyG = Mid$(bodyG, 2)
        If Left$(LTrim$(bodyG), 13) = "_xlfn.LAMBDA(" Then GoTo NextG

        Dim nameG As String: nameG = n.Name
        Dim bangG As Long: bangG = InStr(nameG, "!")
        If bangG > 0 Then nameG = Mid$(nameG, bangG + 1)
        out(nameG) = rt
NextG:
    Next
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim sn As Name
        For Each sn In ws.Names
            On Error Resume Next
            Dim srt As String: srt = sn.refersTo
            On Error GoTo 0
            If Len(srt) = 0 Then GoTo NextS
            Dim bodyS As String: bodyS = srt
            If Left$(bodyS, 1) = "=" Then bodyS = Mid$(bodyS, 2)
            If Left$(LTrim$(bodyS), 13) = "_xlfn.LAMBDA(" Then GoTo NextS

            Dim nameS As String: nameS = sn.Name
            Dim bangS As Long: bangS = InStr(nameS, "!")
            If bangS > 0 Then nameS = Mid$(nameS, bangS + 1)
            out(nameS & "@" & ws.Name) = srt
NextS:
        Next
    Next

    ' Sort keys alphabetically by rebuilding dictionary
    Dim sorted As Object: Set sorted = CreateObject("Scripting.Dictionary")
    Dim keys As Object: Set keys = SortedKeys(out)
    Dim k As Variant
    For Each k In keys
        sorted(CStr(k)) = out(CStr(k))
    Next
    Set BuildDefinedNames = sorted
End Function

'====================  MANIFEST HELPERS  =====================
Private Function SheetIdFromCodeName(codeName As String, fallback As Long) As Long
    ' "Sheet1" -> 1, "Sheet13" -> 13. Sheets renamed by the user keep their
    ' codeName, so this is stable. Returns fallback (tab index) if unparseable.
    Dim m As String
    If Left$(codeName, 5) = "Sheet" Then
        m = Mid$(codeName, 6)
        If IsNumeric(m) And Len(m) > 0 Then
            SheetIdFromCodeName = CLng(m)
            Exit Function
        End If
    End If
    SheetIdFromCodeName = fallback
End Function

Private Function SheetStateString(v As XlSheetVisibility) As String
    Select Case v
        Case -1: SheetStateString = "visible"    ' xlSheetVisible
        Case 0: SheetStateString = "hidden"    ' xlSheetHidden
        Case 2: SheetStateString = "veryHidden"    ' xlSheetVeryHidden
        Case Else: SheetStateString = "visible"
    End Select
End Function

Private Function TabColorHex(ws As Worksheet) As Variant
    On Error Resume Next
    Dim c As Variant: c = ws.Tab.Color
    On Error GoTo 0
    If IsEmpty(c) Or IsNull(c) Then
        TabColorHex = Null
    ElseIf CLng(c) = -4142 Or CLng(c) = 0 Then
        ' xlColorIndexNone or default
        TabColorHex = Null
    Else
        TabColorHex = OleColorToHex(CLng(c))
    End If
End Function

Private Function CalcModeString(m As XlCalculation) As String
    Select Case m
        Case -4135: CalcModeString = "manual"    ' xlCalculationManual
        Case -4105: CalcModeString = "auto"    ' xlCalculationAutomatic
        Case 2: CalcModeString = "autoNoTable"    ' xlCalculationSemiautomatic
        Case Else: CalcModeString = "auto"
    End Select
End Function

Private Function SortTableEntries(list As Object) As Object
    ' Sort collection of dicts by (sheet, name) ascending. Bubble for simplicity.
    Dim n As Long: n = list.Count
    If n <= 1 Then Set SortTableEntries = list: Exit Function
    Dim arr() As Variant: ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1: Set arr(i) = list(i + 1): Next
    Dim j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            Dim aKey As String: aKey = arr(j)("sheet") & vbTab & arr(j)("name")
            Dim bKey As String: bKey = arr(j + 1)("sheet") & vbTab & arr(j + 1)("name")
            If aKey > bKey Then
                Dim tmp As Object: Set tmp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = tmp
            End If
        Next
    Next
    Dim out As Collection: Set out = New Collection
    For i = 0 To n - 1: out.Add arr(i): Next
    Set SortTableEntries = out
End Function

Private Sub SortDictListBySqref(ByRef list As Object)
    Dim n As Long: n = list.Count
    If n <= 1 Then Exit Sub
    Dim arr() As Variant: ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1: Set arr(i) = list(i + 1): Next
    Dim j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If CStr(arr(j)("sqref")) > CStr(arr(j + 1)("sqref")) Then
                Dim tmp As Object: Set tmp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = tmp
            End If
        Next
    Next
    Set list = New Collection
    For i = 0 To n - 1: list.Add arr(i): Next
End Sub

Private Sub SortDictListBySqrefPriority(ByRef list As Object)
    Dim n As Long: n = list.Count
    If n <= 1 Then Exit Sub
    Dim arr() As Variant: ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1: Set arr(i) = list(i + 1): Next
    Dim j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            Dim aSq As String: aSq = CStr(arr(j)("sqref"))
            Dim bSq As String: bSq = CStr(arr(j + 1)("sqref"))
            Dim swap As Boolean: swap = False
            If aSq > bSq Then
                swap = True
            ElseIf aSq = bSq Then
                Dim aP As Long: aP = 0
                Dim bP As Long: bP = 0
                If arr(j).Exists("priority") Then aP = CLng(arr(j)("priority"))
                If arr(j + 1).Exists("priority") Then bP = CLng(arr(j + 1)("priority"))
                If aP > bP Then swap = True
            End If
            If swap Then
                Dim tmp As Object: Set tmp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = tmp
            End If
        Next
    Next
    Set list = New Collection
    For i = 0 To n - 1: list.Add arr(i): Next
End Sub

Private Function SortMergedByTopLeft(list As Object) As Object
    ' Sort collection of {range, value} dicts by top-left row*100000+col.
    Dim n As Long: n = list.Count
    If n <= 1 Then Set SortMergedByTopLeft = list: Exit Function
    Dim arr() As Variant: ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1: Set arr(i) = list(i + 1): Next
    Dim j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            Dim aStart As String: aStart = Split(CStr(arr(j)("range")), ":")(0)
            Dim bStart As String: bStart = Split(CStr(arr(j + 1)("range")), ":")(0)
            Dim aKey As Long: aKey = AddressSortKeyRowCol(aStart)
            Dim bKey As Long: bKey = AddressSortKeyRowCol(bStart)
            If aKey > bKey Then
                Dim tmp As Object: Set tmp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = tmp
            End If
        Next
    Next
    Dim out As Collection: Set out = New Collection
    For i = 0 To n - 1: out.Add arr(i): Next
    Set SortMergedByTopLeft = out
End Function

Private Function AddressSortKeyRowCol(addr As String) As Long
    ' Same as AddressSortKey but row*100000+col (matches PS Compress-CellsToRanges sort)
    Dim i As Long
    Dim letters As String, digits As String
    For i = 1 To Len(addr)
        Dim ch As String: ch = Mid$(addr, i, 1)
        If ch >= "A" And ch <= "Z" Then
            letters = letters & ch
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If
    Next
    Dim col As Long: col = 0
    For i = 1 To Len(letters)
        col = col * 26 + (Asc(Mid$(letters, i, 1)) - 64)
    Next
    Dim row As Long: row = 0
    If Len(digits) > 0 Then row = CLng(digits)
    AddressSortKeyRowCol = row * 100000 + col
End Function

Private Function SortedKeys(d As Object) As Object
    Dim n As Long: n = d.Count
    Dim out As Collection: Set out = New Collection
    If n = 0 Then Set SortedKeys = out: Exit Function
    Dim arr() As String: ReDim arr(0 To n - 1)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In d.Keys: arr(i) = CStr(k): i = i + 1: Next
    SortStringArray arr
    For i = 0 To n - 1: out.Add arr(i): Next
    Set SortedKeys = out
End Function

Private Function SortedNumericKeys(d As Object) As Object
    ' For Dictionary keyed by Long row numbers. Returns Collection sorted ascending.
    Dim n As Long: n = d.Count
    Dim out As Collection: Set out = New Collection
    If n = 0 Then Set SortedNumericKeys = out: Exit Function
    Dim arr() As Long: ReDim arr(0 To n - 1)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In d.Keys: arr(i) = CLng(k): i = i + 1: Next
    SortLongArray arr
    For i = 0 To n - 1: out.Add arr(i): Next
    Set SortedNumericKeys = out
End Function

Private Sub SortStringCollection(ByRef list As Object)
    Dim n As Long: n = list.Count
    If n <= 1 Then Exit Sub
    Dim arr() As String: ReDim arr(0 To n - 1)
    Dim i As Long: i = 0
    Dim x As Variant
    For Each x In list: arr(i) = CStr(x): i = i + 1: Next
    SortStringArray arr
    Set list = New Collection
    For i = 0 To n - 1: list.Add arr(i): Next
End Sub

Private Sub SortStringArray(ByRef arr() As String)
    Dim n As Long: n = UBound(arr) - LBound(arr) + 1
    If n <= 1 Then Exit Sub
    Dim lb As Long: lb = LBound(arr)
    Dim i As Long, j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(lb + j) > arr(lb + j + 1) Then
                Dim tmp As String: tmp = arr(lb + j)
                arr(lb + j) = arr(lb + j + 1)
                arr(lb + j + 1) = tmp
            End If
        Next
    Next
End Sub

Private Sub SortLongArray(ByRef arr() As Long)
    Dim n As Long: n = UBound(arr) - LBound(arr) + 1
    If n <= 1 Then Exit Sub
    Dim lb As Long: lb = LBound(arr)
    Dim i As Long, j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(lb + j) > arr(lb + j + 1) Then
                Dim tmp As Long: tmp = arr(lb + j)
                arr(lb + j) = arr(lb + j + 1)
                arr(lb + j + 1) = tmp
            End If
        Next
    Next
End Sub

Private Function JoinCollection(list As Object, sep As String) As String
    Dim s As String, first As Boolean: first = True
    Dim x As Variant
    For Each x In list
        If first Then first = False Else s = s & sep
        s = s & CStr(x)
    Next
    JoinCollection = s
End Function

'====================  IO UTILITIES  =========================
Private Sub EnsureFolder(p As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parts() As String, accum As String
    parts = Split(p, "\")
    Dim i As Long
    For i = 0 To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo NextPart
        If Len(accum) = 0 Then
            accum = parts(i)
        Else
            accum = accum & "\" & parts(i)
        End If
        ' Skip drive letter (e.g. "C:")
        If Right$(accum, 1) = ":" Then GoTo NextPart
        If Not fso.FolderExists(accum) Then
            On Error Resume Next
            fso.CreateFolder accum
            On Error GoTo 0
        End If
NextPart:
    Next
End Sub

' WriteIfChanged: writes content to path only if the existing file's content
' differs. Skips otherwise. Records the path in exported regardless.
' useCRLF=True normalises LF to CRLF (for TSV/MD); False uses LF (for JSON/lambda).
Public Sub WriteIfChanged(path As String, content As String, exported As Object, useCRLF As Boolean)
    Dim normalised As String: normalised = NormaliseLineEndings(content, useCRLF)

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parent As String: parent = Left$(path, InStrRev(path, "\"))
    If Len(parent) > 0 Then EnsureFolder parent

    Dim existingMatches As Boolean: existingMatches = False
    If fso.FileExists(path) Then
        Dim ex As String
        On Error Resume Next
        ex = ReadFileBytes(path)
        On Error GoTo 0
        If ex = normalised Then existingMatches = True
    End If

    If Not existingMatches Then
        WriteFileBytes path, normalised
    End If

    If Not exported Is Nothing Then
        exported(AddSlashLocal(path)) = True
    End If
End Sub

Private Function AddSlashLocal(p As String) As String
    AddSlashLocal = Replace(p, "/", "\")
End Function

Private Function NormaliseLineEndings(s As String, useCRLF As Boolean) As String
    ' Normalise to LF first, then optionally to CRLF.
    Dim out As String: out = s
    out = Replace(out, vbCrLf, vbLf)
    out = Replace(out, vbCr, vbLf)
    If useCRLF Then out = Replace(out, vbLf, vbCrLf)
    NormaliseLineEndings = out
End Function

' Read file as a String. Uses ADO Stream for UTF-8.
Private Function ReadFileBytes(path As String) As String
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2  ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadFileBytes = stm.ReadText
    stm.Close
End Function

' Write file as UTF-8 (no BOM). Caller passes already line-ending-normalised string.
Private Sub WriteFileBytes(path As String, content As String)
    ' Write UTF-8 without BOM via ADO Stream
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2  ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    ' Convert to binary so we can strip the BOM ADO writes by default
    stm.Position = 0
    stm.Type = 1  ' adTypeBinary
    stm.Position = 3  ' skip BOM
    Dim binStm As Object: Set binStm = CreateObject("ADODB.Stream")
    binStm.Type = 1
    binStm.Open
    stm.CopyTo binStm
    stm.Close
    binStm.SaveToFile path, 2  ' adSaveCreateOverWrite
    binStm.Close
End Sub

'====================  STRING / FORMAT UTILITIES  ============
Private Function NewStringBuilder() As Object
    ' Use a Collection as a fast string builder (append-only, then Join).
    Set NewStringBuilder = New Collection
End Function

Private Sub sbAppend(sb As Object, s As String)
    sb.Add s
End Sub

Private Function sbToString(sb As Object) As String
    Dim n As Long: n = sb.Count
    If n = 0 Then sbToString = "": Exit Function
    Dim arr() As String: ReDim arr(0 To n - 1)
    Dim i As Long: i = 0
    Dim x As Variant
    For Each x In sb
        arr(i) = CStr(x)
        i = i + 1
    Next
    sbToString = Join(arr, "")
End Function

Private Function ColLetters(col As Long) As String
    Dim s As String, n As Long: n = col
    Do While n > 0
        Dim r As Long: r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColLetters = s
End Function

Private Function PadLeft(s As String, w As Long, ch As String) As String
    Dim out As String: out = s
    Do While Len(out) < w
        out = ch & out
    Loop
    PadLeft = out
End Function

Private Function SafeFileName(name As String) As String
    ' Percent-encode illegal Windows filename chars (and % itself).
    If Len(name) = 0 Then SafeFileName = "_": Exit Function
    Dim sb As String
    Dim i As Long
    For i = 1 To Len(name)
        Dim ch As String: ch = Mid$(name, i, 1)
        Dim escape As Boolean: escape = False
        Select Case ch
            Case "%", "/", "\", ":", "*", "?", """", "<", ">", "|": escape = True
        End Select
        If escape Then
            sb = sb & "%" & UCase(Right$("00" & Hex(Asc(ch)), 2))
        Else
            sb = sb & ch
        End If
    Next
    ' Trim trailing dots / spaces (illegal on Windows)
    Do While Len(sb) > 0
        Dim last As String: last = Right$(sb, 1)
        If last = " " Or last = "." Then
            sb = Left$(sb, Len(sb) - 1)
        Else
            Exit Do
        End If
    Loop
    If Len(sb) > 100 Then sb = Left$(sb, 100)
    If Len(sb) = 0 Then sb = "_"
    SafeFileName = sb
End Function

Private Function TsvEscape(s As String) As String
    Dim out As String: out = s
    out = Replace(out, "\", "\\")
    out = Replace(out, vbTab, "\t")
    out = Replace(out, vbLf, "\n")
    out = Replace(out, vbCr, "\r")
    TsvEscape = out
End Function

'====================  JSON PRETTY-PRINT  ====================
' Use JsonConverter.ConvertToJson with whitespace=2 for canonical 2-space
' indent, then post-process to collapse empty containers to "{}" / "[]"
' (matches the PS Format-Json output).
Private Function JsonPretty(obj As Object) As String
    Dim s As String
    s = JsonConverter.ConvertToJson(obj, Whitespace:=2)
    ' Collapse empty containers (JsonConverter renders them with a newline+indent inside)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{\s*\r?\n\s*\}"
    s = re.Replace(s, "{}")
    re.Pattern = "\[\s*\r?\n\s*\]"
    s = re.Replace(s, "[]")
    JsonPretty = s
End Function
