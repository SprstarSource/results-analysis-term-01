Attribute VB_Name = "modBuildPanel01"
Option Explicit

Private m_workflowBusy As Boolean

' Import only the matching term panel into each independent tool workbook.
Private Const PANEL_SHEET As String = "ControlPanel"
Private Const RESULTS_FILE As String = "Term01Results.xlsx"
Private Const GUIDANCE_FILE As String = "Term01Guidance.xlsx"
Private Const EXPORT_FILE As String = "OrientationResults01.xlsx"

' Work folder name (hex code points): Term01Results.xlsx, Term01Guidance.xlsx and OrientationResults01.xlsx live in it
Private Const WF_FOLDER_CODES As String = "645 639 627 644 62C 629 20 628 64A 627 646 627 62A 20 627 644 641 635 644 20 30 31"
Private Const BUTTON_WIDTH As Double = 164
Private Const BUTTON_HEIGHT As Double = 54
Private Const BUTTON_GAP As Double = 12

' Term 01 Analysis panel (merged from the former modBiuldAnalysisfilesPanel01)
Private Const ANALYSIS_PANEL_SHEET As String = "Term01Panel"
Private Const BTN_MAIN As String = "btnTerm01Analysis"
Private Const BTN_OPEN As String = "btnOpenOutput"
Private Const ROW_FIRST As Long = 16

Private Function PanelActions() As Variant
    PanelActions = Array( _
        Array("btnCollect", "First Term Results", "CollectTerm01Results", False, &H60AE27), _
        Array("btnTerm01Analysis", "Term 01 Analysis", "modAnalysisBuilderBuildAll01.Build_All", False, &HB98029), _
        Array("btnExamen", "Examen Analysis", "modExamenAnalysis01.Run_ExamenAnalysis01", False, &HA05568), _
        Array("btnFirstTermOrientation", "First Term Orientation", "RunTermAll01", False, &HB98029))
End Function

Public Sub BuildControlPanel()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, count As Long, columns As Long, gridRow As Long, gridCol As Long
    Dim rowItems As Long, canvasWidth As Double, rowWidth As Double
    Dim originLeft As Double, originTop As Double, x As Double, y As Double
    Dim prevScreen As Boolean, errorText As String

    prevScreen = Application.ScreenUpdating
    On Error GoTo BuildError
    Application.ScreenUpdating = False
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 801, , "The tool workbook is read-only."
    Set ws = PanelSheet()
    If ws Is Nothing Then
        If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 802, , "The workbook structure is protected."
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET
    End If
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        Err.Raise vbObjectError + 803, , "Unprotect the control panel sheet before rebuilding it."
    End If
    ws.Activate
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    With ws
        .Cells.UnMerge
        .Cells.Clear
        .DisplayRightToLeft = True
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 5
        .Columns("C").ColumnWidth = 22
        .Columns("D").ColumnWidth = 14
        .Columns("E").ColumnWidth = 16
        .Columns("F:K").ColumnWidth = 22
        .Rows(1).RowHeight = 48
        .Rows(2).RowHeight = 14
        With .Range("B1:K1")
            .Merge
            .Value = Panel_Ar("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 30 31") & "  |  First Term Management Dashboard"
            .Font.Name = "Segoe UI"
            .Font.Size = 19
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(28, 55, 103)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End With

    actions = PanelActions()
    count = UBound(actions) - LBound(actions) + 1
    columns = 2
    canvasWidth = columns * BUTTON_WIDTH + (columns - 1) * BUTTON_GAP
    originLeft = ws.Range("F5").Left + (ws.Range("F5:K5").Width - canvasWidth) / 2
    originTop = ws.Range("F5").Top
    For i = 0 To count - 1
        item = actions(i)
        gridRow = i \ columns
        gridCol = i Mod columns
        rowItems = count - gridRow * columns
        If rowItems > columns Then rowItems = columns
        rowWidth = rowItems * BUTTON_WIDTH + (rowItems - 1) * BUTTON_GAP
        x = originLeft + (canvasWidth - rowWidth) / 2 + gridCol * (BUTTON_WIDTH + BUTTON_GAP)
        y = originTop + gridRow * (BUTTON_HEIGHT + BUTTON_GAP)
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, BUTTON_WIDTH, BUTTON_HEIGHT)
        With btn
            .Name = CStr(item(0))
            .Adjustments(1) = 0.12
            .Placement = xlFreeFloating
            .Line.Weight = 1.25
            With .TextFrame2
                .TextRange.Text = CStr(item(1))
                .TextRange.Font.Name = "Segoe UI"
                .TextRange.Font.Size = 12
                .TextRange.Font.Bold = msoTrue
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .VerticalAnchor = msoAnchorMiddle
                .WordWrap = msoTrue
                .MarginLeft = 7
                .MarginRight = 7
                .MarginTop = 4
                .MarginBottom = 4
            End With
        End With
    Next i
    y = originTop + ((count + columns - 1) \ columns) * (BUTTON_HEIGHT + BUTTON_GAP)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, originLeft, y, canvasWidth, 40)
    With btn
        .Name = "btnOpenOutput"
        .Fill.Solid: .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Line.ForeColor.RGB = RGB(89, 89, 89)
        .OnAction = PanelMacroAction("modBuildPanel01.Panel_OpenOutput")
        With .TextFrame2
            .TextRange.Text = Panel_Ar("641 62A 62D 20 645 62C 644 62F 20 627 644 625 635 62F 627 631")
            .TextRange.Font.Size = 14: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With
    Panel_BuildOutputTables ws
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    RefreshPanelState
    CheckPanelMacros actions
    Application.ScreenUpdating = prevScreen
    MsgBox "Control panel updated: " & count & " buttons (" & columns & " per row).", vbInformation, "Done"
    Exit Sub
BuildError:
    errorText = CStr(Err.Number) & " : " & Err.Description
    Application.ScreenUpdating = prevScreen
    MsgBox "Panel build failed: " & errorText, vbExclamation, "Error"
End Sub

Public Sub RefreshPanelState()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, enabled As Boolean, resultsExist As Boolean, stage As Long, stepNo As Long
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    resultsExist = PanelResultsExist()
    stage = TermWorkflowStage01()
    actions = PanelActions()
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        Set btn = Nothing
        On Error Resume Next
        Set btn = ws.Shapes(CStr(item(0)))
        Err.Clear
        On Error GoTo 0
        If Not btn Is Nothing Then
            enabled = True
            If CBool(item(3)) Then enabled = resultsExist
            stepNo = WorkflowStep01(CStr(item(0)))
            If stepNo > 1 Then enabled = (stage >= stepNo - 1)
            With btn
                .Fill.Solid
                If enabled Then
                    .Fill.ForeColor.RGB = CLng(item(4))
                    .Line.ForeColor.RGB = CLng(item(4))
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    If stepNo > 0 Then
                        .OnAction = PanelMacroAction("RunTermStep01")
                    Else
                        .OnAction = PanelMacroAction(CStr(item(2)))
                    End If
                Else
                    .Fill.ForeColor.RGB = RGB(220, 220, 220)
                    .Line.ForeColor.RGB = RGB(180, 180, 180)
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
                    If stepNo > 0 Then
                        .OnAction = PanelMacroAction("RunTermStep01")
                    Else
                        .OnAction = PanelMacroAction("ResultsGuard")
                    End If
                End If
            End With
        End If
    Next i
    Panel_BuildOutputTables ws
End Sub

Private Sub Panel_BuildOutputTables(ByVal ws As Worksheet)
    Dim map As Object, key As Variant, fso As Object, folder As Object, f As Object
    Dim r As Long, startRow As Long, outDir As String, statusText As String, stage As Long
    On Error Resume Next
    startRow = 4
    ws.Range("B" & startRow & ":D500").Clear
    stage = TermWorkflowStage01()
    Panel_TableTitle ws, startRow, "Orientation steps"
    Panel_TableHeader ws, startRow + 1
    Panel_TableRow ws, startRow + 2, 1, "Data transfer", IIf(stage >= 1, "Done", "Pending")
    Panel_TableRow ws, startRow + 3, 2, "Stream preferences", IIf(stage >= 2, "Done", "Pending")
    Panel_TableRow ws, startRow + 4, 3, "Orientation", IIf(stage >= 3, "Done", "Pending")
    Panel_TableRow ws, startRow + 5, 4, "Export results", IIf(stage >= 4, "Done", "Pending")

    startRow = startRow + 7
    Panel_TableTitle ws, startRow, "Analysis files"
    Panel_TableHeader ws, startRow + 1
    Set map = GetFileMap()
    r = startRow + 2
    outDir = ThisWorkbook.Path & Application.PathSeparator & WfFolderName()
    For Each key In map.Keys
        statusText = IIf(Panel_HasOutputStem(outDir, CStr(key)), "Done", "Pending")
        Panel_TableRow ws, r, r - startRow - 1, CStr(key), statusText
        r = r + 1
    Next key

    startRow = r + 1
    Panel_TableTitle ws, startRow, "Other button outputs"
    Panel_TableHeader ws, startRow + 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(outDir) Then
        Set folder = fso.GetFolder(outDir)
        r = startRow + 2
        For Each f In folder.Files
            If LCase$(fso.GetExtensionName(f.Name)) Like "xls*" Then
                If Not Panel_IsAnalysisStem(map, f.Name) Then
                    Panel_TableRow ws, r, r - startRow - 1, f.Name, "Available"
                    r = r + 1
                End If
            End If
        Next f
    End If
    Err.Clear
End Sub

Private Function Panel_HasOutputStem(ByVal folder As String, ByVal stem As String) As Boolean
    Dim f As String
    f = Dir$(folder & Application.PathSeparator & stem & "*.xlsx")
    Panel_HasOutputStem = (Len(f) > 0)
End Function

Private Function Panel_IsAnalysisStem(ByVal map As Object, ByVal fileName As String) As Boolean
    Dim key As Variant
    For Each key In map.Keys
        If StrComp(Left$(fileName, Len(CStr(key))), CStr(key), vbTextCompare) = 0 Then
            Panel_IsAnalysisStem = True
            Exit Function
        End If
    Next key
End Function

Private Sub Panel_TableTitle(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Merge: .Value2 = title: .Font.Bold = True: .Font.Size = 14
        .Font.Color = RGB(28, 55, 103): .HorizontalAlignment = xlRight
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ws.Rows(r).RowHeight = 25
End Sub

Private Sub Panel_TableHeader(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, 2).Value2 = "#": ws.Cells(r, 3).Value2 = "File / step": ws.Cells(r, 4).Value2 = "Status": ws.Cells(r, 5).Value2 = "Duration (s)"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Font.Bold = True: .Interior.Color = RGB(28, 55, 103): .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Function Panel_Ar(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    Panel_Ar = s
End Function

Private Sub Panel_TableRow(ByVal ws As Worksheet, ByVal r As Long, ByVal idx As Long, ByVal label As String, ByVal statusText As String)
    ws.Cells(r, 2).Value2 = idx: ws.Cells(r, 3).Value2 = label: ws.Cells(r, 4).Value2 = statusText
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(191, 191, 191)
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub ResultsGuard()
    MsgBox "Run First Term Results first to generate " & RESULTS_FILE & ".", _
           vbExclamation, "Step 1 Required"
End Sub

' Compatibility with buttons created by earlier panel versions.
Public Sub AverageGuard()
    ResultsGuard
End Sub
Public Sub FirstTermOnlyGuard()
    ResultsGuard
End Sub
Public Sub SecondTermOnlyGuard()
    ResultsGuard
End Sub
Public Sub ThirdTermOnlyGuard()
    ResultsGuard
End Sub

Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

Private Function PanelResultsExist() As Boolean
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Function
    PanelResultsExist = WfFileExists(WfResolvePath(RESULTS_FILE))
End Function

Private Function PanelMacroAction(ByVal macroName As String) As String
    PanelMacroAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

Private Sub CheckPanelMacros(ByVal actions As Variant)
    Dim i As Long, item As Variant, missing As String
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        If Not PanelMacroExists(CStr(item(2))) Then missing = missing & "  - " & CStr(item(2)) & vbCrLf
    Next i
    If Len(missing) > 0 Then
        MsgBox "Import the missing modules before using these actions:" & vbCrLf & missing, _
               vbExclamation, "Missing Macros"
    End If
End Sub

Private Function PanelMacroExists(ByVal macroName As String) As Boolean
    Dim comp As Object, codeMod As Object, txt As String
    Dim moduleName As String, procedureName As String, dotPos As Long
    On Error GoTo AccessDenied
    dotPos = InStrRev(macroName, ".")
    procedureName = macroName
    If dotPos > 0 Then
        moduleName = Left$(macroName, dotPos - 1)
        procedureName = Mid$(macroName, dotPos + 1)
    End If
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then
            If Len(moduleName) = 0 Or StrComp(comp.Name, moduleName, vbTextCompare) = 0 Then
                Set codeMod = comp.CodeModule
                If codeMod.CountOfLines > 0 Then
                    txt = codeMod.Lines(1, codeMod.CountOfLines)
                    If InStr(1, txt, "Sub " & procedureName & "(", vbTextCompare) > 0 Then
                        PanelMacroExists = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next comp
    Exit Function
AccessDenied:
    ' Macro inspection is optional when VBA project access is disabled.
    PanelMacroExists = True
    Err.Clear
End Function

Private Function WorkflowStep01(ByVal shapeName As String) As Long
    Select Case shapeName
    Case "btnDataTransfer": WorkflowStep01 = 1
    Case "btnStreamPreferences": WorkflowStep01 = 2
    Case "btnExportOrientationResults": WorkflowStep01 = 4
    End Select
End Function

Public Sub RunTermAll01()
    Dim errorText As String
    If m_workflowBusy Then Exit Sub
    On Error GoTo Failed
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 810, , "The tool workbook is read-only."
    m_workflowBusy = True
    WorkflowSaveStage01 0
    RefreshPanelState
    Application.Run PanelMacroAction("modOrientationAll01.Run_OrientationAll01")
    If modOrientationAll01.OrientationAll01Succeeded() Then WorkflowSaveStage01 4
Finished:
    m_workflowBusy = False
    On Error Resume Next
    RefreshPanelState
    On Error GoTo 0
    If Len(errorText) > 0 Then MsgBox errorText, vbExclamation, "Term 01 guidance"
    Exit Sub
Failed:
    errorText = Err.Description
    Resume Finished
End Sub

Private Function WorkflowRead01(ByVal key As String) As String
    Dim s As String
    On Error GoTo Missing
    s = ThisWorkbook.Names("_Term01_" & key).RefersTo
    If Left$(s, 2) = "=" & Chr$(34) And Right$(s, 1) = Chr$(34) Then
        WorkflowRead01 = Replace(Mid$(s, 3, Len(s) - 3), Chr$(34) & Chr$(34), Chr$(34))
    End If
Missing:
End Function

Private Sub WorkflowWrite01(ByVal key As String, ByVal value As String)
    On Error Resume Next
    ThisWorkbook.Names("_Term01_" & key).Delete
    Err.Clear
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="_Term01_" & key, RefersTo:="=" & Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34), Visible:=False
End Sub

' ---- Work-folder path helpers: the files produced by the tool live in the work folder ----
' (the producing modules create the folder; this module only looks files up)
Private Function WfFolderName() As String
    Dim parts() As String, i As Long, t As String
    parts = Split(WF_FOLDER_CODES, " ")
    For i = 0 To UBound(parts)
        t = t & ChrW(CLng("&H" & parts(i)))
    Next i
    WfFolderName = t
End Function

Private Function WfResolvePath(ByVal fileName As String) As String
    Select Case LCase$(fileName)
        Case LCase$(GUIDANCE_FILE), LCase$(EXPORT_FILE), LCase$(RESULTS_FILE)
            WfResolvePath = ThisWorkbook.Path & Application.PathSeparator & WfFolderName() & Application.PathSeparator & fileName
        Case Else
            WfResolvePath = ThisWorkbook.Path & Application.PathSeparator & fileName
    End Select
End Function

' FileSystemObject is Unicode-safe; Dir is ANSI and can fail on Arabic folder names.
Private Function WfFileExists(ByVal p As String) As Boolean
    WfFileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
End Function

Private Function WorkflowFile01(ByVal fileName As String, Optional ByVal fullHash As Boolean = False) As String
    Dim path As String, stamp As String, h As Integer, isOpen As Boolean
    Dim size As Long, remaining As Long, block As Long, data() As Byte, j As Long
    Dim a As Long, b As Long, fso As Object, fil As Object
    On Error GoTo Missing
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    path = WfResolvePath(fileName)
    Set fso = CreateObject("Scripting.FileSystemObject")   ' Unicode-safe (the work folder name is Arabic)
    Set fil = fso.GetFile(path)
    size = fil.Size
    stamp = CStr(size) & "|" & Format$(fil.DateLastModified, "yyyymmddhhnnss")
    If Not fullHash Then
        WorkflowFile01 = stamp
        Exit Function
    End If
    ' Hash source bytes around collection: an unchanged old file cannot unlock transfer after cancellation.
    h = FreeFile
    Open path For Binary Access Read Shared As #h
    isOpen = True
    remaining = LOF(h): a = 1
    Do While remaining > 0
        block = remaining
        If block > 32768 Then block = 32768
        ReDim data(0 To block - 1)
        Get #h, , data
        For j = 0 To block - 1
            a = (a + CLng(data(j))) Mod 65521
            b = (b + a) Mod 65521
        Next j
        remaining = remaining - block
    Loop
    Close #h
    isOpen = False
    WorkflowFile01 = stamp & "|" & CStr(a) & ":" & CStr(b)
    Exit Function
Missing:
    On Error Resume Next
    If isOpen Then Close #h
    WorkflowFile01 = ""
End Function

Private Function TermWorkflowStage01() As Long
    Dim stage As Long, s As String
    s = WorkflowRead01("Stage")
    If IsNumeric(s) Then
        If Val(s) >= 0 And Val(s) <= 4 Then stage = CLng(s)
    End If
    If stage < 0 Or stage > 4 Then stage = 0
    If stage >= 1 Then
        s = WorkflowFile01("Term01Guidance.xlsx")
        If Len(s) = 0 Or s <> WorkflowRead01("Guidance") Then stage = 0
    End If
    If stage >= 2 Then
        s = WorkflowFile01("StreamPreferences01.xlsx")
        If Len(s) = 0 Or s <> WorkflowRead01("Wishes") Then stage = 1
    End If
    If stage >= 4 Then
        s = WorkflowFile01("OrientationResults01.xlsx")
        If Len(s) = 0 Or s <> WorkflowRead01("Export") Then stage = 3
    End If
    TermWorkflowStage01 = stage
End Function

Private Sub WorkflowSaveStage01(ByVal stage As Long)
    WorkflowWrite01 "Guidance", WorkflowFile01("Term01Guidance.xlsx")
    WorkflowWrite01 "Wishes", WorkflowFile01("StreamPreferences01.xlsx")
    WorkflowWrite01 "Export", WorkflowFile01("OrientationResults01.xlsx")
    WorkflowWrite01 "Stage", CStr(stage)
End Sub

Private Sub WorkflowGuard01(ByVal stepNo As Long)
    Dim previous As String
    Select Case stepNo
    Case 2: previous = "Data Transfer"
    Case 3: previous = "Stream Preferences"
    Case 4: previous = "First Term Orientation"
    End Select
    MsgBox "Complete '" & previous & "' successfully before this action.", vbExclamation, "Required previous step"
End Sub

Public Sub RunTermStep01()
    Dim caller As Variant, stepNo As Long, stage As Long, actions As Variant, item As Variant
    Dim i As Long, macro As String, resultFunction As String, succeeded As Boolean
    Dim errorText As String
    If m_workflowBusy Then Exit Sub
    caller = Application.Caller
    If VarType(caller) <> vbString Then
        MsgBox "Use the ordered buttons on ControlPanel.", vbExclamation, "Term 01 guidance"
        Exit Sub
    End If
    stepNo = WorkflowStep01(CStr(caller))
    If stepNo = 0 Then Exit Sub
    stage = TermWorkflowStage01()
    If stage < stepNo - 1 Then
        WorkflowGuard01 stepNo
        RefreshPanelState
        Exit Sub
    End If
    On Error GoTo Failed
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 810, , "The tool workbook is read-only."
    m_workflowBusy = True
    ' Retrying an earlier action invalidates every later completion, even if it fails or is cancelled.
    WorkflowSaveStage01 stepNo - 1
    RefreshPanelState
    actions = PanelActions()
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        If CStr(item(0)) = CStr(caller) Then
            macro = CStr(item(2))
            Exit For
        End If
    Next i
    If Len(macro) = 0 Then Err.Raise vbObjectError + 811, , "Panel action not found."
    Application.Run PanelMacroAction(macro)
    Select Case stepNo
    Case 1: resultFunction = "modTransfer01.Transfer01Succeeded"
    Case 2: resultFunction = "modStreamPreferences01.Preferences01Succeeded"
    Case 3: resultFunction = "modOrientation01.Orientation01Succeeded"
    Case 4: resultFunction = "modOrientationResults01.OrientationResults01Succeeded"
    End Select
    If Len(resultFunction) > 0 Then succeeded = CBool(Application.Run(PanelMacroAction(resultFunction)))
    If succeeded Then WorkflowSaveStage01 stepNo
Finished:
    m_workflowBusy = False
    On Error Resume Next
    RefreshPanelState
    On Error GoTo 0
    If Len(errorText) > 0 Then MsgBox errorText, vbExclamation, "Term 01 guidance"
    Exit Sub
Failed:
    errorText = Err.Description
    Resume Finished
End Sub

' ============================================================================
' Term 01 Analysis panel (sheet Term01Panel): file status table updated by Build_All
' (merged from the former module modBiuldAnalysisfilesPanel01)
' ============================================================================
Public Sub Panel_Create()
    Dim ws As Worksheet, map As Object, k As Variant, i As Long, n As Long
    Dim shp As Shape, rngBtn As Range, base As String

    On Error GoTo EH
    Application.ScreenUpdating = False
    EnsureFonts

    Set ws = PanelWs()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = ANALYSIS_PANEL_SHEET
    Else
        Do While ws.Shapes.Count > 0
            ws.Shapes(1).Delete
        Loop
        ws.Cells.UnMerge
        ws.Cells.Clear
    End If

    ws.DisplayRightToLeft = True
    With ws.Cells
        .Font.Name = gFontBase
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With

    ws.Columns(1).ColumnWidth = 2.5
    ws.Columns(2).ColumnWidth = 5
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 20
    ws.Columns(6).ColumnWidth = 62
    ws.Columns(7).ColumnWidth = 11
    ws.Columns(8).ColumnWidth = 2.5

    ' ---- banner: "   "
    ws.Rows(1).RowHeight = 10
    MergeText ws, 2, 2, 7, ArT("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644")
    With ws.Range("B2:G2")
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = gFontBase: .Font.Size = 28: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(2).RowHeight = 58

   ' Subtitle: "   (Data - Data01 - Analysis01 - AnalysisAnnuel - Cart_Tec)    "
    MergeText ws, 3, 2, 7, ArT("625 635 62F 627 631 20 645 644 641 627 62A 20 627 644 62A 62D 644 64A 644 20") & _
                           "(Data - Analysis01 - AnalyseCompare - Cart_Tec) " & _
                           ArT("644 643 644 20 627 644 634 639 628 20 628 636 63A 637 629 20 648 627 62D 62F 629")

       With ws.Range("B3:G3")
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = gFontName: .Font.Size = 14
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 12

    ' ---- information block
    base = ThisWorkbook.Path
    PutInfo ws, 5, ArT("645 62C 644 62F 20 627 644 645 635 62F 631"), base & "\Source01"
    PutInfo ws, 6, ArT("645 62C 644 62F 20 627 644 625 635 62F 627 631"), base & "\" & OutputFolderName()
    PutInfo ws, 7, ArT("622 62E 631 20 62A 634 63A 64A 644"), "-"
    PutInfo ws, 8, ArT("627 644 646 62A 64A 62C 629"), "-"
    ws.Rows(9).RowHeight = 14

    ' ---- buttons
    For i = 10 To 12
        ws.Rows(i).RowHeight = 26
    Next i
    ws.Rows(13).RowHeight = 14

    Set rngBtn = ws.Range("C10:D12")
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, rngBtn.Left + 2, rngBtn.Top + 2, rngBtn.Width - 4, rngBtn.Height - 4)
    shp.Name = BTN_MAIN
    StyleButton shp, "Term 01 Analysis", RGB(0, 112, 192), 24
    shp.OnAction = "'" & ThisWorkbook.Name & "'!modAnalysisBuilderBuildAll01.Build_All"

    Set rngBtn = ws.Range("E10:E12")
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, rngBtn.Left + 2, rngBtn.Top + 2, rngBtn.Width - 4, rngBtn.Height - 4)
    shp.Name = BTN_OPEN
    StyleButton shp, ArT("641 62A 62D 20 645 62C 644 62F 20 627 644 625 635 62F 627 631"), RGB(89, 89, 89), 14
    shp.OnAction = "'" & ThisWorkbook.Name & "'!modBuildPanel01.Panel_OpenOutput"

    ' ---- status table
    MergeText ws, 14, 2, 7, ArT("62D 627 644 629 20 627 644 645 644 641 627 62A")
    With ws.Range("B14:G14")
        .Font.Name = gFontBase: .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlRight
        .ReadingOrder = xlRTL
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(28, 55, 103)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(14).RowHeight = 30

    ws.Range("B15").Value2 = "#"
    ws.Range("C15").Value2 = ArT("627 644 645 644 641")
    ws.Range("D15").Value2 = ArT("627 644 623 642 633 627 645")
    ws.Range("E15").Value2 = ArT("627 644 62D 627 644 629")
    ws.Range("F15").Value2 = ArT("627 644 62A 641 627 635 64A 644")
    ws.Range("G15").Value2 = ArT("627 644 645 62F 629 20 28 62B 29")
    StyleHead ws.Range("B15:G15")
    ws.Rows(15).RowHeight = 28

    Set map = GetFileMap()
    n = 0
    For Each k In map.Keys
        n = n + 1
        With ws.Cells(ROW_FIRST + n - 1, 2)
            .Value2 = n
            .NumberFormat = "00"
        End With
        ws.Cells(ROW_FIRST + n - 1, 3).Value2 = CStr(k)
        ws.Cells(ROW_FIRST + n - 1, 4).Value2 = UBound(map(k)) - LBound(map(k)) + 1
        ws.Rows(ROW_FIRST + n - 1).RowHeight = 26
    Next k
    With ws.Range(ws.Cells(ROW_FIRST, 2), ws.Cells(ROW_FIRST + n - 1, 7))
        .Font.Name = gFontNum: .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(ROW_FIRST, 6), ws.Cells(ROW_FIRST + n - 1, 6)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(ROW_FIRST, 6), ws.Cells(ROW_FIRST + n - 1, 6)).WrapText = True
    ws.Range(ws.Cells(ROW_FIRST, 6), ws.Cells(ROW_FIRST + n - 1, 6)).Font.Size = 12
    For i = 1 To n
        Panel_Update i, 0, "", 0
    Next i

    ' ---- footer
    ws.Rows(ROW_FIRST + n).RowHeight = 12
    MergeText ws, ROW_FIRST + n + 1, 2, 7, ArT("644 627 20 64A 64F 633 62A 628 62F 644 20 623 64A 20 645 644 641 20 645 648 62C 648 62F 3A 20 639 646 62F 20 625 639 627 62F 629 20 627 644 62A 634 63A 64A 644 20 62A 64F 62D 641 638 20 627 644 645 644 641 627 62A 20 627 644 62C 62F 64A 62F 629 20 628 627 644 644 627 62D 642 629 20 5F 30 31 20 648 20 5F 30 32 20 2E 2E 2E 20 648 64A 64F 62D 641 638 20 633 62C 644 20 643 627 645 644 20 644 643 644 20 62A 634 63A 64A 644 20 641 64A 20 645 62C 644 62F 20 627 644 625 635 62F 627 631 2E")
    With ws.Range(ws.Cells(ROW_FIRST + n + 1, 2), ws.Cells(ROW_FIRST + n + 1, 7))
        .Font.Name = gFontName: .Font.Size = 12
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlRight
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
    ws.Rows(ROW_FIRST + n + 1).RowHeight = 34

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Panel_Create: " & Err.Number & " : " & Err.Description, vbCritical, "Term 01 Analysis"
End Sub

Private Sub PutInfo(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal txt As String)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Merge
    ws.Cells(r, 2).Value2 = lbl
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = gFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 7)).Merge
    ws.Cells(r, 4).NumberFormat = "@"
    ws.Cells(r, 4).Value2 = txt
    With ws.Range(ws.Cells(r, 4), ws.Cells(r, 7))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Name = gFontNum: .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With
    ws.Rows(r).RowHeight = 28
End Sub

Private Sub StyleButton(ByVal shp As Shape, ByVal txt As String, ByVal clr As Long, ByVal sz As Long)
    On Error Resume Next
    shp.Fill.ForeColor.RGB = clr
    shp.Line.ForeColor.RGB = RGB(28, 55, 103)
    shp.Line.Weight = 1.5
    shp.Adjustments.Item(1) = 0.25
    shp.Shadow.Visible = msoTrue
    shp.Shadow.OffsetX = 2
    shp.Shadow.OffsetY = 3
    shp.Shadow.Transparency = 0.65
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Name = gFontBase
        .TextRange.Font.Size = sz
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .HorizontalAnchor = msoAnchorCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
    End With
    Err.Clear
End Sub

Private Function PanelWs() As Worksheet
    On Error Resume Next
    Set PanelWs = ThisWorkbook.Worksheets(ANALYSIS_PANEL_SHEET)
    Err.Clear
End Function

Public Sub Panel_Update(ByVal idx As Long, ByVal kind As Long, ByVal detail As String, ByVal secs As Double)
    Dim ws As Worksheet, r As Long
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    EnsureFonts
    r = ROW_FIRST + idx - 1
    On Error Resume Next
    With ws.Cells(r, 5)
        Select Case kind
            Case 0
                .Value2 = ArT("641 64A 20 627 644 627 646 62A 638 627 631")
                .Interior.Color = RGB(242, 242, 242): .Font.Color = RGB(100, 100, 100)
            Case 1
                .Value2 = ArT("642 64A 62F 20 627 644 625 646 62C 627 632")
                .Interior.Color = RGB(255, 242, 204): .Font.Color = RGB(156, 101, 0)
            Case 2
                .Value2 = ArT("62A 645")
                .Interior.Color = RGB(198, 239, 206): .Font.Color = RGB(0, 97, 0)
            Case Else
                .Value2 = ArT("641 634 644")
                .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
        End Select
        .Font.Bold = True
        .Font.Name = gFontBase
    End With
    Dim board As Worksheet, boardRow As Long
    Set board = PanelSheet()
    If Not board Is Nothing Then
        boardRow = 13 + idx - 1
        Select Case kind
            Case 0: board.Cells(boardRow, 4).Value2 = "Pending": board.Cells(boardRow, 5).ClearContents
            Case 1: board.Cells(boardRow, 4).Value2 = "Running"
            Case 2: board.Cells(boardRow, 4).Value2 = "Done": board.Cells(boardRow, 5).Value2 = Round(secs, 1)
            Case Else: board.Cells(boardRow, 4).Value2 = "Failed": board.Cells(boardRow, 5).Value2 = Round(secs, 1)
        End Select
    End If
    ws.Cells(r, 6).NumberFormat = "@"
    ws.Cells(r, 6).Value2 = detail
    If kind >= 2 Then
        ws.Cells(r, 7).Value2 = Round(secs, 1)
        ws.Cells(r, 7).NumberFormat = "0.0"
    ElseIf kind = 0 Then
        ws.Cells(r, 7).ClearContents
    End If
    If kind = 3 Then
        ws.Rows(r).RowHeight = 54
    Else
        ws.Rows(r).RowHeight = 26
    End If
    Err.Clear
End Sub

Public Sub Panel_Reset(ByVal map As Object)
    Dim ws As Worksheet, i As Long
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To map.Count
        Panel_Update i, 0, "", 0
    Next i
    ws.Cells(7, 4).Value2 = Format$(Now, "yyyy-mm-dd  hh:nn")
    ws.Cells(8, 4).Value2 = ArT("642 64A 62F 20 627 644 62A 646 641 64A 630 20 2E 2E 2E")
    ws.Cells(8, 4).Font.Color = RGB(156, 101, 0)
    Err.Clear
End Sub

Public Sub Panel_Finish(ByVal okCount As Long, ByVal total As Long, ByVal secs As Double, ByVal note As String)
    Dim ws As Worksheet
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If note <> "" Then
        ws.Cells(8, 4).Value2 = note
        ws.Cells(8, 4).Font.Color = RGB(156, 0, 6)
    Else
        ws.Cells(8, 4).Value2 = ArT("62A 645 20 625 646 634 627 621 20") & okCount & ArT("20 20 645 646 20 20") & total & ArT("20 20 645 644 641 627 64B 20 20 2D 20 20") & Format$(secs, "0") & ArT("20 20 62B 627 646 64A 629")
        If okCount = total Then
            ws.Cells(8, 4).Font.Color = RGB(0, 97, 0)
        Else
            ws.Cells(8, 4).Font.Color = RGB(156, 0, 6)
        End If
    End If
    ws.Cells(8, 4).Font.Bold = True
    Err.Clear
End Sub

Public Sub Panel_Show()
    Dim ws As Worksheet
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
    ws.Range("A1").Select
    Err.Clear
End Sub

Public Sub Panel_OpenOutput()
    Dim p As String
    p = ThisWorkbook.Path & "\" & OutputFolderName()
    If Not EnsureFolderP(p) Then
        MsgBox "The output folder could not be opened or created:" & vbCrLf & p, vbInformation, "Output folder"
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=p, NewWindow:=True
    Err.Clear
End Sub
