Attribute VB_Name = "modBuildPanel"
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

Private Function PanelActions() As Variant
    PanelActions = Array( _
        Array("btnCollect", "First Term Results", "CollectTerm01Results", False, &H60AE27), _
        Array("btnTerm01Analysis", "Term 01 Analysis", "modBiuldAnalysisfilesBuildAll01.Build_All", False, &HB98029), _
        Array("btnExamen", "Examen Analysis", "modExamenAnalysis01.Run_ExamenAnalysis01", False, &HA05568), _
        Array("btnStreamPreferences", "Stream Preferences", "modStreamPreferences01.Run_StreamPreferences01", False, &HA05568), _
        Array("btnDataTransfer", "Data Transfer", "modTransfer01.Run_Transfer01", False, &H60AE27), _
        Array("btnFirstTermOrientation", "First Term Orientation", "modOrientation01.Run_Orientation01", False, &HB98029), _
        Array("btnExportOrientationResults", "Export Orientation Results", "modOrientationResults01.Run_OrientationResults01", False, &H60AE27))
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
        .Columns("B:K").ColumnWidth = 9.6
        .Rows(1).RowHeight = 48
        .Rows(2).RowHeight = 14
        With .Range("B1:K1")
            .Merge
            .Value = "First Term Management Dashboard"
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
    columns = 3
    canvasWidth = columns * BUTTON_WIDTH + (columns - 1) * BUTTON_GAP
    originLeft = ws.Range("B3").Left + (ws.Range("B1:K1").Width - canvasWidth) / 2
    originTop = ws.Range("B3").Top
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
    Case "btnFirstTermOrientation": WorkflowStep01 = 3
    Case "btnExportOrientationResults": WorkflowStep01 = 4
    End Select
End Function

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
