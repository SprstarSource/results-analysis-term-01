Attribute VB_Name = "modBuildPanel"
Option Explicit

'==============================================================================
' Module  : modBuildPanel
' Purpose : builds the ControlPanel sheet. Every button is wired DIRECTLY to the
'           entry macro of one module (no steps, no locks, no state):
'
'   First Term Results   -> modCollector.CollectTerm01Results
'   Term 01 Analysis     -> modBiuldAnalysisfilesBuildAll01.Build_All
'   Examen Analysis      -> modExamenAnalysis01.Run_ExamenAnalysis01
'   Term 01 Orientation  -> modOrientationAll01.Run_OrientationAll01
'                           (runs the whole orientation block in one go)
'
' Run BuildControlPanel once (or again after adding / renaming a module).
' Panel_Reset / Panel_Update / Panel_Finish / Panel_Show (at the end of this module) draw the
'           progress table of Term 01 Analysis under the buttons.
' RefreshPanelState is kept because modCollector calls it after a collection;
' it simply re-links the buttons (useful after "Save As" / renaming the file).
'==============================================================================

Private Const PANEL_SHEET As String = "ControlPanel"
Private Const PANEL_TITLE As String = "First Term Management Dashboard"
Private Const BUTTON_WIDTH As Double = 164
Private Const BUTTON_HEIGHT As Double = 54
Private Const BUTTON_GAP As Double = 12
Private Const BUTTON_COLUMNS As Long = 2

' First row of the progress table (set by Panel_Reset)
Private m_top As Long

' Each action: Array(shape name, caption, macro (Module.Procedure), colour)
Private Function PanelActions() As Variant
    PanelActions = Array( _
        Array("btnCollect", "First Term Results", "modCollector.CollectTerm01Results", &H60AE27), _
        Array("btnTerm01Analysis", "Term 01 Analysis", "modBiuldAnalysisfilesBuildAll01.Build_All", &HB98029), _
        Array("btnExamen", "Examen Analysis", "modExamenAnalysis01.Run_ExamenAnalysis01", &HA05568), _
        Array("btnOrientation", "Term 01 Orientation", "modOrientationAll01.Run_OrientationAll01", &HB98029))
End Function

Public Sub BuildControlPanel()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, count As Long, gridRow As Long, gridCol As Long
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

    ' Start from a clean sheet: every old button (including obsolete ones) is removed.
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
            .Value = PANEL_TITLE
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
    canvasWidth = BUTTON_COLUMNS * BUTTON_WIDTH + (BUTTON_COLUMNS - 1) * BUTTON_GAP
    originLeft = ws.Range("B3").Left + (ws.Range("B1:K1").Width - canvasWidth) / 2
    originTop = ws.Range("B3").Top
    For i = 0 To count - 1
        item = actions(LBound(actions) + i)
        gridRow = i \ BUTTON_COLUMNS
        gridCol = i Mod BUTTON_COLUMNS
        rowItems = count - gridRow * BUTTON_COLUMNS
        If rowItems > BUTTON_COLUMNS Then rowItems = BUTTON_COLUMNS
        rowWidth = rowItems * BUTTON_WIDTH + (rowItems - 1) * BUTTON_GAP
        x = originLeft + (canvasWidth - rowWidth) / 2 + gridCol * (BUTTON_WIDTH + BUTTON_GAP)
        y = originTop + gridRow * (BUTTON_HEIGHT + BUTTON_GAP)
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, BUTTON_WIDTH, BUTTON_HEIGHT)
        With btn
            .Name = CStr(item(0))
            .Adjustments(1) = 0.12
            .Placement = xlFreeFloating
            .Fill.Solid
            .Fill.ForeColor.RGB = CLng(item(3))
            .Line.Weight = 1.25
            .Line.ForeColor.RGB = CLng(item(3))
            .OnAction = PanelMacroAction(CStr(item(2)))
            With .TextFrame2
                .TextRange.Text = CStr(item(1))
                .TextRange.Font.Name = "Segoe UI"
                .TextRange.Font.Size = 12
                .TextRange.Font.Bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
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
    CheckPanelMacros actions
    Application.ScreenUpdating = prevScreen
    MsgBox "Control panel updated: " & count & " buttons.", vbInformation, "Done"
    Exit Sub
BuildError:
    errorText = CStr(Err.Number) & " : " & Err.Description
    Application.ScreenUpdating = prevScreen
    MsgBox "Panel build failed: " & errorText, vbExclamation, "Error"
End Sub

' Re-links every button to its macro (the workbook name is part of the link, so this
' is needed after "Save As" / renaming). Called by modCollector after a collection.
Public Sub RefreshPanelState()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant, i As Long
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    actions = PanelActions()
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        Set btn = Nothing
        On Error Resume Next
        Set btn = ws.Shapes(CStr(item(0)))
        Err.Clear
        On Error GoTo 0
        If Not btn Is Nothing Then btn.OnAction = PanelMacroAction(CStr(item(2)))
    Next i
End Sub

Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

Private Function PanelMacroAction(ByVal macroName As String) As String
    PanelMacroAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

' Warns about any module that has not been imported yet.
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

'==============================================================================
'  Progress table of "Term 01 Analysis" (called by modBiuldAnalysisfilesBuildAll01)
'  It is written on the ControlPanel sheet, under the buttons:
'    Panel_Reset   lists every file to build (status "Waiting")
'    Panel_Update  idx = position of the file, state 1 = running, 2 = done, 3 = failed
'    Panel_Finish  one summary line
'    Panel_Show    brings the panel sheet to the front
'  Every procedure is silent when the sheet is missing or protected.
'==============================================================================

Private Function ProgressTop(ByVal ws As Worksheet) As Long
    Dim shp As Shape, r As Long, maxRow As Long
    maxRow = 2
    On Error Resume Next
    For Each shp In ws.Shapes
        r = shp.BottomRightCell.Row
        If r > maxRow Then maxRow = r
    Next shp
    On Error GoTo 0
    ProgressTop = maxRow + 2
End Function

Public Sub Panel_Reset(ByVal map As Object)
    Dim ws As Worksheet, key As Variant, r As Long, n As Long
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    m_top = ProgressTop(ws)
    With ws.Range(ws.Rows(m_top), ws.Rows(m_top + 120)).Columns("B:K")
        .UnMerge
        .Clear
    End With
    With ws.Range(ws.Cells(m_top, 2), ws.Cells(m_top, 11))
        .Merge
        .Value = "Term 01 Analysis - progress"
        .Font.Name = "Segoe UI"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(m_top).RowHeight = 26
    PanelSummary ws, "Running ...", False
    ' column headers
    ws.Cells(m_top + 2, 2).Value = "#"
    ws.Range(ws.Cells(m_top + 2, 3), ws.Cells(m_top + 2, 5)).Merge
    ws.Cells(m_top + 2, 3).Value = "File"
    ws.Cells(m_top + 2, 6).Value = "State"
    ws.Cells(m_top + 2, 7).Value = "Sec."
    ws.Range(ws.Cells(m_top + 2, 8), ws.Cells(m_top + 2, 11)).Merge
    ws.Cells(m_top + 2, 8).Value = "Result"
    With ws.Range(ws.Cells(m_top + 2, 2), ws.Cells(m_top + 2, 11))
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Interior.Color = RGB(221, 228, 240)
        .HorizontalAlignment = xlCenter
    End With
    ' one row per file
    If Not map Is Nothing Then
        For Each key In map.Keys
            n = n + 1
            r = m_top + 2 + n
            ws.Cells(r, 2).Value = n
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).Merge
            ws.Cells(r, 3).Value = CStr(key)
            ws.Cells(r, 6).Value = "Waiting"
            ws.Range(ws.Cells(r, 8), ws.Cells(r, 11)).Merge
            PanelRowStyle ws, r, RGB(245, 245, 245), RGB(110, 110, 110)
        Next key
    End If
    On Error GoTo 0
End Sub

Public Sub Panel_Update(ByVal idx As Long, ByVal state As Long, ByVal info As String, ByVal secs As Double)
    Dim ws As Worksheet, r As Long, fillC As Long, fontC As Long, caption As String
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    If m_top = 0 Then Exit Sub
    On Error Resume Next
    r = m_top + 2 + idx
    Select Case state
        Case 1: caption = "Running": fillC = RGB(221, 235, 247): fontC = RGB(31, 78, 121)
        Case 2: caption = "Done":    fillC = RGB(226, 239, 218): fontC = RGB(56, 87, 35)
        Case Else: caption = "Failed": fillC = RGB(252, 228, 214): fontC = RGB(156, 0, 6)
    End Select
    ws.Cells(r, 6).Value = caption
    If state <> 1 Then ws.Cells(r, 7).Value = Format$(secs, "0.0")
    If Len(info) > 0 Or state <> 1 Then ws.Cells(r, 8).Value = Left$(info, 140)
    PanelRowStyle ws, r, fillC, fontC
    ws.Cells(r, 8).HorizontalAlignment = xlLeft
    On Error GoTo 0
End Sub

Public Sub Panel_Finish(ByVal okCount As Long, ByVal total As Long, ByVal secs As Double, ByVal note As String)
    Dim ws As Worksheet, txt As String
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If m_top = 0 Then m_top = ProgressTop(ws)
    If Len(note) > 0 Then
        txt = Left$(note, 200)
        PanelSummary ws, txt, True
    Else
        txt = "Created " & okCount & " / " & total & "   (" & Format$(secs, "0") & " s)"
        PanelSummary ws, txt, (okCount < total)
    End If
    On Error GoTo 0
End Sub

Public Sub Panel_Show()
    Dim ws As Worksheet
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Sub PanelSummary(ByVal ws As Worksheet, ByVal txt As String, ByVal isProblem As Boolean)
    With ws.Range(ws.Cells(m_top + 1, 2), ws.Cells(m_top + 1, 11))
        .UnMerge
        .Merge
        .Value = txt
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        If isProblem Then
            .Interior.Color = RGB(252, 228, 214)
            .Font.Color = RGB(156, 0, 6)
        Else
            .Interior.Color = RGB(226, 239, 218)
            .Font.Color = RGB(56, 87, 35)
        End If
    End With
    ws.Rows(m_top + 1).RowHeight = 22
End Sub

Private Sub PanelRowStyle(ByVal ws As Worksheet, ByVal r As Long, ByVal fillC As Long, ByVal fontC As Long)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 11))
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = fontC
        .Interior.Color = fillC
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
End Sub
