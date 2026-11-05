Attribute VB_Name = "modOrientationResults01"
Option Explicit

Private m_runSucceeded As Boolean

' ============================================================================
' Module: modOrientationResults01
' Purpose:
'   Export the final guidance results from Term01Guidance.xlsx to a separate
'   workbook named OrientationResults01.xlsx.
'
' Output:
'   - Sheet "Litr": all arts students from Data01.
'   - One sheet per science section from Data, named Tecno1, Tecno2, ...
'   - The title row holds only the report title (arts) or the report title plus
'     the two-digit section number (science); the sheet name is never shown.
'   - The science section column (N) is removed from every output sheet.
'   - Science students are sorted by Term 1 average (column M), descending.
'   - The first column in each science sheet is replaced by the rank number.
'
' The source workbook is opened read-only and is never modified.
' Place this module in the tool workbook. The source and output files are
' expected beside the tool workbook.
'
' Entry point: Run_OrientationResults01
' ============================================================================

Private Const SOURCE_FILE As String = "Term01Guidance.xlsx"
Private Const OUTPUT_FILE As String = "OrientationResults01.xlsx"

Private Const SCIENCE_SHEET As String = "Data"
Private Const ARTS_SHEET As String = "Data01"
Private Const ARTS_OUTPUT_SHEET As String = "Litr"

Private Const FIRST_DATA_ROW As Long = 3
Private Const HEADER_ROW As Long = 2

Private Const SCIENCE_WIDTH As Long = 21
Private Const ARTS_WIDTH As Long = 15

Private Const SCIENCE_SECTION_COL As Long = 14   ' N
Private Const SCIENCE_AVERAGE_COL As Long = 13   ' M

Private Const FONT_ARABIC As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK As String = "Calibri"
Private Const TITLE_FONT As String = "Cairo"
Private Const TITLE_FONT_FALLBACK As String = "Arial"
Private Const SCORE_COLUMN_WIDTH As Double = 5.7
Private Const BODY_FONT_SIZE As Long = 14
Private Const TITLE_FONT_SIZE As Long = 11

' ---------------------------------------------------------------------------
' Public entry point
' ---------------------------------------------------------------------------
Public Sub Run_OrientationResults01()

    Dim wbSource As Workbook
    Dim wbOutput As Workbook
    Dim wsScience As Worksheet
    Dim wsArts As Worksheet
    Dim loScience As ListObject
    Dim loArts As ListObject

    Dim sourceOpened As Boolean
    Dim outputCreated As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim oldEvents As Boolean
    Dim oldCalc As XlCalculation
    Dim stateSaved As Boolean

    Dim outputPath As String
    Dim artsCount As Long
    Dim scienceCount As Long
    Dim sectionCount As Long
    Dim skippedScience As Long
    Dim reportText As String
    Dim answer As VbMsgBoxResult
    Dim failed As Boolean
    Dim errText As String

    m_runSucceeded = False
    On Error GoTo EH

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 1000, , _
                  "Save the tool workbook before running Orientation Results."
    End If

    outputPath = ThisWorkbook.Path & Application.PathSeparator & OUTPUT_FILE

    If IsWorkbookOpenAtPath(outputPath) Then
        Err.Raise vbObjectError + 1001, , _
                  OUTPUT_FILE & " is already open. Close it and run again."
    End If

    If Len(Dir$(outputPath)) > 0 Then
        answer = MsgBox(OUTPUT_FILE & " already exists." & vbCrLf & _
                        "Do you want to replace it with a new results file?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, _
                        "Orientation Results 01")
        If answer <> vbYes Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    stateSaved = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wbSource = OpenSourceWorkbook(SOURCE_FILE, sourceOpened)

    Set wsScience = FindSheet(wbSource, SCIENCE_SHEET)
    Set wsArts = FindSheet(wbSource, ARTS_SHEET)

    If wsScience Is Nothing And wsArts Is Nothing Then
        Err.Raise vbObjectError + 1002, , _
                  "Neither Data nor Data01 was found in " & SOURCE_FILE & "."
    End If

    If Not wsScience Is Nothing Then
        Set loScience = FindGuidanceTable(wsScience, SCIENCE_WIDTH)
    End If

    If Not wsArts Is Nothing Then
        Set loArts = FindGuidanceTable(wsArts, ARTS_WIDTH)
    End If

    Set wbOutput = Workbooks.Add(xlWBATWorksheet)
    outputCreated = True

    ' Arts are written first to the sheet named "01".
    If Not wsArts Is Nothing Then
        artsCount = WriteArtsResults(wbSource, wsArts, loArts, wbOutput)
    End If

    ' Science is split into one sheet per section.
    If Not wsScience Is Nothing Then
        WriteScienceResults wbSource, wsScience, loScience, wbOutput, _
                            scienceCount, sectionCount, skippedScience
    End If

    If wbOutput.Worksheets.Count = 1 Then
        If wbOutput.Worksheets(1).UsedRange.Cells.Count = 1 _
           And Len(CStr(wbOutput.Worksheets(1).Range("A1").Value2)) = 0 Then
            If artsCount = 0 And sectionCount = 0 Then
                Err.Raise vbObjectError + 1003, , _
                          "No student rows were found in Data or Data01."
            End If
        End If
    End If

    wbOutput.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing
    outputCreated = False
    m_runSucceeded = True

    reportText = "Orientation results were exported successfully." & vbCrLf & vbCrLf & _
                 "Arts students (sheet Litr): " & CStr(artsCount) & vbCrLf & _
                 "Science students: " & CStr(scienceCount) & vbCrLf & _
                 "Science section sheets: " & CStr(sectionCount)

    If skippedScience > 0 Then
        reportText = reportText & vbCrLf & _
                     "Science rows without a valid section: " & CStr(skippedScience)
    End If

    reportText = reportText & vbCrLf & vbCrLf & outputPath

CleanExit:
    On Error Resume Next

    If Not wbOutput Is Nothing Then
        If failed Then
            wbOutput.Close SaveChanges:=False
        End If
    End If

    If sourceOpened And Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If

    If stateSaved Then
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
    End If

    On Error GoTo 0

    If failed Then
        MsgBox "Orientation Results failed:" & vbCrLf & errText, _
               vbCritical, "Orientation Results 01"
    Else
        MsgBox reportText, vbInformation, "Orientation Results 01"
    End If

    Exit Sub

EH:
    m_runSucceeded = False
    failed = True
    errText = Err.Description
    Resume CleanExit

End Sub

' ---------------------------------------------------------------------------
' Arts output: one sheet named "Litr", all arts students, original source order.
' Data01 columns:
' A:O = number, identity, term average, two wishes, two profiles, decision.
' ---------------------------------------------------------------------------
Private Function WriteArtsResults(ByVal wbSource As Workbook, _
                                  ByVal wsSource As Worksheet, _
                                  ByVal loSource As ListObject, _
                                  ByVal wbOutput As Workbook) As Long

    Dim wsOutput As Worksheet
    Dim sourceData As Variant
    Dim headers As Variant
    Dim rowCount As Long

    headers = loSource.HeaderRowRange.Value2

    If loSource.DataBodyRange Is Nothing Then
        rowCount = 0
    Else
        sourceData = loSource.DataBodyRange.Value2
        rowCount = UBound(sourceData, 1)
    End If

    Set wsOutput = GetAvailableOutputSheet(wbOutput, ARTS_OUTPUT_SHEET)
    wsOutput.Name = ARTS_OUTPUT_SHEET

    WriteOutputSheet wsSource, wsOutput, headers, sourceData, _
                     rowCount, ARTS_WIDTH, _
                     CStr(wsSource.Range("A1").Value2), False

    WriteArtsResults = rowCount

End Function

' ---------------------------------------------------------------------------
' Science output: one worksheet per section (Tecno1, Tecno2, ...).
' Data columns:
' M = Term 1 average, N = section, O:P = wishes, Q:T = profiles, U = decision.
' ---------------------------------------------------------------------------
Private Sub WriteScienceResults(ByVal wbSource As Workbook, _
                                ByVal wsSource As Worksheet, _
                                ByVal loSource As ListObject, _
                                ByVal wbOutput As Workbook, _
                                ByRef totalStudents As Long, _
                                ByRef sectionSheets As Long, _
                                ByRef skippedRows As Long)

    Dim sourceData As Variant
    Dim headers As Variant
    Dim groups As Object
    Dim rowList As Collection
    Dim keys As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim sectionNo As Long
    Dim key As String
    Dim wsOutput As Worksheet
    Dim groupIndex As Long
    Dim sectionName As String

    If loSource.DataBodyRange Is Nothing Then Exit Sub

    sourceData = loSource.DataBodyRange.Value2
    rowCount = UBound(sourceData, 1)
    headers = loSource.HeaderRowRange.Value2

    Set groups = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        If TryGetSectionNumber(sourceData(i, SCIENCE_SECTION_COL), sectionNo) Then
            key = CStr(sectionNo)

            If Not groups.Exists(key) Then
                Set rowList = New Collection
                groups.Add key, rowList
            End If

            Set rowList = groups(key)
            rowList.Add i
        Else
            skippedRows = skippedRows + 1
        End If
    Next i

    If groups.Count = 0 Then Exit Sub

    keys = groups.Keys
    SortNumericKeys keys

    For groupIndex = LBound(keys) To UBound(keys)

        sectionNo = CLng(keys(groupIndex))
        sectionName = "Tecno" & CStr(sectionNo)

        Set rowList = groups(CStr(sectionNo))
        Set wsOutput = GetAvailableOutputSheet(wbOutput, sectionName)

        WriteScienceGroup wsSource, wsOutput, headers, sourceData, _
                          rowList, ScienceTitle(sectionNo)

        totalStudents = totalStudents + rowList.Count
        sectionSheets = sectionSheets + 1

    Next groupIndex

End Sub

' ---------------------------------------------------------------------------
' Build one science section and sort by Term 1 average descending.
' Ties retain the original order from the source table.
' ---------------------------------------------------------------------------
Private Sub WriteScienceGroup(ByVal wsSource As Worksheet, _
                              ByVal wsOutput As Worksheet, _
                              ByVal headers As Variant, _
                              ByVal sourceData As Variant, _
                              ByVal rowList As Collection, _
                              ByVal titleText As String)

    Dim order() As Long
    Dim outputData() As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim c As Long
    Dim sourceRow As Long

    rowCount = rowList.Count
    ReDim order(1 To rowCount)

    For i = 1 To rowCount
        order(i) = CLng(rowList(i))
    Next i

    StableSortRowsByAverage order, sourceData

    ReDim outputData(1 To rowCount, 1 To SCIENCE_WIDTH)

    For i = 1 To rowCount
        sourceRow = order(i)

        For c = 1 To SCIENCE_WIDTH
            outputData(i, c) = sourceData(sourceRow, c)
        Next c

        ' The first column is the rank within this section.
        outputData(i, 1) = i
    Next i

    WriteOutputSheet wsSource, wsOutput, headers, outputData, _
                     rowCount, SCIENCE_WIDTH, _
                     titleText, True

End Sub

' ---------------------------------------------------------------------------
' Write data and reproduce the collector-style layout:
' title row, rotated/filterable header row, centered body, RTL, identity
' columns AutoFit, score columns fixed width, and a plain Excel table.
' ---------------------------------------------------------------------------
Private Sub WriteOutputSheet(ByVal wsSource As Worksheet, _
                             ByVal wsOutput As Worksheet, _
                             ByVal headers As Variant, _
                             ByVal data As Variant, _
                             ByVal rowCount As Long, _
                             ByVal columnCount As Long, _
                             ByVal titleText As String, _
                             ByVal isScience As Boolean)

    Dim lastRow As Long
    Dim tbl As ListObject
    Dim targetRange As Range
    Dim sourceTitle As Range
    Dim sourceHeader As Range
    Dim sourceBodyRow As Range
    Dim targetTitle As Range
    Dim c As Long
    Dim arabicFont As String
    Dim numericFont As String
    Dim titleFont As String

    wsOutput.Cells.Clear
    wsOutput.DisplayRightToLeft = True

    lastRow = HEADER_ROW + rowCount

    ' Title row: copy the source title's visual style where available.
    Set targetTitle = wsOutput.Range(wsOutput.Cells(1, 1), _
                                     wsOutput.Cells(1, columnCount))
    targetTitle.Merge
    wsOutput.Cells(1, 1).Value2 = titleText

    On Error Resume Next
    Set sourceTitle = wsSource.Range(wsSource.Cells(1, 1), _
                                     wsSource.Cells(1, columnCount))
    sourceTitle.Copy
    targetTitle.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0

    titleFont = AvailableFont(TITLE_FONT, TITLE_FONT_FALLBACK)
    With targetTitle
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Font.Name = titleFont
        .Font.Size = TITLE_FONT_SIZE
        .Font.Bold = True
    End With
    wsOutput.Rows(1).RowHeight = 30

    ' Header values and formatting come directly from the guidance table.
    wsOutput.Cells(HEADER_ROW, 1).Resize(1, columnCount).Value2 = headers

    Set sourceHeader = wsSource.Range(wsSource.Cells(HEADER_ROW, 1), _
                                      wsSource.Cells(HEADER_ROW, columnCount))
    sourceHeader.Copy
    wsOutput.Cells(HEADER_ROW, 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Data rows are copied as values only.
    If rowCount > 0 Then
        wsOutput.Cells(FIRST_DATA_ROW, 1).Resize(rowCount, columnCount).Value2 = data

        Set sourceBodyRow = wsSource.Range(wsSource.Cells(FIRST_DATA_ROW, 1), _
                                           wsSource.Cells(FIRST_DATA_ROW, columnCount))
        sourceBodyRow.Copy
        wsOutput.Cells(FIRST_DATA_ROW, 1).Resize(rowCount, columnCount).PasteSpecial _
            Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' Keep the source number formats for every output column.
        For c = 1 To columnCount
            wsOutput.Columns(c).NumberFormat = sourceBodyRow.Cells(1, c).NumberFormat
        Next c
    End If

    Set targetRange = wsOutput.Range(wsOutput.Cells(HEADER_ROW, 1), _
                                     wsOutput.Cells(lastRow, columnCount))

    Set tbl = wsOutput.ListObjects.Add(SourceType:=xlSrcRange, _
                                       Source:=targetRange, _
                                       XlListObjectHasHeaders:=xlYes)

    On Error Resume Next
    tbl.TableStyle = ""
    tbl.ShowTableStyleColumnStripes = False
    tbl.ShowTableStyleRowStripes = False
    tbl.ShowTableStyleFirstColumn = False
    tbl.ShowTableStyleLastColumn = False
    tbl.ShowAutoFilter = True
    On Error GoTo 0

    ' Collector-style typography and alignment.
    arabicFont = AvailableFont(FONT_ARABIC, FONT_FALLBACK)
    numericFont = AvailableFont(FONT_NUMERIC, FONT_FALLBACK)

    With wsOutput
        .Range(.Cells(1, 1), .Cells(lastRow, columnCount)).HorizontalAlignment = xlCenter
        .Range(.Cells(1, 1), .Cells(lastRow, columnCount)).VerticalAlignment = xlCenter
        .Range(.Cells(HEADER_ROW, 1), .Cells(lastRow, columnCount)).WrapText = False

        .Columns("A:E").AutoFit
        If columnCount >= 6 Then
            .Range(.Columns(6), .Columns(columnCount)).ColumnWidth = SCORE_COLUMN_WIDTH
        End If

        .Columns("B").Font.Name = arabicFont
        .Columns("B").Font.Size = BODY_FONT_SIZE

        .Columns("C").Font.Name = numericFont
        .Columns("C").Font.Size = BODY_FONT_SIZE

        If columnCount >= 4 Then
            .Range(.Columns(4), .Columns(columnCount)).Font.Name = arabicFont
            .Range(.Columns(4), .Columns(columnCount)).Font.Size = BODY_FONT_SIZE
        End If

        If columnCount >= 6 Then
            .Range(.Columns(6), .Columns(columnCount)).Font.Name = numericFont
            .Range(.Columns(6), .Columns(columnCount)).Font.Size = BODY_FONT_SIZE
        End If

        If columnCount >= 4 Then
            .Range(.Cells(HEADER_ROW, 4), .Cells(HEADER_ROW, columnCount)).Orientation = xlUpward
            .Range(.Cells(HEADER_ROW, 4), .Cells(HEADER_ROW, columnCount)).Font.Bold = True
        End If

        .Rows(HEADER_ROW).RowHeight = 120
        If rowCount > 0 Then
            .Range(.Rows(FIRST_DATA_ROW), .Rows(lastRow)).RowHeight = 20
        End If
    End With

    If isScience And rowCount > 0 Then
        wsOutput.Columns(1).NumberFormat = "0"
    End If

    ' The section is already encoded in the title and sheet name: drop column N.
    If isScience Then wsOutput.Columns(SCIENCE_SECTION_COL).Delete

End Sub

' ---------------------------------------------------------------------------
' Title helpers (Arabic text is built from code points to keep the file ASCII)
' ---------------------------------------------------------------------------
Private Function ScienceTitle(ByVal sectionNo As Long) As String
    ScienceTitle = Txt("645 62D 636 631 20 627 644 62A 648 62C 64A 647 20 627 644 62A 62F 631 64A 62C 64A 20 " & _
                       "62D 633 628 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644 20 " & _
                       "644 644 633 646 629 20 627 644 623 648 644 649 20 62C 630 639 20 645 634 62A 631 643 20 " & _
                       "639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627") & _
                   " " & Format$(sectionNo, "00")
End Function

Private Function Txt(ByVal codes As String) As String
    Dim parts() As String, i As Long, result As String
    parts = Split(codes, " ")
    For i = 0 To UBound(parts)
        result = result & ChrW(CLng("&H" & parts(i)))
    Next i
    Txt = result
End Function

' ---------------------------------------------------------------------------
' Workbook and table helpers
' ---------------------------------------------------------------------------
Private Function OpenSourceWorkbook(ByVal fileName As String, _
                                    ByRef openedHere As Boolean) As Workbook

    Dim wb As Workbook
    Dim fullPath As String

    openedHere = False
    fullPath = ThisWorkbook.Path & Application.PathSeparator & fileName

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenSourceWorkbook = wb
            Exit Function
        End If

        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1010, , _
                      "A different workbook with the same name is already open: " & wb.FullName
        End If
    Next wb

    If Len(Dir$(fullPath)) = 0 Then
        Err.Raise vbObjectError + 1011, , "Source file not found: " & fullPath
    End If

    Set OpenSourceWorkbook = Workbooks.Open( _
        Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    openedHere = True

End Function

Private Function FindSheet(ByVal wb As Workbook, _
                           ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FindSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindGuidanceTable(ByVal ws As Worksheet, _
                                   ByVal requiredWidth As Long) As ListObject

    Dim lo As ListObject

    For Each lo In ws.ListObjects
        If lo.Range.Row = HEADER_ROW And lo.Range.Column = 1 Then
            If lo.ListColumns.Count < requiredWidth Then
                Err.Raise vbObjectError + 1020, , _
                          "Table on sheet '" & ws.Name & "' has fewer than " & _
                          CStr(requiredWidth) & " columns."
            End If

            Set FindGuidanceTable = lo
            Exit Function
        End If
    Next lo

    Err.Raise vbObjectError + 1021, , _
              "A guidance table starting at A2 was not found on sheet '" & ws.Name & "'."

End Function

Private Function GetAvailableOutputSheet(ByVal wb As Workbook, _
                                         ByVal desiredName As String) As Worksheet

    Dim ws As Worksheet

    If wb.Worksheets.Count = 1 Then
        Set ws = wb.Worksheets(1)
        If ws.UsedRange.Cells.Count = 1 And Len(CStr(ws.Range("A1").Value2)) = 0 _
           And ws.ListObjects.Count = 0 Then
            ws.Name = desiredName
            Set GetAvailableOutputSheet = ws
            Exit Function
        End If
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = desiredName
    Set GetAvailableOutputSheet = ws

End Function

Private Function IsWorkbookOpenAtPath(ByVal fullPath As String) As Boolean

    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsWorkbookOpenAtPath = True
            Exit Function
        End If
    Next wb

End Function

' ---------------------------------------------------------------------------
' Section grouping and stable sorting
' ---------------------------------------------------------------------------
Private Function TryGetSectionNumber(ByVal value As Variant, _
                                     ByRef sectionNo As Long) As Boolean

    Dim s As String

    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function

    s = Trim$(CStr(value))
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    If CDbl(s) < 1 Or CDbl(s) <> Fix(CDbl(s)) Then Exit Function

    sectionNo = CLng(s)
    TryGetSectionNumber = True

End Function

Private Sub SortNumericKeys(ByRef keys As Variant)

    Dim i As Long
    Dim j As Long
    Dim temp As Variant

    If Not IsArray(keys) Then Exit Sub

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(j)) < CLng(keys(i)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i

End Sub

Private Sub StableSortRowsByAverage(ByRef order() As Long, _
                                    ByVal sourceData As Variant)

    Dim i As Long
    Dim j As Long
    Dim currentRow As Long
    Dim currentAverage As Double

    For i = LBound(order) + 1 To UBound(order)

        currentRow = order(i)
        currentAverage = NumericValue(sourceData(currentRow, SCIENCE_AVERAGE_COL))
        j = i - 1

        Do While j >= LBound(order)
            If NumericValue(sourceData(order(j), SCIENCE_AVERAGE_COL)) >= currentAverage Then
                Exit Do
            End If

            order(j + 1) = order(j)
            j = j - 1
        Loop

        order(j + 1) = currentRow

    Next i

End Sub

Private Function NumericValue(ByVal value As Variant) As Double

    If IsError(value) Or IsEmpty(value) Then
        NumericValue = -1E+99
    ElseIf IsNumeric(value) Then
        NumericValue = CDbl(value)
    Else
        NumericValue = -1E+99
    End If

End Function

' ---------------------------------------------------------------------------
' Font fallback
' ---------------------------------------------------------------------------
Private Function AvailableFont(ByVal preferred As String, _
                               ByVal fallback As String) As String

    ' Excel substitutes unavailable fonts automatically. Keep the same
    ' preferred/fallback choices used by the collector module.
    AvailableFont = preferred

End Function

Public Function OrientationResults01Succeeded() As Boolean
    OrientationResults01Succeeded = m_runSucceeded
End Function
