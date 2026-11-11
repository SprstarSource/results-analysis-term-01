Attribute VB_Name = "modAnalysisBuilderCartTec01"
Option Explicit

'==============================================================================
' Module  : modAnalysisBuilderCartTec01   (BuildCartTecSheet)
' Purpose : Adds the Cart_Tec sheet (student technical card) of the term 1 file.
'   Layout (same as the Term 03 card):  previous-year averages [D] | [F] | term 1 [H]
'     [F] = exam results (BEPC / Baccalaureate) when the Data sheet has them,
'           otherwise the difference  term 1 - previous year.
'   The columns of the previous year / exam are found in Data by subject
'   (SubjKey), so the card follows whatever the model of the file contains.
'==============================================================================

Private Const FMT_MARK As String = "00.00;-00.00;;@"
Private Const FMT_DIFF As String = "+00.00;-00.00;00.00"
Private Const FMT_HIDE0 As String = "General;-General;;@"
Private Const FMT_DATE0 As String = "yyyy-mm-dd;;;@"
Private Const FMT_RANK0 As String = "00;-00;;@"

Public Function BuildCartTecSheet(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet, sp As Shape
    Dim n As Long, k As Long, r As Long, rLast As Long, rBottom As Long
    Dim multi As Boolean, maxV As Long, rS As Long
    Dim rowRef As String, yes As String, fmtExempt As String, annHdr As String
    Dim fcs As Object, stopTxt As String, stopSign As String, hideRng As String
    Dim clrYel As Long, clrGry As Long, clrCya As Long, clrBan As Long, clrRow As Long
    Dim rI As Range
    Dim cPrevAvg As Long, cExamAvg As Long, cAcc As Long, cMamen As Long, cPic As Long, cNotes As Long
    Dim cPrev() As Long, cExam() As Long, key As String
    Dim hasExam As Boolean, firstExam As Long, accIsTrans As Boolean
    Dim examTitle As String, c As Long

    BuildCartTecSheet = False
    errText = ""
    On Error GoTo EH

    If Not DiscoverData(wb, errText) Then Exit Function

    clrYel = RGB(255, 255, 102)
    clrGry = RGB(235, 235, 235)
    clrCya = RGB(102, 255, 255)
    clrBan = RGB(255, 255, 0)

    n = gSubjN
    multi = (gSecN > 1)
    yes = ArT("646 639 645")
    stopSign = ChrW(&HD83D) & ChrW(&HDED1)
    fmtExempt = "00.00;-00.00;" & QT & ArT("625 639 641 640 640 627 621") & QT & ";@"
    rowRef = ",$I$1,1)"
    rLast = 4 + n

    ' ---- columns of the model: previous year, exam, extras
    cPrevAvg = FindCol(ArT("627 644 645 639 62F 644 20 627 644 633 646 648 64A"))
    cExamAvg = FindCol(ArT("645 639 62F 644 20 627 644 634 647 627 62F 629"))
    If cExamAvg = 0 Then cExamAvg = FindCol(ArT("627 644 645 639 62F 644 20 627 644 639 627 645"))
    cAcc = FindCol(ArT("645 639 62F 644 20 627 644 627 646 62A 642 627 644"))
    accIsTrans = (cAcc > 0)
    If cAcc = 0 Then cAcc = FindCol(ArT("627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"))
    cMamen = FindCol(ArT("627 644 645 623 645 646"))
    cPic = FindCol(ArT("627 644 635 648 631 629"))
    cNotes = FindCol(ArT("645 644 627 62D 638 627 62A"))
    If cNotes = 0 Then cNotes = FindCol(ArT("627 644 645 644 627 62D 638 627 62A"))

    ReDim cPrev(1 To n)
    ReDim cExam(1 To n)
    For k = 1 To n
        key = SubjKey(gSubjName(k))
        cPrev(k) = FindSubjCol(key, "2")
        cExam(k) = FindSubjCol(key, "3")
        If cExam(k) > 0 Then
            If firstExam = 0 Or cExam(k) < firstExam Then firstExam = cExam(k)
        End If
    Next k
    If cExamAvg > 0 Then
        If firstExam = 0 Or cExamAvg < firstExam Then firstExam = cExamAvg
    End If
    hasExam = (firstExam > 0)
    If hasExam Then
        examTitle = Trim$(CStr(wb.Worksheets(DATA_SHEET).Cells(5, firstExam).Value))
        If examTitle = "" Then examTitle = ArT("646 62A 627 626 62C 20 627 644 627 645 62A 62D 627 646")
    End If

    rS = 12
    If cMamen > 0 Then rS = 14
    rBottom = rLast
    If rBottom < rS + 2 Then rBottom = rS + 2

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = CART_SHEET
    ws.DisplayRightToLeft = True

    ' ---- widths
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 17
    ws.Columns(3).ColumnWidth = 28
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 0.8
    ws.Columns(6).ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 0.8
    ws.Columns(8).ColumnWidth = 12
    ws.Columns(9).ColumnWidth = 7.1
    ws.Columns(10).ColumnWidth = 4.2
    ws.Columns(11).ColumnWidth = 5.6

    ' ---- record number (I1)
    ws.Range("I1").Value = 1
    With ws.Range("I1")
        .Font.Name = gFontNum: .Font.Size = 12: .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .NumberFormat = "0"
    End With

    ' ---- title
    MergeText ws, 1, 1, 8, ArT("628 637 627 642 629 20 641 646 64A 629 20 644 644 62A 644 645 64A 630")
    With ws.Range("A1:H1")
        .Interior.Color = clrBan
        .Font.Name = gFontBase: .Font.Size = 28: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(1).RowHeight = 44

    ' ---- left block : name, birth date, repeat, transition average, rank, origin school
    PutCardLabel ws, "A2", ArT("627 644 627 633 645 20 648 627 644 644 642 628"), clrYel
    PutCardValue ws, "A3", "=IFERROR(INDEX(" & ColRef(gHdr(2)) & rowRef & "," & QT & QT & ")", FMT_HIDE0
    PutCardLabel ws, "A4", ArT("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F"), clrYel
    PutCardValue ws, "A5", "=IFERROR(INDEX(" & ColRef(gHdr(3)) & rowRef & "," & QT & QT & ")", FMT_DATE0
    PutCardLabel ws, "A6", gHdr(5), clrYel
    PutCardValue ws, "A7", "=IFERROR(INDEX(" & ColRef(gHdr(5)) & rowRef & "," & QT & QT & ")", "General"
    If cAcc > 0 Then
        PutCardLabel ws, "A8", gHdr(cAcc), clrYel
        If accIsTrans Then
            ws.Range("A8").Formula = "=IF($A$7=" & QT & yes & QT & "," & QT & ArT("627 644 645 639 62F 644 20 627 644 633 646 648 64A") & QT & "," & QT & gHdr(cAcc) & QT & ")"
        End If
        PutCardValue ws, "A9", "=IFERROR(INDEX(" & ColRef(gHdr(cAcc)) & rowRef & "," & QT & QT & ")", FMT_MARK
    End If
    PutCardLabel ws, "A10", ArT("627 644 631 62A 628 629 20 641 64A 20 627 644 642 633 645"), clrYel
    PutCardValue ws, "A11", "=IFERROR(INDEX(" & ColRef(gHdr(1)) & rowRef & "," & QT & QT & ")", FMT_RANK0
    If cMamen > 0 Then
        PutCardLabel ws, "A12", ArT("627 644 645 624 633 633 629 20 627 644 623 635 644 64A 629 20 28 627 644 645 62A 648 633 637 629 29"), clrYel
        PutCardValue ws, "A13", "=IFERROR(INDEX(" & ColRef(gHdr(cMamen)) & rowRef & "," & QT & QT & ")", FMT_HIDE0
        ws.Range("A13").Font.Size = 14
    End If

    ' ---- picture and notes
    ws.Range("B2:B7").Merge
    If cPic > 0 Then
        If multi Then
            ws.Range("B2").Formula = "=IFERROR(IF($A$" & rS & "<>" & QT & QT & "," & QT & QT & ",INDEX(" & ColRef(gHdr(cPic)) & rowRef & ")," & QT & QT & ")"
        Else
            ws.Range("B2").Formula = "=IFERROR(INDEX(" & ColRef(gHdr(cPic)) & rowRef & "," & QT & QT & ")"
        End If
    End If
    With ws.Range("B2:B7")
        .NumberFormat = FMT_HIDE0
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    ws.Range("B8:B11").Merge
    If cNotes > 0 Then ws.Range("B8").Formula = "=IFERROR(INDEX(" & ColRef(gHdr(cNotes)) & rowRef & "," & QT & QT & ")"
    With ws.Range("B8:B11")
        .NumberFormat = FMT_HIDE0
        .Font.Name = gFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With

    ' ---- class title
    If multi Then
        ws.Range("C2").Value2 = gClassTitle
        With ws.Range("C2")
            .Interior.Color = clrYel
            .Font.Name = gFontBase: .Font.Size = 20: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .ShrinkToFit = True
            .ReadingOrder = xlRTL
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
        With ws.Range("C3")
            .NumberFormat = FMT_COUNT
            .Interior.Color = clrGry
            .Font.Name = gFontNum: .Font.Size = 22: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
    Else
        ws.Range("C2:C3").Merge
        ws.Range("C2").Value2 = gClassTitle
        With ws.Range("C2:C3")
            .Interior.Color = clrYel
            .Font.Name = gFontBase: .Font.Size = 20: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
            .ReadingOrder = xlRTL
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
    End If

    ' ---- results block : previous year (D) | exam or difference (F) | term 1 (H)
    ws.Range("D2:D3").Merge
    ws.Range("D2").Formula = "=IF($A$7=" & QT & yes & QT & "," & QT & ArT("645 639 62F 644 20 641 640 20 30 31 20 644 644 639 627 645 20 627 644 645 627 636 64A") & QT & "," & QT & ArT("627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629") & QT & ")"
    ws.Range("F2:F3").Merge
    If hasExam Then
        ws.Range("F2").Value2 = examTitle
    Else
        ws.Range("F2").Value2 = ArT("627 644 641 627 631 642")
    End If
    ws.Range("H2:H3").Merge
    ws.Range("H2").Value2 = ArT("646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644")
    StyleHead ws.Range("D2:D3")
    StyleHead ws.Range("F2:F3")
    StyleHead ws.Range("H2:H3")
    For Each rI In ws.Range("D2:D3,F2:F3,H2:H3").Areas
        rI.Interior.Color = clrCya
        rI.Font.Color = RGB(0, 0, 0)
        rI.Font.Size = 15
    Next rI
    If Not hasExam Then ws.Range("F2:F3").Font.Size = 18

    ' ---- averages (row 4)
    ws.Range("C4").Value2 = ArT("627 644 645 639 62F 644 20 627 644 639 627 645")
    If cPrevAvg > 0 Then ws.Range("D4").Formula = "=IFERROR(INDEX(" & ColRef(gHdr(cPrevAvg)) & rowRef & "," & QT & QT & ")"
    If hasExam Then
        If cExamAvg > 0 Then ws.Range("F4").Formula = "=IFERROR(INDEX(" & ColRef(gHdr(cExamAvg)) & rowRef & "," & QT & QT & ")"
    Else
        ws.Range("F4").Formula = "=IFERROR(IF(OR(H4=0,D4=0)," & QT & QT & ",H4-D4)," & QT & QT & ")"
    End If
    ws.Range("H4").Formula = "=IFERROR(INDEX(" & ColRef(gHdr(gAvgC1)) & rowRef & "," & QT & QT & ")"

    ' ---- subjects (rows 5..rLast)
    For k = 1 To n
        r = 4 + k
        ws.Cells(r, 3).Value2 = gSubjName(k)
        If cPrev(k) > 0 Then ws.Cells(r, 4).Formula = "=IFERROR(INDEX(" & ColRef(gHdr(cPrev(k))) & rowRef & "," & QT & QT & ")"
        If hasExam Then
            If cExam(k) > 0 Then ws.Cells(r, 6).Formula = "=IFERROR(INDEX(" & ColRef(gHdr(cExam(k))) & rowRef & "," & QT & QT & ")"
        Else
            ws.Cells(r, 6).Formula = "=IFERROR(IF(OR(H" & r & "=0,D" & r & "=0)," & QT & QT & ",H" & r & "-D" & r & ")," & QT & QT & ")"
        End If
        ws.Cells(r, 8).Formula = "=IFERROR(INDEX(" & ColRef(gHdr(gSubjC1(k))) & rowRef & "," & QT & QT & ")"
    Next k

    StyleGrid ws.Range(ws.Cells(4, 3), ws.Cells(rLast, 3))
    StyleGrid ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 4))
    StyleGrid ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6))
    StyleGrid ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8))
    With ws.Range(ws.Cells(4, 3), ws.Cells(rLast, 3))
        .Font.Name = gFontName: .Font.Size = 17: .Font.Bold = True
    End With
    ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 4)).Font.Size = 20
    ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8)).Font.Size = 20
    ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6)).Font.Size = 20
    ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6)).Font.Bold = True
    ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 4)).NumberFormat = FMT_MARK
    ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8)).NumberFormat = FMT_MARK
    If hasExam Then
        ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6)).NumberFormat = FMT_MARK
    Else
        ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6)).NumberFormat = FMT_DIFF
    End If
    If gPEIdx > 0 Then
        ws.Cells(4 + gPEIdx, 4).NumberFormat = fmtExempt
        ws.Cells(4 + gPEIdx, 8).NumberFormat = fmtExempt
        If hasExam Then ws.Cells(4 + gPEIdx, 6).NumberFormat = fmtExempt
    End If

    For Each rI In ws.Range("C4,D4,F4,H4").Areas
        rI.Interior.Color = clrCya
        rI.Font.Bold = True
    Next rI
    For k = 1 To n
        r = 4 + k
        If k Mod 2 = 1 Then clrRow = clrYel Else clrRow = clrGry
        ws.Cells(r, 3).Interior.Color = clrRow
        ws.Cells(r, 4).Interior.Color = clrRow
        ws.Cells(r, 6).Interior.Color = clrRow
        ws.Cells(r, 8).Interior.Color = clrRow
    Next k

    For r = 2 To rBottom
        ws.Rows(r).RowHeight = 32
    Next r

    ' ---- end-of-section warning (files that merge several sections)
    If multi Then
        stopTxt = stopSign & " " & ArT("646 640 647 640 627 64A 640 629 20 627 644 640 642 640 633 640 645") & " " & stopSign
        ws.Range("A" & rS & ":B" & (rS + 2)).Merge
        ws.Range("A" & rS).Formula = "=IFERROR(IF($C$3=" & QT & QT & "," & QT & QT & ",IF(INDEX(" & ColRef(gHdr(gSecCol)) & rowRef & _
                                  "<>$C$3," & QT & stopTxt & QT & "," & QT & QT & "))," & QT & QT & ")"
        With ws.Range("A" & rS & ":B" & (rS + 2))
            .Font.Name = gFontBase: .Font.Size = 26: .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = False
            .ShrinkToFit = True
            .ReadingOrder = xlRTL
        End With
    End If

    With ws.Range(ws.Cells(4, 6), ws.Cells(rLast, 6))
        Set fcs = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        fcs.Font.Color = RGB(0, 150, 50)
        Set fcs = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        fcs.Font.Color = RGB(226, 54, 0)
    End With
    If multi Then
        Set fcs = ws.Range("A" & rS & ":B" & (rS + 2)).FormatConditions.Add(Type:=xlExpression, Formula1:="=$A$" & rS & "<>" & QT & QT)
        fcs.Interior.Color = RGB(255, 255, 0)
        fcs.Font.Color = RGB(255, 0, 0)
        hideRng = "A3,A5,A7,B8:B11,D4:H" & rLast
        If cAcc > 0 Then hideRng = hideRng & ",A9"
        hideRng = hideRng & ",A11"
        If cMamen > 0 Then hideRng = hideRng & ",A13"
        Set fcs = ws.Range(hideRng).FormatConditions.Add(Type:=xlExpression, Formula1:="=$A$" & rS & "<>" & QT & QT)
        fcs.NumberFormat = ";;;"
        fcs.Font.Color = RGB(255, 255, 255)
    End If

    ' ---- note under the card (repeaters)
    If accIsTrans Then
        ws.Range(ws.Cells(rBottom + 2, 1), ws.Cells(rBottom + 3, 8)).Merge
        ws.Cells(rBottom + 2, 1).Value2 = ArT("627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 62A 644 645 64A 630 20 627 644 645 639 64A 62F 20 64A 643 62A 628 20 641 64A 20 62E 627 646 629 20 645 639 62F 644 20 627 644 627 646 62A 642 627 644 20 628 64A 646 645 627 20 627 644 645 639 62F 644 20 627 644 641 635 644 64A 20 644 644 641 635 644 20 30 31 20 64A 643 62A 628 20 641 64A 20 62E 627 646 629 20 627 644 645 639 62F 644 20 627 644 633 646 648 64A")
        With ws.Range(ws.Cells(rBottom + 2, 1), ws.Cells(rBottom + 3, 8))
            .Font.Name = gFontName: .Font.Size = 12
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
            .WrapText = True
            .ReadingOrder = xlRTL
        End With
    End If

    maxV = gRows
    If maxV < 1 Then maxV = 1
    If maxV > 30000 Then maxV = 30000
    Set sp = ws.Shapes.AddFormControl(xlSpinner, ws.Range("I1").Left + 2, ws.Range("I1").Top + 2, _
                                      ws.Range("I1").Width - 4, ws.Range("I1").Height + ws.Range("I2").Height - 4)
    sp.Name = "Spinner 1"
    With sp.ControlFormat
        .LinkedCell = "$I$1"
        .Min = 1
        .Max = maxV
        .SmallChange = 1
        .Value = 1
    End With

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    ws.Range("A1").Select
    Err.Clear
    On Error GoTo EH

    BuildCartTecSheet = True
    Exit Function
EH:
    errText = "BuildCartTecSheet: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function

Private Sub PutCardLabel(ByVal ws As Worksheet, ByVal addr As String, ByVal txt As String, ByVal clr As Long)
    ws.Range(addr).Value2 = txt
    With ws.Range(addr)
        .Interior.Color = clr
        .Font.Name = gFontName: .Font.Size = 17: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub PutCardValue(ByVal ws As Worksheet, ByVal addr As String, ByVal f As String, ByVal fmt As String)
    With ws.Range(addr)
        .Formula = f
        .NumberFormat = fmt
        .Font.Name = gFontName: .Font.Size = 17: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub
