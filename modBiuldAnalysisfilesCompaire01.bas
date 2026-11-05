Attribute VB_Name = "modBiuldAnalysisfilesCompaire01"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesCompaire01   (BuildCompaireSheet)
' Purpose : Adds the AnalyseCompaire sheet ONLY to the files that merge several
'           sections (1TecnoGlobal, 2SEGlobal, 3SEGlobal ...). It compares the
'           sections with the statistical indicators of the term 1 results.
'   - Table 1 : mean of every subject (and of the term 1 average) per section
'   - Table 2 : success rate of every subject per section
'   - Table 3 : general data (numbers, gender, repeaters, success rates)
'   - Then one block of indicators per subject (3 blocks per row):
'     success rate, mean, median, mode, max, min, range, standard deviation,
'     variance, coefficient of variation and the rank of each section.
'   Formulas read the Data table directly (same formulas / formats as Analysis01).
'==============================================================================

Private Const BLOCK_ROWS As Long = 14
Private Const IND_N As Long = 11

' ---- formula builders ---------------------------------------------------------
' secN = "" means "all sections"
Private Function CondP(ByVal sc As String, ByVal secN As String) As String
    If secN = "" Then CondP = "" Else CondP = sc & "," & secN & ","
End Function

Private Function MaskP(ByVal c As String, ByVal sc As String, ByVal secN As String) As String
    If secN = "" Then
        MaskP = "(" & c & ">0)"
    Else
        MaskP = "((" & sc & "=" & secN & ")*(" & c & ">0))"
    End If
End Function

Private Function FRate(ByVal c As String, ByVal sc As String, ByVal secN As String) As String
    FRate = "=IFERROR(COUNTIFS(" & CondP(sc, secN) & c & "," & QT & ">=" & CStr(PASS_MARK) & QT & ")/COUNTIFS(" & _
            CondP(sc, secN) & c & "," & QT & ">0" & QT & ")," & QT & QT & ")"
End Function

Private Function FMean(ByVal c As String, ByVal sc As String, ByVal secN As String) As String
    FMean = "=IFERROR(AVERAGEIFS(" & c & "," & CondP(sc, secN) & c & "," & QT & ">0" & QT & ")," & QT & QT & ")"
End Function

Private Function FMedian(ByVal c As String, ByVal sc As String, ByVal secN As String) As String
    FMedian = "=IFERROR(AGGREGATE(17,6," & c & "/" & MaskP(c, sc, secN) & ",2)," & QT & QT & ")"
End Function

Private Function FMode(ByVal c As String, ByVal sc As String, ByVal secN As String) As String
    FMode = "=IFERROR(MODE.SNGL(IF(" & MaskP(c, sc, secN) & "," & c & "))," & QT & QT & ")"
End Function

Private Function FMax(ByVal c As String, ByVal sc As String, ByVal secN As String) As String
    FMax = "=IFERROR(MAXIFS(" & c & "," & CondP(sc, secN) & c & "," & QT & ">0" & QT & ")," & QT & QT & ")"
End Function

Private Function FMin(ByVal c As String, ByVal sc As String, ByVal secN As String) As String
    FMin = "=IFERROR(MINIFS(" & c & "," & CondP(sc, secN) & c & "," & QT & ">0" & QT & ")," & QT & QT & ")"
End Function

Private Function FSd(ByVal c As String, ByVal sc As String, ByVal secN As String, ByVal mAddr As String) As String
    FSd = "=IFERROR(SQRT(SUMPRODUCT(" & MaskP(c, sc, secN) & "*(" & c & "-" & mAddr & ")^2)/COUNTIFS(" & _
          CondP(sc, secN) & c & "," & QT & ">0" & QT & "))," & QT & QT & ")"
End Function

' success rate of a group (gender / repeaters) on the term 1 average
Private Function FGroupRate(ByVal grp As String, ByVal val As String, ByVal avg As String, _
                            ByVal sc As String, ByVal secN As String) As String
    FGroupRate = "=IFERROR(COUNTIFS(" & CondP(sc, secN) & grp & "," & QT & val & QT & "," & avg & "," & QT & ">=" & CStr(PASS_MARK) & QT & ")/COUNTIFS(" & _
                 CondP(sc, secN) & grp & "," & QT & val & QT & "," & avg & "," & QT & ">0" & QT & ")," & QT & QT & ")"
End Function

Private Function FGroupCount(ByVal grp As String, ByVal val As String, ByVal sc As String, ByVal secN As String) As String
    FGroupCount = "=COUNTIFS(" & CondP(sc, secN) & grp & "," & QT & val & QT & ")"
End Function

' ---- small helpers --------------------------------------------------------------
Private Sub BestWorstRow(ByVal rng As Range)
    Dim fc As Object
    On Error Resume Next
    Set fc = rng.FormatConditions.AddTop10
    fc.TopBottom = xlTop10Top
    fc.Rank = 1
    fc.Percent = False
    fc.Interior.Color = RGB(198, 239, 206)
    fc.Font.Color = RGB(0, 97, 0)
    fc.Font.Bold = True
    Set fc = rng.FormatConditions.AddTop10
    fc.TopBottom = xlTop10Bottom
    fc.Rank = 1
    fc.Percent = False
    fc.Interior.Color = RGB(255, 199, 206)
    fc.Font.Color = RGB(156, 0, 6)
    fc.Font.Bold = True
    Err.Clear
End Sub

' Title row (merged over the block) + header row (label, section numbers, "all")
Private Sub PutTableHead(ByVal ws As Worksheet, ByVal r As Long, ByVal c0 As Long, ByVal nS As Long, _
                         ByVal ttl As String, ByVal lbl As String)
    Dim j As Long
    MergeText ws, r, c0, c0 + nS + 1, ttl
    StyleHead ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + nS + 1))
    ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + nS + 1)).Font.Size = 16
    ws.Rows(r).RowHeight = 36
    ws.Cells(r + 1, c0).Value2 = lbl
    For j = 1 To nS
        ws.Cells(r + 1, c0 + j).NumberFormat = FMT_COUNT
        ws.Cells(r + 1, c0 + j).Value2 = gSecList(j)
    Next j
    ws.Cells(r + 1, c0 + nS + 1).Value2 = ArT("627 644 643 644")
    StyleHead ws.Range(ws.Cells(r + 1, c0), ws.Cells(r + 1, c0 + nS + 1))
    ws.Rows(r + 1).RowHeight = 30
End Sub

Private Function SecText(ByVal j As Long, ByVal nS As Long) As String
    If j <= nS Then SecText = CStr(gSecList(j)) Else SecText = ""
End Function

'==============================================================================
' MAIN
'==============================================================================
Public Function BuildCompaireSheet(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet
    Dim n As Long, nS As Long, bw As Long, totalCols As Long
    Dim ref() As String, nm() As String
    Dim sc As String, sx As String, rp As String, avg As String
    Dim male As String, fem As String, yes As String
    Dim i As Long, j As Long, k As Long, b As Long, m As Long
    Dim c0 As Long, col As Long, rTop As Long, nTop As Long, rSub As Long, rGrid0 As Long
    Dim r0 As Long, r As Long, secN As String, rr(0 To IND_N - 1) As Long
    Dim mAddr As String, sdAddr As String, mxAddr As String, mnAddr As String, rngM As String
    Dim lbl As Variant, cLbl As Variant, f As String

    BuildCompaireSheet = False
    errText = ""
    On Error GoTo EH

    If Not DiscoverData(wb, errText) Then Exit Function
    If gSecN < 2 Then
        BuildCompaireSheet = True       ' single section: no comparison sheet
        Exit Function
    End If

    n = gSubjN
    nS = gSecN
    bw = nS + 2
    totalCols = 3 * bw + 2

    ReDim ref(1 To n + 1)
    ReDim nm(1 To n + 1)
    For i = 1 To n
        ref(i) = ColRef(gHdr(gSubjC1(i)))
        nm(i) = gSubjName(i)
    Next i
    ref(n + 1) = ColRef(gHdr(gAvgC1))
    nm(n + 1) = ArT("645 639 62F 644 20 627 644 641 635 644 20 627 644 623 648 644")
    avg = ref(n + 1)
    sc = ColRef(gHdr(gSecCol))
    sx = ColRef(gHdr(4))
    rp = ColRef(gHdr(5))
    male = ArT("630 643 631")
    fem = ArT("623 646 62B 649")
    yes = ArT("646 639 645")

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = COMPARE_SHEET
    ws.DisplayRightToLeft = True

    ' ---- widths
    For b = 0 To 2
        c0 = 1 + b * (bw + 1)
        ws.Columns(c0).ColumnWidth = 26
        ws.Range(ws.Columns(c0 + 1), ws.Columns(c0 + nS + 1)).ColumnWidth = 9.5
        If b < 2 Then ws.Columns(c0 + bw).ColumnWidth = 2.5
    Next b

    ' ---- banner
    MergeText ws, 1, 1, totalCols, ArT("645 642 627 631 646 629 20 628 64A 646 20 627 644 623 642 633 627 645 20 2D 20 62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644 20 2D 20") & gClassTitle
    StyleBanner ws.Range(ws.Cells(1, 1), ws.Cells(1, totalCols))
    ws.Rows(1).RowHeight = TITLE_ROW_HEIGHT
    ws.Rows(2).RowHeight = 6

    ' ---- three summary tables
    rTop = 3
    nTop = n + 1
    If nTop < 8 Then nTop = 8

    ' table 1 (mean) and table 2 (success rate)
    For b = 0 To 1
        c0 = 1 + b * (bw + 1)
        If b = 0 Then
            PutTableHead ws, rTop, c0, nS, ArT("645 642 627 631 646 629 20 645 62A 648 633 637 627 62A 20 627 644 645 648 627 62F 20 628 64A 646 20 627 644 623 642 633 627 645"), ArT("627 644 645 627 62F 629 20 2F 20 627 644 642 633 645")
        Else
            PutTableHead ws, rTop, c0, nS, ArT("645 642 627 631 646 629 20 646 633 628 20 627 644 646 62C 627 62D 20 641 64A 20 627 644 645 648 627 62F"), ArT("627 644 645 627 62F 629 20 2F 20 627 644 642 633 645")
        End If
        For i = 1 To n + 1
            r = rTop + 1 + i
            ws.Cells(r, c0).Value2 = nm(i)
            For j = 1 To nS + 1
                secN = SecText(j, nS)
                If b = 0 Then
                    ws.Cells(r, c0 + j).Formula = FMean(ref(i), sc, secN)
                Else
                    ws.Cells(r, c0 + j).Formula = FRate(ref(i), sc, secN)
                End If
            Next j
        Next i
        StyleGrid ws.Range(ws.Cells(rTop + 2, c0 + 1), ws.Cells(rTop + 1 + n + 1, c0 + nS + 1))
        If b = 0 Then
            ws.Range(ws.Cells(rTop + 2, c0 + 1), ws.Cells(rTop + 1 + n + 1, c0 + nS + 1)).NumberFormat = FMT_DEC
        Else
            ws.Range(ws.Cells(rTop + 2, c0 + 1), ws.Cells(rTop + 1 + n + 1, c0 + nS + 1)).NumberFormat = FMT_PCT
        End If
        StyleLabel ws.Range(ws.Cells(rTop + 2, c0), ws.Cells(rTop + 1 + n + 1, c0))
        For i = 1 To n + 1
            BestWorstRow ws.Range(ws.Cells(rTop + 1 + i, c0 + 1), ws.Cells(rTop + 1 + i, c0 + nS))
        Next i
        ' the average row is highlighted
        With ws.Range(ws.Cells(rTop + 1 + n + 1, c0), ws.Cells(rTop + 1 + n + 1, c0 + nS + 1))
            .Font.Bold = True
            .Borders(xlEdgeTop).Weight = xlMedium
        End With
    Next b

    ' table 3 : general data
    c0 = 1 + 2 * (bw + 1)
    PutTableHead ws, rTop, c0, nS, ArT("645 639 637 64A 627 62A 20 639 627 645 629"), ArT("627 644 628 64A 627 646 20 2F 20 627 644 642 633 645")
    cLbl = Array(ArT("62A 639 62F 627 62F 20 627 644 642 633 645"), ArT("639 62F 62F 20 627 644 630 643 648 631"), ArT("639 62F 62F 20 627 644 625 646 627 62B"), _
                 ArT("639 62F 62F 20 627 644 645 639 64A 62F 64A 646"), ArT("639 62F 62F 20 627 644 646 627 62C 62D 64A 646"), _
                 ArT("646 633 628 629 20 646 62C 627 62D 20 627 644 630 643 648 631"), ArT("646 633 628 629 20 646 62C 627 62D 20 627 644 625 646 627 62B"), _
                 ArT("646 633 628 629 20 646 62C 627 62D 20 627 644 645 639 64A 62F 64A 646"))
    For i = 0 To 7
        r = rTop + 2 + i
        ws.Cells(r, c0).Value2 = cLbl(i)
        For j = 1 To nS + 1
            secN = SecText(j, nS)
            Select Case i
                Case 0
                    If secN = "" Then
                        f = "=COUNTIFS(" & sc & "," & QT & ">0" & QT & ")"
                    Else
                        f = "=COUNTIFS(" & sc & "," & secN & ")"
                    End If
                Case 1: f = FGroupCount(sx, male, sc, secN)
                Case 2: f = FGroupCount(sx, fem, sc, secN)
                Case 3: f = FGroupCount(rp, yes, sc, secN)
                Case 4: f = "=COUNTIFS(" & CondP(sc, secN) & avg & "," & QT & ">=" & CStr(PASS_MARK) & QT & ")"
                Case 5: f = FGroupRate(sx, male, avg, sc, secN)
                Case 6: f = FGroupRate(sx, fem, avg, sc, secN)
                Case 7: f = FGroupRate(rp, yes, avg, sc, secN)
            End Select
            ws.Cells(r, c0 + j).Formula = f
        Next j
    Next i
    StyleGrid ws.Range(ws.Cells(rTop + 2, c0 + 1), ws.Cells(rTop + 9, c0 + nS + 1))
    ws.Range(ws.Cells(rTop + 2, c0 + 1), ws.Cells(rTop + 6, c0 + nS + 1)).NumberFormat = FMT_COUNT
    ws.Range(ws.Cells(rTop + 7, c0 + 1), ws.Cells(rTop + 9, c0 + nS + 1)).NumberFormat = FMT_PCT
    StyleLabel ws.Range(ws.Cells(rTop + 2, c0), ws.Cells(rTop + 9, c0))
    ApplyPassColors ws.Range(ws.Cells(rTop + 7, c0 + 1), ws.Cells(rTop + 9, c0 + nS + 1)), "=0.5", "=1"

    ' ---- sub banner
    ws.Rows(rTop + 2 + nTop).RowHeight = 8
    rSub = rTop + 3 + nTop
    MergeText ws, rSub, 1, totalCols, ArT("62A 641 635 64A 644 20 627 644 645 624 634 631 627 62A 20 627 644 625 62D 635 627 626 64A 629 20 644 643 644 20 645 627 62F 629")
    StyleBanner ws.Range(ws.Cells(rSub, 1), ws.Cells(rSub, totalCols))
    ws.Rows(rSub).RowHeight = TITLE_ROW_HEIGHT
    ws.Rows(rSub + 1).RowHeight = 6
    rGrid0 = rSub + 2

    ' ---- one block per subject (+ the term 1 average)
    lbl = Array(ArT("646 633 628 629 20 627 644 646 62C 627 62D"), ArT("627 644 645 62A 648 633 637 20 627 644 62D 633 627 628 64A"), _
                ArT("627 644 648 633 64A 637"), ArT("627 644 645 646 648 627 644"), _
                ArT("623 639 644 649 20 645 639 62F 644"), ArT("623 62F 646 649 20 645 639 62F 644"), ArT("627 644 645 62F 649"), _
                ArT("627 644 627 646 62D 631 627 641 20 627 644 645 639 64A 627 631 64A"), ArT("627 644 62A 628 627 64A 646"), _
                ArT("645 639 627 645 644 20 627 644 627 62E 62A 644 627 641"), ArT("62A 631 62A 64A 628 20 627 644 642 633 645 20 28 627 644 645 62A 648 633 637 29"))

    For m = 0 To n
        r0 = rGrid0 + (m \ 3) * BLOCK_ROWS
        c0 = 1 + (m Mod 3) * (bw + 1)
        If m < n Then
            PutTableHead ws, r0, c0, nS, ArT("645 627 62F 629 20") & nm(m + 1), ArT("627 644 642 633 645")
        Else
            PutTableHead ws, r0, c0, nS, nm(m + 1), ArT("627 644 642 633 645")
        End If
        For k = 0 To IND_N - 1
            rr(k) = r0 + 2 + k
            ws.Cells(rr(k), c0).Value2 = lbl(k)
        Next k
        rngM = ws.Range(ws.Cells(rr(1), c0 + 1), ws.Cells(rr(1), c0 + nS)).Address(True, True)
        For j = 1 To nS + 1
            col = c0 + j
            secN = SecText(j, nS)
            mAddr = ws.Cells(rr(1), col).Address(False, False)
            mxAddr = ws.Cells(rr(4), col).Address(False, False)
            mnAddr = ws.Cells(rr(5), col).Address(False, False)
            sdAddr = ws.Cells(rr(7), col).Address(False, False)
            ws.Cells(rr(0), col).Formula = FRate(ref(m + 1), sc, secN)
            ws.Cells(rr(1), col).Formula = FMean(ref(m + 1), sc, secN)
            ws.Cells(rr(2), col).Formula = FMedian(ref(m + 1), sc, secN)
            PutModeFormula ws.Cells(rr(3), col), FMode(ref(m + 1), sc, secN)
            ws.Cells(rr(4), col).Formula = FMax(ref(m + 1), sc, secN)
            ws.Cells(rr(5), col).Formula = FMin(ref(m + 1), sc, secN)
            ws.Cells(rr(6), col).Formula = "=IFERROR(" & mxAddr & "-" & mnAddr & "," & QT & QT & ")"
            ws.Cells(rr(7), col).Formula = FSd(ref(m + 1), sc, secN, mAddr)
            ws.Cells(rr(8), col).Formula = "=IFERROR(" & sdAddr & "^2," & QT & QT & ")"
            ws.Cells(rr(9), col).Formula = "=IFERROR(" & sdAddr & "/" & mAddr & "," & QT & QT & ")"
            If j <= nS Then
                ws.Cells(rr(10), col).Formula = "=IFERROR(RANK(" & mAddr & "," & rngM & ")," & QT & QT & ")"
            End If
        Next j
        StyleGrid ws.Range(ws.Cells(rr(0), c0 + 1), ws.Cells(rr(IND_N - 1), c0 + nS + 1))
        StyleLabel ws.Range(ws.Cells(rr(0), c0), ws.Cells(rr(IND_N - 1), c0))
        ws.Range(ws.Cells(rr(0), c0 + 1), ws.Cells(rr(0), c0 + nS + 1)).NumberFormat = FMT_PCT
        ws.Range(ws.Cells(rr(1), c0 + 1), ws.Cells(rr(8), c0 + nS + 1)).NumberFormat = FMT_DEC
        ws.Range(ws.Cells(rr(9), c0 + 1), ws.Cells(rr(9), c0 + nS + 1)).NumberFormat = FMT_PCT
        ws.Range(ws.Cells(rr(10), c0 + 1), ws.Cells(rr(10), c0 + nS + 1)).NumberFormat = FMT_COUNT
        ApplyPassColors ws.Range(ws.Cells(rr(0), c0 + 1), ws.Cells(rr(0), c0 + nS + 1)), "=0.5", "=1"
        ApplyPassColors ws.Range(ws.Cells(rr(1), c0 + 1), ws.Cells(rr(1), c0 + nS + 1)), "=10", "=20"
        ws.Rows(r0 + 2 + IND_N).RowHeight = 6
    Next m

    ' ---- print / view
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(rGrid0 + ((n \ 3) + 1) * BLOCK_ROWS - 2, totalCols)).Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    ws.Range("A1").Select
    Err.Clear
    On Error GoTo EH

    BuildCompaireSheet = True
    Exit Function
EH:
    errText = "BuildCompaireSheet: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function
