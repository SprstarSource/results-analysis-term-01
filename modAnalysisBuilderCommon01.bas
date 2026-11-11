Attribute VB_Name = "modAnalysisBuilderCommon01"
Option Explicit

'==============================================================================
' Module  : modAnalysisBuilderCommon01
' Purpose : Shared constants, text helpers, fonts, cell styles and the
'           "discover what is inside the Data sheet" routine used by the Term 01
'           modules (Data01, Analysis01, Compare01, CartTec01).
'==============================================================================

' ---- sheet / table names
Public Const DATA_SHEET As String = "Data"
Public Const DATA_TABLE As String = "Data"
Public Const DATA_HDR_ROW As Long = 6
Public Const DATA_FIRST_ROW As Long = 7
Public Const ANALYSIS_SHEET As String = "Analysis01"
Public Const COMPARE_SHEET As String = "AnalyseCompare"
Public Const CART_SHEET As String = "Cart_Tec"

' ---- number formats
Public Const FMT_COUNT As String = "00"
Public Const FMT_DEC As String = "00.00"
Public Const FMT_PCT As String = "[=1]%0;%00.00"
Public Const PASS_MARK As Double = 10

' ---- layout
Public Const FIRST_NUM_COL As Long = 4
Public Const TITLE_ROW_HEIGHT As Double = 40
Public Const DATA_ROW_HEIGHT As Double = 24
Public Const HDR_ROW_HEIGHT As Double = 100
Public Const ID_COL_WIDTH As Double = 4.3
Public Const NUM_COL_WIDTH As Double = 6.3
Public Const DATE_COL_MIN_WIDTH As Double = 10.5

' ---- fonts
Public Const CHECK_FONTS As Boolean = True
Public Const FONT_BASE As String = "Dubai"
Public Const FONT_NAME As String = "Dubai Medium"
Public Const FONT_NUM As String = "Aptos Slab"
Public Const FONT_SAFE_TEXT As String = "Arial"
Public Const FONT_SAFE_NUM As String = "Calibri"

Public Const QT As String = """"

Public gFontBase As String
Public gFontName As String
Public gFontNum As String
Public gFontNotes As String

' ---- result of DiscoverData (read the Data table once, share with the builders)
Public gHdr() As String          ' table headers, 1..gNCols
Public gNCols As Long
Public gRows As Long
Public gSecCol As Long           ' column of the section header
Public gSubjN As Long            ' number of analysed subjects
Public gSubjName() As String     ' subject name without the term suffix
Public gSubjC1() As Long         ' column of the term 1 mark in Data
Public gAvgC1 As Long            ' column of the term 1 average in Data
Public gPEIdx As Long            ' index of the physical-education subject (0 = none)
Public gSecN As Long             ' number of sections that have students
Public gSecList() As Long        ' their numbers
Public gClassTitle As String     ' class title (Data!A5)

'==============================================================================
' TEXT HELPERS
'==============================================================================
Public Function ArT(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    ArT = s
End Function

Public Function NormKey(ByVal s As String) As String
    Dim c As Variant
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    For Each c In Array(&H200B, &H200E, &H200F, &H202A, &H202B, &H202C, &H640)
        s = Replace(s, ChrW(CLng(c)), "")
    Next c
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormKey = Trim$(s)
End Function

' Normalised Arabic used only for keyword matching (alef/ya forms)
Public Function NormAr(ByVal s As String) As String
    s = NormKey(s)
    s = Replace(s, ChrW(&H623), ChrW(&H627))
    s = Replace(s, ChrW(&H625), ChrW(&H627))
    s = Replace(s, ChrW(&H622), ChrW(&H627))
    s = Replace(s, ChrW(&H671), ChrW(&H627))
    s = Replace(s, ChrW(&H649), ChrW(&H64A))
    NormAr = s
End Function

' Structured reference to one column of the Data table (always double brackets)
Public Function ColRef(ByVal hdr As String) As String
    Dim t As String
    t = hdr
    t = Replace(t, "'", "''")
    t = Replace(t, "[", "'[")
    t = Replace(t, "]", "']")
    t = Replace(t, "#", "'#")
    ColRef = DATA_TABLE & "[[" & t & "]]"
End Function

Public Function ColLetter(ByVal c As Long) As String
    Dim a As String
    a = Cells(1, c).Address(False, False)
    ColLetter = Left$(a, Len(a) - 1)
End Function

' Replaces {key} by value; arguments come in pairs: key1, value1, key2, value2 ...
Public Function Subst(ByVal f As String, ParamArray kv() As Variant) As String
    Dim i As Long
    For i = LBound(kv) To UBound(kv) - 1 Step 2
        f = Replace(f, "{" & CStr(kv(i)) & "}", CStr(kv(i + 1)))
    Next i
    Subst = f
End Function

'==============================================================================
' FILES (FileSystemObject: safe with Arabic names)
'==============================================================================
Public Function FolderExistsP(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    FolderExistsP = fso.FolderExists(p)
    If Err.Number <> 0 Then FolderExistsP = False: Err.Clear
End Function

Public Function FileExistsP(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsP = fso.FileExists(p)
    If Err.Number <> 0 Then FileExistsP = False: Err.Clear
End Function

Public Function EnsureFolderP(ByVal p As String) As Boolean
    Dim fso As Object
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Not FolderExistsP(p) Then
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder p
        Err.Clear
        On Error GoTo 0
    End If
    EnsureFolderP = FolderExistsP(p)
End Function

Private Function IsOpenWb(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsOpenWb = True
            Exit Function
        End If
    Next wb
End Function

' stem.xlsx, and when it already exists: stem_01.xlsx, stem_02.xlsx ... (never overwrites)
Public Function NextFreePath(ByVal folder As String, ByVal stem As String, ByVal ext As String) As String
    Dim p As String, n As Long
    p = folder & stem & ext
    Do While FileExistsP(p) Or IsOpenWb(p)
        n = n + 1
        p = folder & stem & "_" & Format$(n, "00") & ext
    Loop
    NextFreePath = p
End Function

'==============================================================================
' FONTS
'==============================================================================
Public Sub ResolveFonts(ByVal ws As Worksheet)
    gFontNotes = ""
    gFontBase = PickFont(ws, FONT_BASE, FONT_SAFE_TEXT)
    gFontName = PickFont(ws, FONT_NAME, FONT_BASE, FONT_SAFE_TEXT)
    gFontNum = PickFont(ws, FONT_NUM, FONT_SAFE_NUM, FONT_SAFE_TEXT)
    ws.Cells.Clear
End Sub

' Resolves the fonts once per session (on a scratch workbook)
Public Sub EnsureFonts()
    Dim wbTmp As Workbook
    If gFontBase <> "" And gFontName <> "" And gFontNum <> "" Then Exit Sub
    Set wbTmp = Workbooks.Add(xlWBATWorksheet)
    ResolveFonts wbTmp.Worksheets(1)
    wbTmp.Close SaveChanges:=False
End Sub

' Name of the output folder (first term results)
Public Function OutputFolderName() As String
    OutputFolderName = ArT("645 639 627 644 62C 629 20 628 64A 627 646 627 62A 20 627 644 641 635 644 20 30 31")      ' the work folder (same name in every module)
End Function

Private Function PickFont(ByVal ws As Worksheet, ParamArray names() As Variant) As String
    Dim i As Long, first As String, chosen As String
    first = CStr(names(LBound(names)))
    If Not CHECK_FONTS Then
        PickFont = first
        Exit Function
    End If
    chosen = CStr(names(UBound(names)))
    For i = LBound(names) To UBound(names) - 1
        If FontAvailable(ws, CStr(names(i))) Then
            chosen = CStr(names(i))
            Exit For
        End If
    Next i
    PickFont = chosen
    If StrComp(chosen, first, vbTextCompare) <> 0 Then
        If InStr(gFontNotes, first & " ->") = 0 Then gFontNotes = gFontNotes & vbLf & first & " -> " & chosen
    End If
End Function

Private Function FontAvailable(ByVal ws As Worksheet, ByVal fontName As String) As Boolean
    Dim samples(1 To 2) As String, i As Long, wBase As Double, wTest As Double
    samples(1) = "MWmwIl0123456789 The quick brown fox " & ArT("627 644 639 631 628 64A 629 20 645 62F 631 633 629")
    samples(2) = "00.00 18.75 05.50 Notef11 MoyenneMe"
    On Error GoTo Unknown
    For i = 1 To 2
        wBase = MeasureWidth(ws, "Zz No Such Font Zz", samples(i))
        wTest = MeasureWidth(ws, fontName, samples(i))
        If Abs(wTest - wBase) > 0.3 Then
            FontAvailable = True
            Exit Function
        End If
    Next i
    Exit Function
Unknown:
    FontAvailable = True
End Function

Private Function MeasureWidth(ByVal ws As Worksheet, ByVal fontName As String, ByVal sample As String) As Double
    With ws.Range("A1")
        .Value2 = sample
        .Font.Name = fontName
        .Font.Size = 14
        .WrapText = False
    End With
    ws.Columns(1).AutoFit
    MeasureWidth = ws.Columns(1).ColumnWidth
End Function

'==============================================================================
' CELL STYLES
'==============================================================================
Public Sub StyleBanner(ByVal rng As Range)
    With rng
        .Font.Name = gFontBase: .Font.Size = 16: .Font.Bold = True
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
End Sub

Public Sub StyleTableTitle(ByVal rng As Range)
    With rng
        .Font.Name = gFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub StyleGrid(ByVal rng As Range)
    With rng
        .Font.Name = gFontNum: .Font.Size = 15: .Font.Bold = False
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        .Rows.RowHeight = DATA_ROW_HEIGHT
    End With
End Sub

Public Sub StyleHead(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = gFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub StyleLabel(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = gFontName: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub MergeText(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal txt As String)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        If c2 > c1 Then .Merge
        .Value2 = txt
    End With
End Sub

Public Sub PutLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    MergeText ws, r, 1, 3, txt
    StyleLabel ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
End Sub

'==============================================================================
' SUBJECT KEY : a short code for a subject, whatever the spelling / suffix is
'   (used by Cart_Tec to match the term 1 subject with its previous-year /
'    exam columns, whatever the spelling or the suffix 2 / 3 /  1)
'==============================================================================
Public Function SubjKey(ByVal s As String) As String
    Dim t As String
    t = NormAr(s)
    Do While Len(t) > 0
        If Right$(t, 1) Like "[0-9]" Then t = Left$(t, Len(t) - 1) Else Exit Do
    Loop
    t = Trim$(t)
    If Len(t) > 2 Then
        If Right$(t, 2) = " " & ChrW(&H641) Then t = Left$(t, Len(t) - 2)
    End If
    t = Replace(t, " ", "")
    t = Replace(t, ".", "")
    If InStr(t, ArT("628 62F 646")) > 0 Then
        SubjKey = "PE"
    ElseIf InStr(t, ArT("641 64A 632 64A 627")) > 0 Then
        SubjKey = "PHY"
    ElseIf InStr(t, ArT("62A 643 646 648 644 648 62C")) > 0 Then
        SubjKey = "TEC"
    ElseIf InStr(t, ArT("637 628 64A 639")) > 0 Then
        SubjKey = "NAT"
    ElseIf InStr(t, ArT("645 627 632 64A 63A")) > 0 Then
        SubjKey = "AMZ"
    ElseIf InStr(t, ArT("62C 646 628 64A")) > 0 Then
        SubjKey = "L3"
    ElseIf InStr(t, ArT("639 631 628 64A")) > 0 Then
        SubjKey = "ARB"
    ElseIf InStr(t, ArT("641 631 646 633")) > 0 Then
        SubjKey = "FRA"
    ElseIf InStr(t, ArT("627 646 62C 644 64A 632")) > 0 Then
        SubjKey = "ENG"
    ElseIf InStr(t, ArT("645 627 646 64A")) > 0 Then
        SubjKey = "GER"
    ElseIf InStr(t, ArT("633 628 627 646")) > 0 Then
        SubjKey = "SPA"
    ElseIf InStr(t, ArT("637 627 644 64A")) > 0 Then
        SubjKey = "ITA"
    ElseIf InStr(t, ArT("627 633 644 627 645")) > 0 Then
        SubjKey = "ISL"
    ElseIf InStr(t, ArT("62A 627 631 64A 62E")) > 0 Then
        SubjKey = "HIS"
    ElseIf InStr(t, ArT("631 64A 627 636 64A 627 62A")) > 0 Then
        SubjKey = "MAT"
    ElseIf InStr(t, ArT("645 639 644 648 645 627 62A")) > 0 Then
        SubjKey = "INF"
    ElseIf InStr(t, ArT("641 646 64A 629")) > 0 Then
        SubjKey = "ART"
    ElseIf InStr(t, ArT("641 644 633 641")) > 0 Then
        SubjKey = "PHI"
    ElseIf InStr(t, ArT("645 62D 627 633 628")) > 0 Then
        SubjKey = "ACC"
    ElseIf InStr(t, ArT("627 642 62A 635 627 62F")) > 0 Then
        SubjKey = "ECO"
    ElseIf InStr(t, ArT("642 627 646 648 646")) > 0 Then
        SubjKey = "LAW"
    Else
        SubjKey = t
    End If
End Function

' Column (1..gNCols) of the Data table whose header is exactly "hdr" (0 = not found)
Public Function FindCol(ByVal hdr As String) As Long
    Dim c As Long, k As String
    k = NormKey(hdr)
    For c = 1 To gNCols
        If StrComp(NormKey(gHdr(c)), k, vbTextCompare) = 0 Then
            FindCol = c
            Exit Function
        End If
    Next c
End Function

' Extra column (after the section column) of a subject: dg = "2" previous year, "3" exam
Public Function FindSubjCol(ByVal key As String, ByVal dg As String) As Long
    Dim c As Long, h As String
    For c = gSecCol + 1 To gNCols
        h = NormKey(gHdr(c))
        If Len(h) > 1 Then
            If Right$(h, 1) = dg Then
                If SubjKey(h) = key Then
                    FindSubjCol = c
                    Exit Function
                End If
            End If
        End If
    Next c
End Function

'==============================================================================
' DISCOVER the Data sheet (term 1 subjects, term 1 average, sections)
'==============================================================================
Private Function HasMarks1(ByVal lo As ListObject, ByVal c1 As Long) As Boolean
    Dim n As Double
    On Error Resume Next
    n = Application.WorksheetFunction.CountIf(lo.ListColumns(c1).DataBodyRange, ">0")
    HasMarks1 = (n > 0)
End Function

' Fills the g* variables. Subjects = columns F .. (term 1 average - 1) that contain
' at least one mark > 0. Returns False + errText on trouble.
Public Function DiscoverData(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    Dim c As Long, k As Long, mx As Long
    Dim h As String, b As String, sfx1 As String, mk As String
    Dim rngSec As Range

    DiscoverData = False: errText = ""
    On Error GoTo EH

    Set ws = wb.Worksheets(DATA_SHEET)
    Set lo = ws.ListObjects(DATA_TABLE)
    gNCols = lo.ListColumns.Count
    gRows = lo.ListRows.Count
    If gRows < 1 Then errText = "Data table is empty": Exit Function

    ReDim gHdr(1 To gNCols)
    gSecCol = 0
    For c = 1 To gNCols
        gHdr(c) = CStr(lo.ListColumns(c).Name)
        If gSecCol = 0 Then
            If StrComp(NormKey(gHdr(c)), ArT("627 644 642 633 645"), vbTextCompare) = 0 Then gSecCol = c
        End If
    Next c
    If gSecCol < 7 Then errText = "Section column not found in Data": Exit Function

    sfx1 = " " & ChrW(&H641) & " 1"
    mk = ArT("645 639 62F 644")

    gAvgC1 = 0
    For c = gSecCol - 1 To 6 Step -1
        If Left$(NormKey(gHdr(c)), 4) = mk Then gAvgC1 = c: Exit For
    Next c
    If gAvgC1 = 0 Then gAvgC1 = gSecCol - 1
    If gAvgC1 < 7 Then errText = "Term 1 average column not found in Data": Exit Function

    gSubjN = 0: gPEIdx = 0
    ReDim gSubjName(1 To gSecCol)
    ReDim gSubjC1(1 To gSecCol)

    For c = 6 To gAvgC1 - 1
        h = NormKey(gHdr(c))
        If Len(h) > 0 Then
            b = h
            If Len(h) > 4 Then
                If Right$(h, 4) = sfx1 Then b = Left$(h, Len(h) - 4)
            End If
            If HasMarks1(lo, c) Then
                gSubjN = gSubjN + 1
                gSubjName(gSubjN) = b
                gSubjC1(gSubjN) = c
                If gPEIdx = 0 Then
                    If InStr(NormAr(b), ArT("628 62F 646")) > 0 Then gPEIdx = gSubjN
                End If
            End If
        End If
    Next c

    If gSubjN < 1 Then errText = "No subject with marks was found in Data": Exit Function

    ' sections that really have students
    Set rngSec = lo.ListColumns(gSecCol).DataBodyRange
    mx = CLng(Application.WorksheetFunction.Max(rngSec))
    If mx < 1 Then errText = "Section numbers are missing in Data": Exit Function
    If mx > 60 Then errText = "Section number out of range": Exit Function
    gSecN = 0
    ReDim gSecList(1 To mx)
    For k = 1 To mx
        If Application.WorksheetFunction.CountIf(rngSec, k) > 0 Then
            gSecN = gSecN + 1
            gSecList(gSecN) = k
        End If
    Next k

    gClassTitle = CStr(ws.Range("A5").Value)
    DiscoverData = True
    Exit Function
EH:
    errText = "DiscoverData: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function
