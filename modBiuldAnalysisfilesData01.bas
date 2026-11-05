Attribute VB_Name = "modBiuldAnalysisfilesData01"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesData01
' Purpose : Reads \Source01\, builds the Data sheet (term 1 marks + the extra
'           columns of the model of each file: previous-year averages, exam results,
'           notes, photo ... filled in by hand after the files are created).
'==============================================================================

Private Const SOURCE_FOLDER_NAME As String = "Source01"

' ---- Collection engine constants ----
Private Const DATA_START_ROW       As Long = 6
Private Const TRAILING_ROWS_SKIP   As Long = 1
Private Const IDENTITY_LAST_COL    As Long = 5
Private Const COL_WIDTH            As Double = 5.7
Private Const FONT_SIZE            As Long = 14
Private Const FONT_ARABIC          As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC         As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK        As String = "Calibri"

Private Const PREFERRED_FONT As String = "Cairo"
Private Const FALLBACK_FONT  As String = "Arial"
Private Const HEADER_FONT_SIZE As Double = 11
Private Const HEADER_ROW_HEIGHT As Double = 30

Private g_ActiveArabicFont  As String
Private g_ActiveNumericFont As String

' ==============================================================================
' HELPERS
' ==============================================================================
Private Function AR(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    AR = s
End Function

Private Function NormalizeKey(ByVal s As String) As String
    Dim c As Variant
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    For Each c In Array(&H200B, &H200E, &H200F, &H202A, &H202B, &H202C, &H640)
        s = Replace(s, ChrW(CLng(c)), "")
    Next c
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeKey = Trim(s)
End Function

Private Function CellText(ByVal r As Range) As String
    If IsError(r.Value) Then CellText = "" Else CellText = CStr(r.Value)
End Function

Private Function FileExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FileExistsPath = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExistsPath = False: Err.Clear
End Function

Private Function FolderExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExistsPath = ((GetAttr(p) And vbDirectory) <> 0)
    If Err.Number <> 0 Then FolderExistsPath = False: Err.Clear
End Function

Private Function EnsureFolder(ByVal p As String) As Boolean
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    If Not FolderExistsPath(p) Then
        On Error Resume Next
        MkDir p
        Err.Clear
        On Error GoTo 0
    End If
    EnsureFolder = FolderExistsPath(p)
End Function

Private Function IsWorkbookOpen(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If f Is Nothing Then LastUsedRow = 0 Else LastUsedRow = f.Row
End Function

Private Function SortedNames(ByVal col As Collection) As Variant
    Dim arr() As String, i As Long, j As Long, tmp As String
    If col.Count = 0 Then SortedNames = Array(): Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count: arr(i) = CStr(col(i)): Next i
    For i = 2 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 1
            If StrComp(arr(j), tmp, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    SortedNames = arr
End Function

' ==============================================================================
' COLLECTION OF THE SOURCE FILES INTO ONE WORKBOOK
' ==============================================================================
Private Function FontExists(ByVal FontName As String) As Boolean
    Const HKCU As Long = &H80000001
    Const HKLM As Long = &H80000002
    Dim reg As Object, hv As Variant
    Dim names As Variant, types As Variant
    Dim i As Long, p As Long, nm As String, enumOK As Boolean

    On Error GoTo Fallback
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    For Each hv In Array(HKLM, HKCU)
        names = Empty
        If reg.EnumValues(CLng(hv), "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", names, types) = 0 Then
            enumOK = True
            If IsArray(names) Then
                For i = LBound(names) To UBound(names)
                    nm = CStr(names(i))
                    p = InStr(nm, " (")
                    If p > 0 Then nm = Left(nm, p - 1)
                    If StrComp(nm, FontName, vbTextCompare) = 0 Then FontExists = True: Exit Function
                    If StrComp(Left(FontName, Len(nm) + 1), nm & " ", vbTextCompare) = 0 Then FontExists = True: Exit Function
                Next i
            End If
        End If
    Next hv

    FontExists = Not enumOK
    Exit Function
Fallback:
    FontExists = True
End Function

Private Sub SetupFonts()
    If FontExists(FONT_ARABIC) Then g_ActiveArabicFont = FONT_ARABIC Else g_ActiveArabicFont = FONT_FALLBACK
    If FontExists(FONT_NUMERIC) Then g_ActiveNumericFont = FONT_NUMERIC Else g_ActiveNumericFont = FONT_FALLBACK
End Sub

Private Function SrcProcessFile(ByRef wbResult As Workbook, _
                                ByVal SourcePath As String, _
                                ByVal fileName As String, _
                                ByVal SheetNumber As Long, _
                                ByRef ReasonOut As String) As Boolean
    Dim wbSource As Workbook, wsSource As Worksheet, wsResult As Worksheet
    Dim LastRow As Long, LastCol As Long, DataArr As Variant
    Dim errNum As Long, errDesc As String, lastDataRow As Long, wrapEndRow As Long
    Dim idLastCol As Long, lastFind As Range
    Dim tbl As ListObject, tblRange As Range, seenHeaders As Object
    Dim h As Long, headerText As String, genCounter As Long

    SrcProcessFile = False
    ReasonOut = ""
    On Error GoTo EH

    On Error Resume Next
    Set wbSource = Workbooks.Open(fileName:=SourcePath & fileName, UpdateLinks:=False, _
                                  ReadOnly:=True, AddToMRU:=False, Password:="")
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wbSource Is Nothing Or errNum <> 0 Then
        ReasonOut = "OPEN_FAILED | " & errDesc: Exit Function
    End If

    If wbSource.Sheets.Count = 0 Then
        ReasonOut = "EMPTY_WORKBOOK (no sheets)": wbSource.Close SaveChanges:=False: Exit Function
    End If

    Set wsSource = wbSource.Sheets(1)

    On Error Resume Next
    LastCol = wsSource.Cells(DATA_START_ROW, wsSource.Columns.Count).End(xlToLeft).Column
    Set lastFind = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                                       SearchDirection:=xlPrevious)
    If lastFind Is Nothing Then LastRow = 0 Else LastRow = lastFind.Row - TRAILING_ROWS_SKIP
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Then
        ReasonOut = "DETECTION_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    If LastRow < DATA_START_ROW Or LastCol < 1 Then
        ReasonOut = "EMPTY_WORKBOOK (no data at row " & DATA_START_ROW & ")"
        wbSource.Close SaveChanges:=False: Exit Function
    End If

    On Error Resume Next
    Set wsResult = wbResult.Sheets.Add(After:=wbResult.Sheets(wbResult.Sheets.Count))
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wsResult Is Nothing Then
        ReasonOut = "ADD_SHEET_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    On Error Resume Next
    wsResult.Name = Format(SheetNumber, "00")
    Err.Clear
    On Error GoTo EH

    wsResult.DisplayRightToLeft = True

    On Error Resume Next
    DataArr = wsSource.Range(wsSource.Cells(DATA_START_ROW, 1), wsSource.Cells(LastRow, LastCol)).Value
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Or Not IsArray(DataArr) Then
        ReasonOut = "DATA_READ_FAILED | " & errDesc
        wbSource.Close SaveChanges:=False
        wsResult.Delete
        Exit Function
    End If

    lastDataRow = 1 + UBound(DataArr, 1)

    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        If LastCol >= 4 Then .Columns("D").NumberFormat = "@"
        If LastCol >= 5 Then .Columns("E").NumberFormat = "@"
        If LastCol >= 6 Then .Range(.Columns("F"), .Columns(LastCol)).NumberFormat = "0.00"
    End With

    wsResult.Range("A2").Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    genCounter = 0

    For h = 1 To LastCol
        headerText = Trim(CStr(wsResult.Cells(2, h).Value))
        If headerText = "" Or seenHeaders.Exists(headerText) Then
            genCounter = genCounter + 1
            Do While seenHeaders.Exists("Column" & genCounter): genCounter = genCounter + 1: Loop
            wsResult.Cells(2, h).Value = "Column" & genCounter
            seenHeaders.Add "Column" & genCounter, True
        Else
            seenHeaders.Add headerText, True
        End If
    Next h

    Set tblRange = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(lastDataRow, LastCol))
    Set tbl = Nothing

    On Error Resume Next
    Set tbl = wsResult.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number = 0 Then
        tbl.Name = "Data" & Format(SheetNumber, "00")
        tbl.TableStyle = ""
        tbl.ShowTableStyleColumnStripes = False
        tbl.ShowTableStyleRowStripes = False
        tbl.ShowTableStyleFirstColumn = False
        tbl.ShowTableStyleLastColumn = False
        tbl.ShowAutoFilter = True
    End If
    Err.Clear
    On Error GoTo EH

    With wsResult
        .Columns("B").Font.Name = g_ActiveArabicFont: .Columns("B").Font.Size = FONT_SIZE
        .Columns("C").Font.Name = g_ActiveNumericFont: .Columns("C").Font.Size = FONT_SIZE
        If LastCol >= 4 Then .Columns("D").Font.Name = g_ActiveArabicFont: .Columns("D").Font.Size = FONT_SIZE
        If LastCol >= 5 Then .Columns("E").Font.Name = g_ActiveArabicFont: .Columns("E").Font.Size = FONT_SIZE
        If LastCol >= 6 Then
            With .Range(.Columns("F"), .Columns(LastCol))
                .Font.Name = g_ActiveNumericFont
                .Font.Size = FONT_SIZE
            End With
        End If

        wrapEndRow = lastDataRow
        If wrapEndRow < 2 Then wrapEndRow = 2
        .Range(.Cells(2, 1), .Cells(wrapEndRow, LastCol)).WrapText = False

        With .Range(.Cells(1, 1), .Cells(lastDataRow, LastCol))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        idLastCol = IDENTITY_LAST_COL
        If LastCol < idLastCol Then idLastCol = LastCol
        .Range(.Columns(1), .Columns(idLastCol)).EntireColumn.AutoFit

        If LastCol >= 6 Then .Range(.Columns(IDENTITY_LAST_COL + 1), .Columns(LastCol)).ColumnWidth = COL_WIDTH

        If LastCol >= 4 Then
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Orientation = xlUpward
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Font.Bold = True
        End If
    End With

    On Error Resume Next
    wsResult.Range("A1").Value = CleanText(CStr(wsSource.Range("A1").Value)) & "|" & _
                                 CleanText(CStr(wsSource.Range("A2").Value)) & "|" & _
                                 CleanText(CStr(wsSource.Range("A3").Value)) & "|" & _
                                 CleanText(CStr(wsSource.Range("A4").Value))
    Err.Clear
    On Error GoTo EH

    PlaceHeaderTitle wsSource, wsResult

    On Error Resume Next
    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing
    Err.Clear
    On Error GoTo 0

    SrcProcessFile = True
    Exit Function

EH:
    ReasonOut = "UNEXPECTED | " & Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If Not wsResult Is Nothing Then wsResult.Delete
    Err.Clear
    SrcProcessFile = False
End Function

Private Sub SrcCollectAll(ByRef wbResult As Workbook, _
                          ByVal SourcePath As String, _
                          ByRef CopiedCount As Long, _
                          ByRef SkippedCount As Long, _
                          ByRef SkippedList As String)
    Dim fileName As String, SheetNumber As Long, Reason As String
    Dim names As Collection, sorted As Variant, i As Long
    Dim f As String, ext As String, baseName As String, dotPos As Long

    CopiedCount = 0: SkippedCount = 0: SheetNumber = 1: SkippedList = ""

    Set names = New Collection
    fileName = Dir(SourcePath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            dotPos = InStrRev(fileName, ".")
            ext = LCase(Mid(fileName, dotPos + 1))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Then names.Add fileName
        End If
        fileName = Dir
    Loop

    If names.Count = 0 Then
        SkippedList = "  - " & SOURCE_FOLDER_NAME & " folder is empty or contains no Excel files:" & vbCrLf & _
                      "    " & SourcePath & vbCrLf
        Exit Sub
    End If

    sorted = SortedNames(names)
    SetupFonts

    For i = LBound(sorted) To UBound(sorted)
        f = CStr(sorted(i))
        dotPos = InStrRev(f, ".")
        ext = LCase(Mid(f, dotPos + 1))
        baseName = Left(f, dotPos - 1)

        If StrComp(SourcePath & f, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            ' skip macro workbook
        ElseIf IsOpenInExcel(SourcePath & f) Then
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : ALREADY OPEN in Excel (close it and run again)" & vbCrLf
        ElseIf ext <> "xlsx" And FileExistsPath(SourcePath & baseName & ".xlsx") Then
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : DUPLICATE (.xlsx version exists)" & vbCrLf
        ElseIf SrcProcessFile(wbResult, SourcePath, f, SheetNumber, Reason) Then
            CopiedCount = CopiedCount + 1: SheetNumber = SheetNumber + 1
        Else
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : " & Reason & vbCrLf
        End If
    Next i
End Sub

Private Function IsOpenInExcel(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsOpenInExcel = True
            Exit Function
        End If
    Next wb
End Function

Private Function SrcBuildWorkbook(ByVal srcFolder As String, _
                                  ByRef wbOut As Workbook, _
                                  ByRef copied As Long, _
                                  ByRef skipped As Long, _
                                  ByRef skippedList As String) As Boolean
    Dim wb As Workbook, oldSec As Long, oldAlerts As Boolean

    SrcBuildWorkbook = False
    Set wbOut = Nothing
    copied = 0: skipped = 0: skippedList = ""

    If LCase(Left(srcFolder, 4)) = "http" Then
        skippedList = "  - The workbook is on OneDrive/SharePoint (URL path). Move it to a local folder first." & vbCrLf
        Exit Function
    End If
    If Right(srcFolder, 1) <> "\" Then srcFolder = srcFolder & "\"
    If Dir(srcFolder, vbDirectory) = "" Then
        skippedList = "  - " & SOURCE_FOLDER_NAME & " folder not found: " & srcFolder & vbCrLf
        Exit Function
    End If

    oldSec = Application.AutomationSecurity
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    On Error GoTo EH

    Set wb = Workbooks.Add
    Do While wb.Sheets.Count > 1: wb.Sheets(wb.Sheets.Count).Delete: Loop

    SrcCollectAll wb, srcFolder, copied, skipped, skippedList

    If copied > 0 Then
        If wb.Sheets.Count > 1 Then
            If Left(wb.Sheets(1).Name, 5) = "Sheet" And wb.Sheets(1).UsedRange.Cells.Count = 1 Then wb.Sheets(1).Delete
        End If
        Set wbOut = wb
        SrcBuildWorkbook = True
    Else
        wb.Close SaveChanges:=False
    End If

Done:
    Application.AutomationSecurity = oldSec
    Application.DisplayAlerts = oldAlerts
    Exit Function

EH:
    skippedList = skippedList & "  - Build failed: " & Err.Number & " : " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wbOut = Nothing
    SrcBuildWorkbook = False
    Resume Done
End Function

Private Sub PlaceHeaderTitle(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet)
    Dim rawText As String
    Dim partTitle As String
    Dim partGroup As String

    rawText = CleanText(CStr(wsSource.Range("A5").Value))
    SplitAtSchoolYear rawText, partTitle, partGroup

    WriteMergedCell wsTarget.Range("B1:F1"), partTitle
    WriteMergedCell wsTarget.Range("H1:Q1"), partGroup
    wsTarget.Rows(1).RowHeight = HEADER_ROW_HEIGHT
End Sub

Private Function CleanText(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(160), " ")
    s = Application.WorksheetFunction.Clean(s)
    CleanText = Application.WorksheetFunction.Trim(s)
End Function

Private Sub SplitAtSchoolYear(ByVal txt As String, _
                              ByRef part1 As String, _
                              ByRef part2 As String)
    Dim rx As Object
    Dim m As Object
    Dim splitPos As Long
    Dim digit As String
    Dim seps As String

    part1 = txt
    part2 = ""
    If Len(txt) = 0 Then Exit Sub

    digit = "[0-9\u0660-\u0669]"
    seps = "[/\\\-\u2013\u2014]"

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = digit & "{4}\s*" & seps & "\s*" & digit & "{4}"

    If Not rx.Test(txt) Then
        rx.Pattern = digit & "{4}"
        If Not rx.Test(txt) Then Exit Sub
    End If

    Set m = rx.Execute(txt)
    splitPos = m(0).FirstIndex + m(0).Length

    part1 = Trim(Left(txt, splitPos))

    If splitPos < Len(txt) Then
        part2 = Trim(Mid(txt, splitPos + 1))
        Do While Len(part2) > 0
            If InStr(1, "-:/" & ChrW(&H2013) & ChrW(&H2014), Left(part2, 1)) = 0 Then Exit Do
            part2 = Trim(Mid(part2, 2))
        Loop
    End If
End Sub

Private Sub WriteMergedCell(ByVal rng As Range, ByVal txt As String)
    With rng
        .UnMerge
        .ClearContents
        .Merge
        .Value = txt
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
        .Font.Name = GetAvailableFont(PREFERRED_FONT, FALLBACK_FONT)
        .Font.Size = HEADER_FONT_SIZE
        .Font.Bold = True
    End With
End Sub

Private Function GetAvailableFont(ByVal wanted As String, _
                                  ByVal fallback As String) As String
    If IsFontInstalled(wanted) Then
        GetAvailableFont = wanted
    Else
        GetAvailableFont = fallback
    End If
End Function

Private Function IsFontInstalled(ByVal fontName As String) As Boolean
    Dim folders(1 To 2) As String
    Dim i As Long

    folders(1) = Environ("WINDIR") & "\Fonts\"
    folders(2) = Environ("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\"

    On Error Resume Next
    For i = 1 To 2
        If Len(Dir(folders(i) & fontName & "*.*")) > 0 Then
            IsFontInstalled = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

' ==============================================================================
' PUBLIC API
' ==============================================================================
Public Function LoadSource(ByRef wbGlobal As Workbook, ByRef sheetIndex As Object, _
                           ByRef dupKeys As Object, ByRef report As String) As Boolean
    Dim srcFolder As String, copied As Long, skipped As Long, lst As String
    Dim ws As Worksheet, k As String

    LoadSource = False
    report = ""

    If Len(ThisWorkbook.Path) = 0 Then
        report = "  - Save the tool workbook first (its folder is needed to find " & SOURCE_FOLDER_NAME & ")." & vbCrLf
        Exit Function
    End If
    If Not FolderExistsPath(ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME) Then
        report = "  - " & SOURCE_FOLDER_NAME & " folder not found next to the tool." & vbCrLf
        Exit Function
    End If

    srcFolder = ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME & "\"
    Set wbGlobal = Nothing
    If Not SrcBuildWorkbook(srcFolder, wbGlobal, copied, skipped, lst) Then
        report = lst
        Exit Function
    End If
    report = lst

    Set sheetIndex = CreateObject("Scripting.Dictionary")
    sheetIndex.CompareMode = vbTextCompare
    Set dupKeys = CreateObject("Scripting.Dictionary")
    dupKeys.CompareMode = vbTextCompare

    For Each ws In wbGlobal.Worksheets
        k = NormalizeKey(CellText(ws.Range("H1")))
        If k <> "" Then
            If sheetIndex.Exists(k) Then
                If Not dupKeys.Exists(k) Then dupKeys.Add k, True
            Else
                sheetIndex.Add k, ws
            End If
        End If
    Next ws

    LoadSource = True
End Function

' File stems of the Term 01 analysis files
Public Function GetFileMap() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    AddGroup m, "1LitAnalysis01", AR("623 648 644 649 20 62B 627 646 648 64A 20 62C 62F 639 20 645 634 62A 631 643 20 622 62F 627 628"), 1
    AddGroup m, "1TecnoGlobalAnalysis01", AR("623 648 644 649 20 62B 627 646 648 64A 20 62C 62F 639 20 645 634 62A 631 643 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627"), 4
    AddGroup m, "2GEAnalysis01", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F"), 1
    AddGroup m, "2LEAnalysis01", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 644 63A 627 62A 20 623 62C 646 628 64A 629 20 623 644 645 627 646 64A 629"), 1
    AddGroup m, "2MathAnalysis01", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 631 64A 627 636 64A 627 62A"), 1
    AddGroup m, "2PhiloAnalysis01", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 622 62F 627 628 20 648 641 644 633 641 629"), 1
    AddGroup m, "2SEGlobalAnalysis01", AR("62B 627 646 64A 629 20 62B 627 646 648 64A 20 639 644 648 645 20 62A 62C 631 64A 628 64A 629"), 2
    AddGroup m, "3GEAnalysis01", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F"), 1
    AddGroup m, "3LEAnalysis01", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 644 63A 627 62A 20 623 62C 646 628 64A 629 20 623 644 645 627 646 64A 629"), 1
    AddGroup m, "3MathAnalysis01", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 631 64A 627 636 64A 627 62A"), 1
    AddGroup m, "3PhiloAnalysis01", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 622 62F 627 628 20 648 641 644 633 641 629"), 1
    AddGroup m, "3SEGlobalAnalysis01", AR("62B 627 644 62B 629 20 62B 627 646 648 64A 20 639 644 648 645 20 62A 62C 631 64A 628 64A 629"), 2
    Set GetFileMap = m
End Function

Private Sub AddGroup(ByVal m As Object, ByVal stem As String, ByVal prefix As String, ByVal cnt As Long)
    Dim a() As String, i As Long
    ReDim a(0 To cnt - 1)
    For i = 1 To cnt
        a(i - 1) = prefix & " " & Format$(i, "00")
    Next i
    m.Add stem, a
End Sub

Public Function GroupReady(ByVal sheetIndex As Object, ByVal dupKeys As Object, _
                           ByVal titles As Variant, ByRef why As String) As Boolean
    Dim i As Long, k As String, ok As Boolean, wsT As Worksheet
    ok = True
    why = ""
    For i = LBound(titles) To UBound(titles)
        k = NormalizeKey(CStr(titles(i)))
        If Not sheetIndex.Exists(k) Then
            ok = False: why = why & "    missing section: " & k & vbCrLf
        ElseIf dupKeys.Exists(k) Then
            ok = False: why = why & "    duplicate section: " & k & vbCrLf
        Else
            Set wsT = sheetIndex(k)
            If LastUsedRow(wsT) < 3 Then
                ok = False: why = why & "    empty section: " & k & vbCrLf
            End If
        End If
    Next i
    GroupReady = ok
End Function

Private Function ClassTitleOf(ByVal key As String) As String
    key = NormalizeKey(key)
    If Len(key) > 3 Then
        If Mid$(key, Len(key) - 2, 1) = " " And IsNumeric(Right$(key, 2)) Then key = Left$(key, Len(key) - 3)
    End If
    ClassTitleOf = key
End Function

' ==============================================================================
' EXTRA COLUMNS OF THE DATA SHEET (as in the models; filled in by hand later)
'   "G:<title>"  starts a banner group (row 5)      "T:"  starts the tail columns
'   "H:<header>" one column. Texts are Unicode code points in hexadecimal.
' ==============================================================================
Private Function ExtraSpec(ByVal stem As String) As Collection
    Dim c As Collection
    Set c = New Collection
    Select Case stem
        Case "1LitAnalysis01"
            c.Add "G:646 62A 627 626 62C 20 634 647 627 62F 629 20 627 644 62A 639 644 64A 645 20 627 644 645 62A 648 633 637"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 33"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 33"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 33"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 33"
            c.Add "H:639 20 627 644 641 64A 632 64A 627 626 64A 629 20 648 627 644 62A 643 646 648 644 648 62C 64A 627 33"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 33"
            c.Add "H:627 644 62A 631 628 64A 629 20 627 644 625 633 644 627 645 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 33"
            c.Add "H:645 639 62F 644 20 627 644 634 647 627 62F 629"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:627 644 62A 631 628 64A 629 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 32"
            c.Add "H:639 20 627 644 641 64A 632 64A 627 626 64A 629 20 648 627 644 62A 643 646 648 644 648 62C 64A 627 32"
            c.Add "H:627 644 645 639 644 648 645 627 62A 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "H:645 639 62F 644 20 627 644 627 646 62A 642 627 644"
            c.Add "T:"
            c.Add "H:627 644 645 623 645 646"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "1TecnoGlobalAnalysis01"
            c.Add "G:646 62A 627 626 62C 20 634 647 627 62F 629 20 627 644 62A 639 644 64A 645 20 627 644 645 62A 648 633 637"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 33"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 33"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 33"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 33"
            c.Add "H:639 20 627 644 641 64A 632 64A 627 626 64A 629 20 648 627 644 62A 643 646 648 644 648 62C 64A 627 33"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 33"
            c.Add "H:627 644 62A 631 628 64A 629 20 627 644 625 633 644 627 645 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 33"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33"
            c.Add "H:645 639 62F 644 20 627 644 634 647 627 62F 629"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 633 646 629 20 627 644 631 627 628 639 629 20 645 62A 648 633 637"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:627 644 62A 631 628 64A 629 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 32"
            c.Add "H:639 20 627 644 641 64A 632 64A 627 626 64A 629 20 648 627 644 62A 643 646 648 644 648 62C 64A 627 32"
            c.Add "H:627 644 645 639 644 648 645 627 62A 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "H:645 639 62F 644 20 627 644 627 646 62A 642 627 644"
            c.Add "T:"
            c.Add "H:627 644 645 623 645 646"
            c.Add "H:627 644 635 648 631 629"
            c.Add "H:645 644 627 62D 638 627 62A"
        Case "2GEAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "T:"
            c.Add "H:62A 2E 20 627 644 645 62D 627 633 628 64A 20 648 627 644 645 627 644 64A 32"
            c.Add "H:627 644 627 642 62A 635 627 62F 20 648 627 644 645 646 627 62C 645 646 62A 32"
            c.Add "H:627 644 642 627 646 648 646 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "2LEAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "H:627 644 644 63A 629 20 627 FEF7 62C 646 628 64A 629 20 627 644 62B 627 644 62B 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "T:"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "2MathAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629 32"
            c.Add "H:639 644 648 645 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 31 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "T:"
            c.Add "H:627 644 635 648 631 629"
            c.Add "H:645 644 627 62D 638 627 62A"
        Case "2PhiloAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629 32"
            c.Add "H:639 644 648 645 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "T:"
            c.Add "H:627 644 641 644 633 641 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "2SEGlobalAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629 32"
            c.Add "H:639 644 648 645 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "T:"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "3GEAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:62A 2E 20 627 644 645 62D 627 633 628 64A 20 648 627 644 645 627 644 64A 32"
            c.Add "H:627 644 627 642 62A 635 627 62F 20 648 627 644 645 646 627 62C 645 646 62A 32"
            c.Add "H:627 644 642 627 646 648 646 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "G:646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 33"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 33"
            c.Add "H:62A 2E 20 627 644 645 62D 627 633 628 64A 20 648 627 644 645 627 644 64A 33"
            c.Add "H:627 644 627 642 62A 635 627 62F 20 648 627 644 645 646 627 62C 645 646 62A 33"
            c.Add "H:627 644 642 627 646 648 646 33"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 33"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 33"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 33"
            c.Add "H:627 644 641 644 633 641 629 33"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33"
            c.Add "H:627 644 645 639 62F 644 20 627 644 639 627 645"
            c.Add "T:"
            c.Add "H:627 644 641 644 633 641 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "3LEAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 FEF7 62C 646 628 64A 629 20 627 644 62B 627 644 62B 629 32"
            c.Add "H:627 644 631 64A 640 627 636 64A 627 62A 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "G:646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 33"
            c.Add "H:627 644 641 644 633 641 629 33"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 33"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 FEF7 62C 646 628 64A 629 20 627 644 62B 627 644 62B 629 33"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 33"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 639 627 645"
            c.Add "T:"
            c.Add "H:627 644 641 644 633 641 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "3MathAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629 32"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "G:646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 33"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 33"
            c.Add "H:627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629 33"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 33"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 33"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 33"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 33"
            c.Add "H:627 644 641 644 633 641 629 33"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33"
            c.Add "H:627 644 645 639 62F 644 20 627 644 639 627 645"
            c.Add "T:"
            c.Add "H:627 644 641 644 633 641 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "3PhiloAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 641 644 633 641 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "G:646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 33"
            c.Add "H:627 644 641 644 633 641 629 33"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 33"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 33"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 33"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33"
            c.Add "H:627 644 645 639 62F 644 20 627 644 639 627 645"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "T:"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case "3SEGlobalAnalysis01"
            c.Add "G:627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629 20 644 644 639 627 645 20 627 644 645 627 636 64A"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 32"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629 32"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 32"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 32"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 32"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 32"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 32"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A"
            c.Add "G:646 62A 627 626 62C 20 627 644 628 643 627 644 648 631 64A 627"
            c.Add "H:627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627 33"
            c.Add "H:627 644 631 64A 627 636 64A 627 62A 33"
            c.Add "H:627 644 639 644 648 645 20 627 644 641 64A 632 64A 627 626 64A 629 33"
            c.Add "H:639 20 627 644 637 628 64A 639 629 20 648 627 644 62D 64A 627 629 33"
            c.Add "H:627 644 639 644 648 645 20 627 644 625 633 644 627 645 64A 629 33"
            c.Add "H:627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627 33"
            c.Add "H:627 644 644 63A 629 20 627 644 641 631 646 633 64A 629 33"
            c.Add "H:627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629 33"
            c.Add "H:627 644 641 644 633 641 629 33"
            c.Add "H:62A 20 627 644 628 62F 646 64A 629 20 648 627 644 631 64A 627 636 64A 629 33"
            c.Add "H:627 644 645 639 62F 644 20 627 644 639 627 645"
            c.Add "T:"
            c.Add "H:627 644 641 644 633 641 629 32"
            c.Add "H:627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 644 644 645 639 64A 62F 64A 646"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
        Case Else
            c.Add "T:"
            c.Add "H:645 644 627 62D 638 627 62A"
            c.Add "H:627 644 635 648 631 629"
    End Select
    Set ExtraSpec = c
End Function

Private Function IsTextTail(ByVal h As String) As Boolean
    h = NormalizeKey(h)
    IsTextTail = (StrComp(h, ArT("645 644 627 62D 638 627 62A"), vbTextCompare) = 0) Or (StrComp(h, ArT("627 644 645 644 627 62D 638 627 62A"), vbTextCompare) = 0) Or (StrComp(h, ArT("627 644 635 648 631 629"), vbTextCompare) = 0) _
              Or (StrComp(h, ArT("627 644 645 623 645 646"), vbTextCompare) = 0)
End Function

' ==============================================================================
' DATA SHEET BUILDER (term 1 marks + section + extra columns)
' ==============================================================================
Public Function BuildDataSheet(ByVal wbNew As Workbook, ByVal sheetIndex As Object, _
                               ByVal titles As Variant, ByVal stem As String, _
                               ByRef errText As String) As Boolean
    Dim ws As Worksheet, wsSrc As Worksheet, wsFirst As Worksheet
    Dim nSec As Long, i As Long, c As Long, r As Long, rr As Long
    Dim W As Long, total As Long, lastRow As Long, lastCol As Long
    Dim rowsOf() As Long, secVal() As Long, hdrBase() As String
    Dim key As String, classTitle As String, txt As String, h As String, nm As String
    Dim arr As Variant, outArr() As Variant, hdrRow() As Variant
    Dim seen As Object, gen As Long
    Dim lo As ListObject, lines() As String
    Dim lastC As Long, fRow As Long, lRow As Long, secStart As Long
    Dim spec As Collection, it As Variant, tag As String, body As String
    Dim nEx As Long, nG As Long, inTail As Boolean, lastNum As Long, titleEnd As Long
    Dim exH() As String, exKind() As Long
    Dim gStart() As Long, gEnd() As Long, gTitle() As String

    BuildDataSheet = False
    errText = ""
    On Error GoTo EH

    nSec = UBound(titles) - LBound(titles) + 1
    ReDim rowsOf(1 To nSec)
    ReDim secVal(1 To nSec)

    For i = 1 To nSec
        key = NormalizeKey(CStr(titles(LBound(titles) + i - 1)))
        Set wsSrc = sheetIndex(key)
        lastCol = wsSrc.Cells(2, wsSrc.Columns.Count).End(xlToLeft).Column
        lastRow = LastUsedRow(wsSrc)
        rowsOf(i) = lastRow - 2
        If rowsOf(i) < 1 Then Err.Raise vbObjectError + 801, "BuildDataSheet", "Empty section: " & key
        If i = 1 Then
            Set wsFirst = wsSrc
            W = lastCol
            If W < 7 Then Err.Raise vbObjectError + 804, "BuildDataSheet", "Too few columns in: " & key
            ReDim hdrBase(1 To W)
            For c = 1 To W
                hdrBase(c) = NormalizeKey(CellText(wsSrc.Cells(2, c)))
            Next c
            classTitle = ClassTitleOf(key)
        Else
            If lastCol <> W Then Err.Raise vbObjectError + 802, "BuildDataSheet", "Column count differs in: " & key
            For c = 1 To W
                If StrComp(NormalizeKey(CellText(wsSrc.Cells(2, c))), hdrBase(c), vbTextCompare) <> 0 Then
                    Err.Raise vbObjectError + 803, "BuildDataSheet", "Header differs (column " & c & ") in: " & key
                End If
            Next c
        End If
        If IsNumeric(Right$(key, 2)) Then secVal(i) = CLng(Right$(key, 2)) Else secVal(i) = i
        total = total + rowsOf(i)
    Next i

    ' ---- the extra columns of this file (model)
    Set spec = ExtraSpec(stem)
    ReDim exH(1 To spec.Count + 1)
    ReDim exKind(1 To spec.Count + 1)
    ReDim gStart(1 To spec.Count + 1)
    ReDim gEnd(1 To spec.Count + 1)
    ReDim gTitle(1 To spec.Count + 1)
    For Each it In spec
        tag = Left$(CStr(it), 2)
        body = Mid$(CStr(it), 3)
        Select Case tag
            Case "G:"
                nG = nG + 1
                gTitle(nG) = AR(body)
                gStart(nG) = nEx + 1
                gEnd(nG) = nEx
                inTail = False
            Case "T:"
                inTail = True
            Case "H:"
                nEx = nEx + 1
                exH(nEx) = AR(body)
                If inTail Then
                    If IsTextTail(exH(nEx)) Then exKind(nEx) = 2 Else exKind(nEx) = 3
                Else
                    exKind(nEx) = 1
                    gEnd(nG) = nEx
                End If
        End Select
    Next it

    lastC = W + 1 + nEx
    ReDim hdrRow(1 To lastC)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For c = 1 To W
        h = hdrBase(c)
        If h = "" Or seen.Exists(h) Then
            Do
                gen = gen + 1
                h = "Column" & gen
            Loop While seen.Exists(h)
        End If
        seen.Add h, True
        hdrRow(c) = h
    Next c
    ' section column, then the extra columns
    nm = AR("627 644 642 633 645")
    Do While seen.Exists(nm)
        nm = nm & " 2"
    Loop
    seen.Add nm, True
    hdrRow(W + 1) = nm
    For i = 1 To nEx
        nm = exH(i)
        Do While seen.Exists(nm)
            nm = nm & " 2"
        Loop
        seen.Add nm, True
        hdrRow(W + 1 + i) = nm
    Next i

    ReDim outArr(1 To total, 1 To W)
    rr = 0
    For i = 1 To nSec
        key = NormalizeKey(CStr(titles(LBound(titles) + i - 1)))
        Set wsSrc = sheetIndex(key)
        arr = wsSrc.Range(wsSrc.Cells(3, 1), wsSrc.Cells(2 + rowsOf(i), W)).Value
        For r = 1 To rowsOf(i)
            rr = rr + 1
            For c = 1 To W
                outArr(rr, c) = arr(r, c)
            Next c
        Next r
    Next i

    Set ws = wbNew.Worksheets(1)
    ws.Name = DATA_SHEET
    ws.DisplayRightToLeft = True
    fRow = DATA_FIRST_ROW
    lRow = DATA_FIRST_ROW + total - 1

    With ws
        .Range(.Cells(DATA_HDR_ROW, 1), .Cells(DATA_HDR_ROW, lastC)).NumberFormat = "@"
        .Range(.Cells(DATA_HDR_ROW, 1), .Cells(DATA_HDR_ROW, lastC)).Value = hdrRow

        .Range(.Cells(fRow, 1), .Cells(lRow, 1)).NumberFormat = FMT_COUNT
        .Range(.Cells(fRow, 2), .Cells(lRow, 2)).NumberFormat = "@"
        .Range(.Cells(fRow, 3), .Cells(lRow, 3)).NumberFormat = "yyyy-mm-dd"
        .Range(.Cells(fRow, 4), .Cells(lRow, 5)).NumberFormat = "@"
        .Range(.Cells(fRow, 6), .Cells(lRow, W)).NumberFormat = FMT_DEC
        .Range(.Cells(fRow, W + 1), .Cells(lRow, W + 1)).NumberFormat = FMT_COUNT
        For i = 1 To nEx
            If exKind(i) = 2 Then
                .Range(.Cells(fRow, W + 1 + i), .Cells(lRow, W + 1 + i)).NumberFormat = "@"
            Else
                .Range(.Cells(fRow, W + 1 + i), .Cells(lRow, W + 1 + i)).NumberFormat = FMT_DEC
            End If
        Next i

        .Cells(fRow, 1).Resize(total, W).Value = outArr

        secStart = fRow
        For i = 1 To nSec
            .Range(.Cells(secStart, W + 1), .Cells(secStart + rowsOf(i) - 1, W + 1)).Value = secVal(i)
            secStart = secStart + rowsOf(i)
        Next i
    End With

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(DATA_HDR_ROW, 1), ws.Cells(lRow, lastC)), , xlYes)
    lo.Name = DATA_TABLE
    lo.TableStyle = ""
    lo.ShowTableStyleRowStripes = False
    lo.ShowTableStyleColumnStripes = False
    lo.ShowTableStyleFirstColumn = False
    lo.ShowTableStyleLastColumn = False
    lo.ShowAutoFilter = True

    ' ---- banner lines (rows 1 to 4), class title and group banners (row 5)
    titleEnd = W + 1
    If nG > 0 Then
        If gEnd(nG) >= gStart(nG) Then
            If W + 1 + gEnd(nG) > titleEnd Then titleEnd = W + 1 + gEnd(nG)
        End If
    End If
    lines = Split(CellText(wsFirst.Range("A1")), "|")
    For r = 1 To 4
        txt = ""
        If UBound(lines) >= r - 1 Then txt = Trim$(lines(r - 1))
        MergeText ws, r, 1, titleEnd, txt
    Next r
    MergeText ws, 5, 1, 5, classTitle
    MergeText ws, 5, 6, W, AR("646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644")
    For i = 1 To nG
        If gEnd(i) >= gStart(i) Then MergeText ws, 5, W + 1 + gStart(i), W + 1 + gEnd(i), gTitle(i)
    Next i

    lastNum = lastC
    Do While lastNum > W + 1
        If exKind(lastNum - W - 1) <> 2 Then Exit Do
        lastNum = lastNum - 1
    Loop

    With ws
        StyleGrid .Range(.Cells(DATA_HDR_ROW, 1), .Cells(lRow, lastC))
        StyleHead .Range(.Cells(DATA_HDR_ROW, 1), .Cells(DATA_HDR_ROW, lastC))
        .Range(.Cells(DATA_HDR_ROW, 4), .Cells(DATA_HDR_ROW, lastNum)).Orientation = 90
        .Rows(DATA_HDR_ROW).RowHeight = HDR_ROW_HEIGHT

        With .Range(.Cells(fRow, 2), .Cells(lRow, 2))
            .Font.Name = gFontName: .Font.Size = 14
            .HorizontalAlignment = xlRight
            .ShrinkToFit = True
        End With
        With .Range(.Cells(fRow, 4), .Cells(lRow, 5))
            .Font.Name = gFontName: .Font.Size = 13
            .ShrinkToFit = True
        End With
        With .Range(.Cells(fRow, 3), .Cells(lRow, 3))
            .Font.Size = 14
        End With

        .Columns(1).ColumnWidth = ID_COL_WIDTH
        .Columns(2).ColumnWidth = 30
        .Columns(3).ColumnWidth = DATE_COL_MIN_WIDTH + 1
        .Range(.Columns(4), .Columns(lastNum)).ColumnWidth = NUM_COL_WIDTH
        For i = 1 To nEx
            If exKind(i) = 2 Then
                With .Range(.Cells(fRow, W + 1 + i), .Cells(lRow, W + 1 + i))
                    .Font.Name = gFontName: .Font.Size = 13
                    .HorizontalAlignment = xlRight
                    .WrapText = True
                End With
                If StrComp(NormalizeKey(exH(i)), AR("627 644 635 648 631 629"), vbTextCompare) = 0 Then
                    .Columns(W + 1 + i).ColumnWidth = 14
                Else
                    .Columns(W + 1 + i).ColumnWidth = 30
                End If
            End If
        Next i

        For r = 1 To 4
            .Range(.Cells(r, 1), .Cells(r, titleEnd)).Font.Name = gFontName
            .Range(.Cells(r, 1), .Cells(r, titleEnd)).Font.Size = 14
            .Range(.Cells(r, 1), .Cells(r, titleEnd)).HorizontalAlignment = xlCenter
            .Range(.Cells(r, 1), .Cells(r, titleEnd)).ReadingOrder = xlRTL
            .Rows(r).RowHeight = 24
        Next r
        With .Range(.Cells(5, 1), .Cells(5, 5))
            .Font.Name = gFontBase: .Font.Size = 16: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
            .ReadingOrder = xlRTL
        End With
        StyleTableTitle .Range(.Cells(5, 6), .Cells(5, W))
        For i = 1 To nG
            If gEnd(i) >= gStart(i) Then StyleTableTitle .Range(.Cells(5, W + 1 + gStart(i)), .Cells(5, W + 1 + gEnd(i)))
        Next i
        .Rows(5).RowHeight = 32
    End With

    BuildDataSheet = True
    Exit Function
EH:
    errText = "BuildDataSheet: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function
