Attribute VB_Name = "modResultFlow01"
Option Explicit

'==============================================================================
' Module  : ResultFlow01 (DistributeOnly_Final_01)
' Version : Final - Standalone Phase 2 (English, Dynamic Section Column)
' Purpose : Builds the Term 01 results in memory straight from \Source01\
'           (Collector01 logic embedded: same conversion to .xlsx and same
'           collection) and distributes the sections into target
'           workbooks according to Map, starting at row 7 of the "Data"
'           worksheet. Term01Results.xlsx is NOT needed and is never touched.
'           SELF-CONTAINED: needs no other module. It runs in any workbook that
'           has the Source01 folder and the target analysis files next to it.
'
' Variant : Target files use the "01" suffix (e.g. 1LitAnalysis01.xlsm).
'
' Notes   : - All Arabic text is stored as hex code points (see AR) so the
'             module is pure ASCII and safe to import on any system locale.
'           - Every target file is backed up to \_Backup\<timestamp>\ first.
'           - ALLOW_PARTIAL = False : a target file is NOT touched if any of
'             its required sections is missing in the source workbook.
'==============================================================================

Private Const SOURCE_FOLDER_NAME As String = "Source01"
Private Const BACKUP_FOLDER    As String = "_Backup"
Private Const ALLOW_PARTIAL    As Boolean = False
Private Const HEADER_ROW       As Long = 6
Private Const PASTE_START_ROW  As Long = 7

' ---- Collector01 constants (embedded) ----
Private Const DATA_START_ROW       As Long = 6      ' header row in the source sheet
Private Const TRAILING_ROWS_SKIP   As Long = 1      ' footer rows after last used row (original logic: -1). Set 0 if none
Private Const IDENTITY_LAST_COL    As Long = 5      ' columns A:E = identity columns (AutoFit)
Private Const COL_WIDTH            As Double = 5.7  ' width for columns F:LastCol
Private Const FONT_SIZE            As Long = 14
Private Const FONT_ARABIC          As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC         As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK        As String = "Calibri"
Private Const ARCHIVE_FOLDER       As String = "_Original"

Private Const PREFERRED_FONT As String = "Cairo"
Private Const FALLBACK_FONT  As String = "Arial"
Private Const HEADER_FONT_SIZE As Double = 11
Private Const HEADER_ROW_HEIGHT As Double = 30


Private g_ActiveArabicFont  As String
Private g_ActiveNumericFont As String

' ==============================================================================
' HELPERS
' ==============================================================================

' Builds a string from space-separated hex code points, e.g. AR("0627 0644")
Private Function AR(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    AR = s
End Function

' Official header text of the section column in the target sheets
Private Function SectionHeaderText() As String
    SectionHeaderText = AR("0627 0644 0642 0633 0645")
End Function

' Removes invisible differences: NBSP, direction marks, tatweel, double spaces
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

' Last row containing anything, in any column (0 if the sheet is empty)
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastUsedRow = 0 Else LastUsedRow = f.Row
End Function

' ==============================================================================
' MAIN
' ==============================================================================
Public Sub DistributeDataFromTerm01Results()

    ' ==================== Pre-flight checks ====================
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save this workbook before running the macro.", vbExclamation, "Distribution Aborted"
        Exit Sub
    End If

    If LCase(Left(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "This workbook is on OneDrive/SharePoint (URL path)." & vbCrLf & _
               "Move it to a local folder, or sync the folder, then run again.", _
               vbExclamation, "Distribution Aborted"
        Exit Sub
    End If

    ' ==================== Performance & Security settings ====================
    ' NOTE: calculation mode is deliberately NOT changed (it is stored inside
    '       the saved workbooks and would leak into the target files).
    Dim prevScreen  As Boolean
    Dim prevEvents  As Boolean
    Dim prevAlerts  As Boolean
    Dim prevSec     As MsoAutomationSecurity

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevSec = Application.AutomationSecurity

    On Error GoTo UnexpectedError

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    ' ==================== Variables ====================
    Dim wbGlobal        As Workbook
    Dim wsGlobal        As Worksheet
    Dim wbTarget        As Workbook
    Dim wsTarget        As Worksheet
    Dim SourcePath      As String
    Dim LastRow         As Long
    Dim LastCol         As Long
    Dim PasteRow        As Long
    Dim RowsCopied      As Long
    Dim Map             As Object
    Dim SheetIndex      As Object
    Dim UsedSheets      As Object
    Dim Key             As Variant
    Dim kSheet          As Variant
    Dim keyName         As String
    Dim targetPath      As String
    Dim TextH1          As String
    Dim SuccessCount    As Long
    Dim SheetCount      As Long
    Dim SrcFolder       As String
    Dim SrcCopied       As Long
    Dim SrcSkipped      As Long
    Dim SrcList         As String
    Dim ConvTotal       As Long
    Dim ConvOK          As Long
    Dim ConvFailed      As Long
    Dim DataArr         As Variant
    Dim TargetKeys      As Variant
    Dim i               As Long
    Dim SectionNumber   As String
    Dim HasValidSection As Boolean
    Dim FileMissing     As Long
    Dim FileCopiedRows  As Long
    Dim FileSections    As Long
    Dim saveErr         As Long
    Dim saveDesc        As String
    Dim RunStamp        As String
    Dim backupDir       As String
    Dim backupReady     As Boolean
    Dim backupErr       As Long
    Dim errMsg          As String

    Dim SkippedFiles     As String
    Dim SkippedCount     As Long
    Dim MissingSections  As String
    Dim MissingCount     As Long
    Dim UnmappedSections As String
    Dim UnmappedCount    As Long
    Dim UnindexedSheets  As String
    Dim UnindexedCount   As Long
    Dim Report           As String

    ' ==================== Target file map (Term 01) ====================
    Set Map = CreateObject("Scripting.Dictionary")
    Map.CompareMode = vbTextCompare
    ' 1st year common core - Literature
    Map.Add "1LitAnalysis01.xlsm", Array(AR("0623 0648 0644 0649 0020 062B 0627 0646 0648 064A 0020 062C 062F 0639 0020 0645 0634 062A 0631 0643 0020 0622 062F 0627 0628 0020 0030 0031"))
    ' 1st year common core - Science & Technology (01-04)
    Map.Add "1TecnoGlobalAnalysis01.xlsm", Array( _
        AR("0623 0648 0644 0649 0020 062B 0627 0646 0648 064A 0020 062C 062F 0639 0020 0645 0634 062A 0631 0643 0020 0639 0644 0648 0645 0020 0648 062A 0643 0646 0648 0644 0648 062C 064A 0627 0020 0030 0031"), _
        AR("0623 0648 0644 0649 0020 062B 0627 0646 0648 064A 0020 062C 062F 0639 0020 0645 0634 062A 0631 0643 0020 0639 0644 0648 0645 0020 0648 062A 0643 0646 0648 0644 0648 062C 064A 0627 0020 0030 0032"), _
        AR("0623 0648 0644 0649 0020 062B 0627 0646 0648 064A 0020 062C 062F 0639 0020 0645 0634 062A 0631 0643 0020 0639 0644 0648 0645 0020 0648 062A 0643 0646 0648 0644 0648 062C 064A 0627 0020 0030 0033"), _
        AR("0623 0648 0644 0649 0020 062B 0627 0646 0648 064A 0020 062C 062F 0639 0020 0645 0634 062A 0631 0643 0020 0639 0644 0648 0645 0020 0648 062A 0643 0646 0648 0644 0648 062C 064A 0627 0020 0030 0034") _
        )
    ' 2nd year - Management & Economics
    Map.Add "2GEAnalysis01.xlsm", Array(AR("062B 0627 0646 064A 0629 0020 062B 0627 0646 0648 064A 0020 062A 0633 064A 064A 0631 0020 0648 0627 0642 062A 0635 0627 062F 0020 0030 0031"))
    ' 2nd year - Foreign Languages (German)
    Map.Add "2LEAnalysis01.xlsm", Array(AR("062B 0627 0646 064A 0629 0020 062B 0627 0646 0648 064A 0020 0644 063A 0627 062A 0020 0623 062C 0646 0628 064A 0629 0020 0623 0644 0645 0627 0646 064A 0629 0020 0030 0031"))
    ' 2nd year - Mathematics
    Map.Add "2MathAnalysis01.xlsm", Array(AR("062B 0627 0646 064A 0629 0020 062B 0627 0646 0648 064A 0020 0631 064A 0627 0636 064A 0627 062A 0020 0030 0031"))
    ' 2nd year - Literature & Philosophy
    Map.Add "2PhiloAnalysis01.xlsm", Array(AR("062B 0627 0646 064A 0629 0020 062B 0627 0646 0648 064A 0020 0622 062F 0627 0628 0020 0648 0641 0644 0633 0641 0629 0020 0030 0031"))
    ' 2nd year - Experimental Sciences (01-02)
    Map.Add "2SEGlobalAnalysis01.xlsm", Array( _
        AR("062B 0627 0646 064A 0629 0020 062B 0627 0646 0648 064A 0020 0639 0644 0648 0645 0020 062A 062C 0631 064A 0628 064A 0629 0020 0030 0031"), _
        AR("062B 0627 0646 064A 0629 0020 062B 0627 0646 0648 064A 0020 0639 0644 0648 0645 0020 062A 062C 0631 064A 0628 064A 0629 0020 0030 0032") _
        )
    ' 3rd year - Management & Economics
    Map.Add "3GEAnalysis01.xlsm", Array(AR("062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 062A 0633 064A 064A 0631 0020 0648 0627 0642 062A 0635 0627 062F 0020 0030 0031"))
    ' 3rd year - Foreign Languages (German)
    Map.Add "3LEAnalysis01.xlsm", Array(AR("062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0644 063A 0627 062A 0020 0623 062C 0646 0628 064A 0629 0020 0623 0644 0645 0627 0646 064A 0629 0020 0030 0031"))
    ' 3rd year - Mathematics
    Map.Add "3MathAnalysis01.xlsm", Array(AR("062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0631 064A 0627 0636 064A 0627 062A 0020 0030 0031"))
    ' 3rd year - Literature & Philosophy
    Map.Add "3PhiloAnalysis01.xlsm", Array(AR("062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0622 062F 0627 0628 0020 0648 0641 0644 0633 0641 0629 0020 0030 0031"))
    ' 3rd year - Experimental Sciences (01-02)
    Map.Add "3SEGlobalAnalysis01.xlsm", Array( _
        AR("062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0639 0644 0648 0645 0020 062A 062C 0631 064A 0628 064A 0629 0020 0030 0031"), _
        AR("062B 0627 0644 062B 0629 0020 062B 0627 0646 0648 064A 0020 0639 0644 0648 0645 0020 062A 062C 0631 064A 0628 064A 0629 0020 0030 0032") _
        )

    ' ==================== Source path ====================
    SourcePath = ThisWorkbook.Path & "\"
    RunStamp = Format(Now, "yyyymmdd_hhnnss")
    backupDir = SourcePath & BACKUP_FOLDER & "\" & RunStamp
    SkippedFiles = "": SkippedCount = 0
    MissingSections = "": MissingCount = 0
    UnmappedSections = "": UnmappedCount = 0
    UnindexedSheets = "": UnindexedCount = 0
    SuccessCount = 0

    ' ==================== Build the results workbook in memory from Source01 ====================
    SrcFolder = SourcePath & SOURCE_FOLDER_NAME & "\"

    If Not FolderExistsPath(SourcePath & SOURCE_FOLDER_NAME) Then
        MsgBox SOURCE_FOLDER_NAME & " folder was not found next to the tool." & vbCrLf & vbCrLf & _
               "Expected path:" & vbCrLf & SrcFolder, vbCritical, "Distribution Aborted"
        GoTo CleanExit
    End If

    Set wbGlobal = Nothing
    If Not SrcBuildWorkbook(SrcFolder, wbGlobal, SrcCopied, SrcSkipped, SrcList, _
                               ConvTotal, ConvOK, ConvFailed) Then
        MsgBox "No usable source file was found in " & SOURCE_FOLDER_NAME & "." & vbCrLf & _
               "Nothing was distributed." & vbCrLf & vbCrLf & SrcList, _
               vbCritical, "Distribution Aborted"
        GoTo CleanExit
    End If

    SheetCount = wbGlobal.Sheets.Count

    ' ==================== Build sheet index (H1 text -> Worksheet) ====================
    Set SheetIndex = CreateObject("Scripting.Dictionary")
    SheetIndex.CompareMode = vbTextCompare
    Set UsedSheets = CreateObject("Scripting.Dictionary")
    UsedSheets.CompareMode = vbTextCompare

    For Each wsGlobal In wbGlobal.Worksheets
        TextH1 = NormalizeKey(CellText(wsGlobal.Range("H1")))
        If TextH1 = "" Then
            UnindexedCount = UnindexedCount + 1
            UnindexedSheets = UnindexedSheets & "  - Sheet '" & wsGlobal.Name & "' (empty H1)" & vbCrLf
        ElseIf SheetIndex.Exists(TextH1) Then
            UnindexedCount = UnindexedCount + 1
            UnindexedSheets = UnindexedSheets & "  - Sheet '" & wsGlobal.Name & _
                              "' (duplicate H1, first sheet '" & SheetIndex(TextH1).Name & "' is used)" & vbCrLf
        Else
            SheetIndex.Add TextH1, wsGlobal
        End If
    Next wsGlobal

    ' ==================== Main loop over target files ====================
    For Each Key In Map.Keys

        On Error GoTo FileError

        keyName = CStr(Key)
        targetPath = SourcePath & keyName
        TargetKeys = Map(Key)
        HasValidSection = False
        FileMissing = 0

        ' Pre-check: which required sections exist in the source workbook
        For i = LBound(TargetKeys) To UBound(TargetKeys)
            TextH1 = NormalizeKey(CStr(TargetKeys(i)))
            If SheetIndex.Exists(TextH1) Then
                HasValidSection = True
            Else
                FileMissing = FileMissing + 1
                MissingCount = MissingCount + 1
                MissingSections = MissingSections & "  - " & TextH1 & _
                                  " (required by " & keyName & ")" & vbCrLf
            End If
        Next i

        If Not HasValidSection Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (skipped: no matching sections in " & SOURCE_FOLDER_NAME & ")" & vbCrLf
            GoTo NextKey
        End If

        If FileMissing > 0 And Not ALLOW_PARTIAL Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (skipped: " & FileMissing & " of " & _
                           (UBound(TargetKeys) - LBound(TargetKeys) + 1) & _
                           " required sections missing - file left untouched)" & vbCrLf
            GoTo NextKey
        End If

        If Not FileExistsPath(targetPath) Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (file not found)" & vbCrLf
            GoTo NextKey
        End If

        If IsWorkbookOpen(targetPath) Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (already open in this Excel session - close it first)" & vbCrLf
            GoTo NextKey
        End If

        ' ==================== Backup before touching the file ====================
        If Not backupReady Then
            backupReady = EnsureFolder(SourcePath & BACKUP_FOLDER) And EnsureFolder(backupDir)
        End If

        If Not backupReady Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (skipped: cannot create backup folder)" & vbCrLf
            GoTo NextKey
        End If

        On Error Resume Next
        FileCopy targetPath, backupDir & "\" & keyName
        backupErr = Err.Number
        Err.Clear
        On Error GoTo FileError

        If backupErr <> 0 Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (skipped: backup failed, file may be locked)" & vbCrLf
            GoTo NextKey
        End If

        ' ==================== Open target ====================
        Set wbTarget = Nothing
        On Error Resume Next
        Set wbTarget = Workbooks.Open(targetPath, _
                                      UpdateLinks:=False, _
                                      ReadOnly:=False, _
                                      AddToMRU:=False, _
                                      Password:="")
        On Error GoTo FileError

        If wbTarget Is Nothing Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (cannot open)" & vbCrLf
            GoTo NextKey
        End If

        If wbTarget.ReadOnly Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (file is locked / Read-Only)" & vbCrLf
            wbTarget.Close SaveChanges:=False
            Set wbTarget = Nothing
            GoTo NextKey
        End If

        Set wsTarget = Nothing
        On Error Resume Next
        Set wsTarget = wbTarget.Sheets("Data")
        On Error GoTo FileError

        If wsTarget Is Nothing Then
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & keyName & " (missing 'Data' sheet)" & vbCrLf
            wbTarget.Close SaveChanges:=False
            Set wbTarget = Nothing
            GoTo NextKey
        End If

        ' ==================== Smart clear up to section column ====================
        ClearDataArea wsTarget, HeaderRow:=HEADER_ROW, SectionHeader:=SectionHeaderText()

        PasteRow = PASTE_START_ROW
        FileCopiedRows = 0
        FileSections = 0

        For i = LBound(TargetKeys) To UBound(TargetKeys)
            TextH1 = NormalizeKey(CStr(TargetKeys(i)))

            If Not SheetIndex.Exists(TextH1) Then GoTo NextSection

            Set wsGlobal = SheetIndex(TextH1)
            If Not UsedSheets.Exists(TextH1) Then UsedSheets.Add TextH1, True

            LastCol = wsGlobal.Cells(2, wsGlobal.Columns.Count).End(xlToLeft).Column
            LastRow = LastUsedRow(wsGlobal)
            RowsCopied = LastRow - 2

            If RowsCopied <= 0 Then GoTo NextSection

            SectionNumber = Right(TextH1, 2)

            ' ==================== Guard 1-row x 1-column case ====================
            If RowsCopied = 1 And LastCol = 1 Then
                wsTarget.Cells(PasteRow, 1).Value = wsGlobal.Cells(3, 1).Value
                wsTarget.Cells(PasteRow, LastCol + 1).Value = SectionNumber
            Else
                DataArr = wsGlobal.Range( _
                              wsGlobal.Cells(3, 1), _
                              wsGlobal.Cells(LastRow, LastCol)).Value

                wsTarget.Cells(PasteRow, 1) _
                        .Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

                wsTarget.Cells(PasteRow, LastCol + 1) _
                        .Resize(UBound(DataArr, 1), 1).Value = SectionNumber
            End If

            PasteRow = PasteRow + RowsCopied
            FileCopiedRows = FileCopiedRows + RowsCopied
            FileSections = FileSections + 1

NextSection:
        Next i

        ' ==================== Protected save ====================
        If FileCopiedRows > 0 Then
            On Error Resume Next
            wbTarget.Save
            saveErr = Err.Number: saveDesc = Err.Description
            Err.Clear
            On Error GoTo FileError

            If saveErr <> 0 Then
                SkippedCount = SkippedCount + 1
                SkippedFiles = SkippedFiles & "  - " & keyName & _
                               " (save failed: " & saveDesc & ")" & vbCrLf
            Else
                SuccessCount = SuccessCount + FileSections
            End If
        End If

        wbTarget.Close SaveChanges:=False
        Set wbTarget = Nothing
        GoTo NextKey

FileError:
        ' Any error inside one file: record it, leave that file unsaved, continue with the next
        SkippedCount = SkippedCount + 1
        SkippedFiles = SkippedFiles & "  - " & keyName & " (error " & Err.Number & ": " & _
                       Err.Description & " - file not saved)" & vbCrLf
        Resume NextKey

NextKey:
        If Not wbTarget Is Nothing Then
            On Error Resume Next
            wbTarget.Close SaveChanges:=False
            Err.Clear
        End If
        On Error GoTo UnexpectedError
        Set wbTarget = Nothing
        Set wsTarget = Nothing
    Next Key

    ' Check for unmapped sheets in the in-memory results
    For Each kSheet In SheetIndex.Keys
        If Not UsedSheets.Exists(CStr(kSheet)) Then
            UnmappedCount = UnmappedCount + 1
            UnmappedSections = UnmappedSections & "  - " & CStr(kSheet) & vbCrLf
        End If
    Next kSheet

    wbGlobal.Close SaveChanges:=False
    Set wbGlobal = Nothing

    ' ==================== Final report ====================
    Report = "Distribution completed." & vbCrLf & vbCrLf & _
             "--- Statistics ---" & vbCrLf & _
             "Source files read from " & SOURCE_FOLDER_NAME & " : " & SrcCopied & vbCrLf & _
             "Successful paste operations : " & SuccessCount & vbCrLf & _
             "Skipped files : " & SkippedCount & vbCrLf & _
             "Missing sections : " & MissingCount & vbCrLf & _
             "Unmapped source sheets : " & UnmappedCount & vbCrLf & _
             "Unindexed source sheets : " & UnindexedCount & vbCrLf & _
             "Converted to .xlsx : " & ConvOK & " of " & ConvTotal & " (failed " & ConvFailed & ")" & vbCrLf & _
             "Unreadable source files : " & SrcSkipped

    If backupReady Then
        Report = Report & vbCrLf & "Backup : " & BACKUP_FOLDER & "\" & RunStamp
    End If

    If SkippedCount > 0 Then
        Report = Report & vbCrLf & vbCrLf & "Skipped files:" & vbCrLf & SkippedFiles
    End If

    If MissingCount > 0 Then
        Report = Report & vbCrLf & vbCrLf & _
                 "Missing sections in " & SOURCE_FOLDER_NAME & ":" & vbCrLf & MissingSections
    End If

    If UnmappedCount > 0 Then
        Report = Report & vbCrLf & vbCrLf & _
                 "Unmapped sections in " & SOURCE_FOLDER_NAME & ":" & vbCrLf & UnmappedSections
    End If

    If UnindexedCount > 0 Then
        Report = Report & vbCrLf & vbCrLf & _
                 "Unindexed sheets in " & SOURCE_FOLDER_NAME & " (not usable):" & vbCrLf & UnindexedSheets
    End If

    If SrcSkipped > 0 Then
        Report = Report & vbCrLf & vbCrLf & _
                 "Unreadable / skipped source files:" & vbCrLf & SrcList
    End If

    MsgBox Report, vbInformation, "Distribution Report"

CleanExit:
    ' ==================== Restore application settings ====================
    Application.AutomationSecurity = prevSec
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Exit Sub

UnexpectedError:
    errMsg = "An unexpected error occurred during execution:" & vbCrLf & _
             "Error " & Err.Number & ": " & Err.Description

    On Error Resume Next
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False
    If Not wbGlobal Is Nothing Then wbGlobal.Close SaveChanges:=False
    On Error GoTo 0

    MsgBox errMsg, vbCritical, "Unexpected Error"
    Resume CleanExit
End Sub

' ------------------------------------------------------------------------------
' Clears the data area from row (HeaderRow + 1) to the last row,
' from column 1 up to the section-header column only.
' ------------------------------------------------------------------------------
Private Sub ClearDataArea(ByVal ws As Worksheet, _
                          ByVal HeaderRow As Long, _
                          ByVal SectionHeader As String)

    Dim colIndex    As Long
    Dim LastRow     As Long
    Dim LastRowSec  As Long
    Dim cell        As Range
    Dim lo          As ListObject

    ' Show all rows of any active filter (ClearContents skips hidden rows).
    ' The filter itself is kept, only its criteria are reset.
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    For Each lo In ws.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
        End If
    Next lo
    Err.Clear
    On Error GoTo 0

    ' Scan the header row to locate the section column
    colIndex = 0
    For Each cell In ws.Rows(HeaderRow).Cells
        If StrComp(Trim(CStr(cell.Value)), SectionHeader, vbTextCompare) = 0 Then
            colIndex = cell.Column
            Exit For
        End If
        If cell.Column > 100 Then Exit For
    Next cell

    ' Fallback: section header not found -> locate last used column in header row
    If colIndex = 0 Then
        colIndex = ws.Cells(HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    End If

    ' Compute the last row by comparing column A with the section column
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If colIndex > 1 Then
        LastRowSec = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
        If LastRowSec > LastRow Then LastRow = LastRowSec
    End If

    ' Clear from the row after the header to the last row
    If LastRow > HeaderRow Then
        ws.Range(ws.Cells(HeaderRow + 1, 1), _
                 ws.Cells(LastRow, colIndex)).ClearContents
    End If
End Sub

' ==============================================================================
' SOURCE01 BUILD - Collector01 logic embedded, so this module needs no other module
' STEP 1: silent conversion of the \Source01\ files to .xlsx (web-exported .xls /
'         .xlsb / .xlsm; originals are moved to \Source01\_Original, never deleted)
' STEP 2: collection into ONE new workbook returned OPEN and UNSAVED (nothing is
'         written to Term01Results.xlsx). The caller closes it with SaveChanges:=False.
' Returns False when nothing usable was collected (reasons in skippedList).
' ==============================================================================
Private Function SrcBuildWorkbook(ByVal srcFolder As String, _
                                  ByRef wbOut As Workbook, _
                                  ByRef copied As Long, _
                                  ByRef skipped As Long, _
                                  ByRef skippedList As String, _
                                  ByRef convTotal As Long, _
                                  ByRef convSuccess As Long, _
                                  ByRef convFailed As Long) As Boolean
    Dim wb As Workbook, oldSec As Long, oldAlerts As Boolean, oldCalc As Long

    SrcBuildWorkbook = False
    Set wbOut = Nothing
    copied = 0: skipped = 0: skippedList = ""
    convTotal = 0: convSuccess = 0: convFailed = 0

    If LCase(Left(srcFolder, 4)) = "http" Then
        skippedList = "  - The workbook is on OneDrive/SharePoint (URL path). Move it to a local folder first." & vbCrLf
        Exit Function
    End If
    If Right(srcFolder, 1) <> "\" Then srcFolder = srcFolder & "\"
    If Dir(srcFolder, vbDirectory) = "" Then
        skippedList = "  - Source01 folder not found: " & srcFolder & vbCrLf
        Exit Function
    End If

    oldSec = Application.AutomationSecurity
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable   ' block Auto_Open in .xlsm/.xls
    On Error GoTo EH

    ' 1) Silent auto-conversion inside \Source01\ first
    ConvertSourceFilesToXLSX srcFolder, convTotal, convSuccess, convFailed

    ' 2) Collect all files into a new (unsaved) workbook
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
    Application.Calculation = oldCalc
    Exit Function

EH:
    skippedList = skippedList & "  - Build failed: " & Err.Number & " : " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wbOut = Nothing
    SrcBuildWorkbook = False
    Resume Done
End Function

' ==============================================================================
' HELPERS
' ==============================================================================
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
' STEP 1: SILENT CONVERSION INSIDE \Source01\ FOLDER
' ==============================================================================
Private Sub ConvertSourceFilesToXLSX(ByVal targetFolderPath As String, _
                                     ByRef totalFiles As Long, _
                                     ByRef convertedFiles As Long, _
                                     ByRef failedFiles As Long)
    Dim fso As Object, folderObj As Object, fileObj As Object
    Dim paths As Collection, p As Variant
    Dim wb As Workbook
    Dim newFilePath As String, baseName As String, ext As String
    Dim cleanFolder As String, archPath As String, archFile As String
    Dim saveErr As Long

    totalFiles = 0: convertedFiles = 0: failedFiles = 0

    If Right(targetFolderPath, 1) = "\" Then
        cleanFolder = Left(targetFolderPath, Len(targetFolderPath) - 1)
    Else
        cleanFolder = targetFolderPath
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(cleanFolder) Then Exit Sub
    Set folderObj = fso.GetFolder(cleanFolder)
    archPath = cleanFolder & "\" & ARCHIVE_FOLDER

    ' Snapshot the file list first (do not modify the folder while enumerating it)
    Set paths = New Collection
    For Each fileObj In folderObj.Files
        If Left(fileObj.Name, 2) <> "~$" Then
            ext = LCase(fso.GetExtensionName(fileObj.Name))
            If ext = "xls" Or ext = "xlsb" Or ext = "xlsm" Then
                If StrComp(fileObj.Path, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
                    paths.Add fileObj.Path
                End If
            End If
        End If
    Next fileObj

    For Each p In paths
        totalFiles = totalFiles + 1
        baseName = fso.GetBaseName(CStr(p))
        newFilePath = cleanFolder & "\" & baseName & ".xlsx"

        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(fileName:=CStr(p), UpdateLinks:=False, ReadOnly:=True, _
                                AddToMRU:=False, Password:="")
        Err.Clear
        On Error GoTo 0

        If wb Is Nothing Then
            failedFiles = failedFiles + 1
        ElseIf fso.FileExists(newFilePath) Then
            ' Target already exists: do not overwrite (duplicate is skipped later in collection)
            wb.Close SaveChanges:=False
            failedFiles = failedFiles + 1
        Else
            On Error Resume Next
            wb.SaveAs fileName:=newFilePath, FileFormat:=xlOpenXMLWorkbook
            saveErr = Err.Number: Err.Clear
            wb.Close SaveChanges:=False
            Err.Clear
            On Error GoTo 0

            If saveErr = 0 Then
                ' Archive the original instead of deleting it
                On Error Resume Next
                If Not fso.FolderExists(archPath) Then fso.CreateFolder archPath
                archFile = archPath & "\" & fso.GetFileName(CStr(p))
                If fso.FileExists(archFile) Then
                    archFile = archPath & "\" & Format(Now, "yyyymmdd_hhnnss") & "_" & fso.GetFileName(CStr(p))
                End If
                fso.MoveFile CStr(p), archFile
                Err.Clear
                On Error GoTo 0
                convertedFiles = convertedFiles + 1
            Else
                failedFiles = failedFiles + 1
            End If
        End If
        Set wb = Nothing
    Next p
End Sub

' ==============================================================================
' STEP 2: FONT CHECK & COLLECTION INTO Term01Results.xlsx
' ==============================================================================
' Checks the Windows font registry (HKLM + HKCU). If the check itself cannot run,
' returns True (keeps the requested font instead of forcing the fallback).
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
                    If p > 0 Then nm = Left(nm, p - 1)          ' strip "(TrueType)" / "(OpenType)"
                    ' exact name, or family name of a variable font (e.g. "Bahnschrift" for "Bahnschrift SemiBold")
                    If StrComp(nm, FontName, vbTextCompare) = 0 Then FontExists = True: Exit Function
                    If StrComp(Left(FontName, Len(nm) + 1), nm & " ", vbTextCompare) = 0 Then FontExists = True: Exit Function
                Next i
            End If
        End If
    Next hv

    FontExists = Not enumOK     ' enumeration worked and nothing matched -> False; enumeration failed -> True
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

    ' ---- Read source data
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

    lastDataRow = 1 + UBound(DataArr, 1)     ' header at row 2 + data rows

    ' ---- Number formats BEFORE writing data (so "@" columns store text)
    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        If LastCol >= 4 Then .Columns("D").NumberFormat = "@"
        If LastCol >= 5 Then .Columns("E").NumberFormat = "@"
        If LastCol >= 6 Then .Range(.Columns("F"), .Columns(LastCol)).NumberFormat = "0.00"
    End With

    wsResult.Range("A2").Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

    ' ---- Unique, non-empty headers (case-insensitive, like Excel tables)
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

    ' ---- Dynamic table (range ends exactly at the last data row)
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

    ' ---- Formatting
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

        ' Identity columns: AutoFit. Score columns: fixed width, rotated bold headers.
        idLastCol = IDENTITY_LAST_COL
        If LastCol < idLastCol Then idLastCol = LastCol
        .Range(.Columns(1), .Columns(idLastCol)).EntireColumn.AutoFit

        If LastCol >= 6 Then .Range(.Columns(IDENTITY_LAST_COL + 1), .Columns(LastCol)).ColumnWidth = COL_WIDTH

        If LastCol >= 4 Then
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Orientation = xlUpward
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Font.Bold = True
        End If
    End With

    ' Apply title formatting last so table column fonts do not override it.
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

    ' Collect names first (Dir cannot be nested with other Dir calls)
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
        MsgBox "Source01 folder is empty or contains no Excel files:" & vbCrLf & SourcePath, vbExclamation, "Empty Source01 Folder"
        Exit Sub
    End If

    sorted = SortedNames(names)     ' deterministic sheet order (01, 02, ...)
    SetupFonts

    For i = LBound(sorted) To UBound(sorted)
        f = CStr(sorted(i))
        dotPos = InStrRev(f, ".")
        ext = LCase(Mid(f, dotPos + 1))
        baseName = Left(f, dotPos - 1)

        If StrComp(SourcePath & f, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            ' skip the macro workbook itself
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

    If CopiedCount = 0 Then
        MsgBox "None of the source files could be processed.", vbCritical, "Collection Failed"
    End If
End Sub


Private Sub PlaceHeaderTitle(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet)

    Dim rawText As String
    Dim partTitle As String
    Dim partGroup As String

    ' 1. Capture and clean the text from A5
    rawText = CleanText(CStr(wsSource.Range("A5").Value))

    ' 2. Split at the end of the school year
    SplitAtSchoolYear rawText, partTitle, partGroup


    ' 3. Write and format both merged ranges
    WriteMergedCell wsTarget.Range("B1:F1"), partTitle
    WriteMergedCell wsTarget.Range("H1:Q1"), partGroup

    ' Merged cells do not auto-fit, so set the row height explicitly
    wsTarget.Rows(1).RowHeight = HEADER_ROW_HEIGHT

End Sub

'----------------------------------------------------------
' Cleans text: line breaks, non-breaking spaces, control
' characters and repeated spaces.
'----------------------------------------------------------
Private Function CleanText(ByVal s As String) As String

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(160), " ")                       ' non-breaking space
    s = Application.WorksheetFunction.Clean(s)          ' control characters
    CleanText = Application.WorksheetFunction.Trim(s)   ' extra spaces

End Function

'----------------------------------------------------------
' Splits text into:
'   part1 = everything up to the end of the school year
'   part2 = the remainder (class / group name)
' If no year is found, part1 = full text and part2 = "".
'----------------------------------------------------------
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

    ' Western (0-9) and Arabic-Indic (U+0660-U+0669) digits
    digit = "[0-9\u0660-\u0669]"
    ' Separators: slash, backslash, hyphen, en dash, em dash
    seps = "[/\\\-\u2013\u2014]"

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True

    ' First choice: full school-year pattern (e.g. 2025-2026)
    rx.Pattern = digit & "{4}\s*" & seps & "\s*" & digit & "{4}"

    If Not rx.Test(txt) Then
        ' Fallback: a single 4-digit year
        rx.Pattern = digit & "{4}"
        If Not rx.Test(txt) Then Exit Sub
    End If

    Set m = rx.Execute(txt)
    splitPos = m(0).FirstIndex + m(0).Length

    part1 = Trim(Left(txt, splitPos))

    If splitPos < Len(txt) Then
        part2 = Trim(Mid(txt, splitPos + 1))

        ' Remove any separator left at the start of the group name
        Do While Len(part2) > 0
            If InStr(1, "-:/" & ChrW(&H2013) & ChrW(&H2014), Left(part2, 1)) = 0 Then Exit Do
            part2 = Trim(Mid(part2, 2))
        Loop
    End If

End Sub

'----------------------------------------------------------
' Merges the range, writes the text and applies formatting.
'----------------------------------------------------------
Private Sub WriteMergedCell(ByVal rng As Range, ByVal txt As String)

    With rng
        .UnMerge                                ' remove any previous merge
        .ClearContents                          ' clear leftovers
        .Merge                                  ' merge the range
        .Value = txt
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL                   ' right-to-left reading order
        .Font.Name = GetAvailableFont(PREFERRED_FONT, FALLBACK_FONT)
        .Font.Size = HEADER_FONT_SIZE
        .Font.Bold = True
    End With

End Sub

'----------------------------------------------------------
' Returns the preferred font if installed, otherwise the fallback.
'----------------------------------------------------------
Private Function GetAvailableFont(ByVal wanted As String, _
                                  ByVal fallback As String) As String

    If IsFontInstalled(wanted) Then
        GetAvailableFont = wanted
    Else
        GetAvailableFont = fallback
    End If

End Function

'----------------------------------------------------------
' Checks the system and per-user font folders for a file whose
' name starts with the font name (e.g. Cairo-Regular.ttf).
'----------------------------------------------------------
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
