Attribute VB_Name = "modClearDataOnly01"
Option Explicit

'==============================================================================
' Module  : modClearDataOnly01
' Purpose : Clears ONLY the data area (from row 7 downwards, up to the section
'           column) in the TARGET workbooks (Suffix 01).
'           Does NOT paste any new data.
' Entry   : ClearAllTargetFilesData01()
' Note    : STANDALONE module. It intentionally shares no code and no names
'           with the other semester modules (each one runs from its own folder
'           and at its own time). Do not merge them.
'==============================================================================

Public Sub ClearAllTargetFilesData01()

    Dim SourcePath    As String
    Dim Map           As Object
    Dim Key           As Variant

    Dim ClearedCount  As Long
    Dim SkippedCount  As Long
    Dim SkippedFiles  As String
    Dim SectionText   As String
    Dim Note          As String
    Dim Report        As String
    Dim prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean
    Dim prevSec As MsoAutomationSecurity
    Dim fso As Object, backupDir As String, backupRoot As String
    Dim runName As String, suffix As Long

    ' Define the Arabic word for "Section" using ChrW to avoid encoding issues
    SectionText = ChrW(1575) & ChrW(1604) & ChrW(1602) & ChrW(1587) & ChrW(1605)  ' Arabic word for "Section"

    ' ==================== Folder check ====================
    SourcePath = ThisWorkbook.Path

    If Len(SourcePath) = 0 Then
        MsgBox "This workbook has not been saved yet." & vbCrLf & _
               "Save it in the folder that contains the target files, then run again.", _
               vbExclamation, "Clear Data Report (01)"
        Exit Sub
    End If

    If InStr(1, SourcePath, "://") > 0 Then
        MsgBox "This workbook is opened from a cloud/web location:" & vbCrLf & _
               SourcePath & vbCrLf & vbCrLf & _
               "Run it from a normal local folder (outside OneDrive/SharePoint) " & _
               "that contains the target files.", _
               vbExclamation, "Clear Data Report (01)"
        Exit Sub
    End If

    If Right$(SourcePath, 1) <> Application.PathSeparator Then
        SourcePath = SourcePath & Application.PathSeparator
    End If

    ' ==================== Target file map (Semester 01) ====================
    Set Map = CreateObject("Scripting.Dictionary")
    Map.CompareMode = vbTextCompare

    Map.Add "1LitAnalysis01.xlsm", True
    Map.Add "1TecnoGlobalAnalysis01.xlsm", True
    Map.Add "2GEAnalysis01.xlsm", True
    Map.Add "2LEAnalysis01.xlsm", True
    Map.Add "2MathAnalysis01.xlsm", True
    Map.Add "2PhiloAnalysis01.xlsm", True
    Map.Add "2SEGlobalAnalysis01.xlsm", True
    Map.Add "3GEAnalysis01.xlsm", True
    Map.Add "3LEAnalysis01.xlsm", True
    Map.Add "3MathAnalysis01.xlsm", True
    Map.Add "3PhiloAnalysis01.xlsm", True
    Map.Add "3SEGlobalAnalysis01.xlsm", True

    If MsgBox("Clear all target tables for Term 01?" & vbCrLf & _
              "The original target files will be saved empty. A backup is required for each file.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Clear Tables") <> vbYes Then Exit Sub

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevSec = Application.AutomationSecurity

    On Error GoTo UnexpectedError

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set fso = CreateObject("Scripting.FileSystemObject")
    backupRoot = SourcePath & "_Backup"
    If Not fso.FolderExists(backupRoot) Then fso.CreateFolder backupRoot
    runName = "Clear01_" & Format(Now, "yyyymmdd_hhnnss")
    backupDir = backupRoot & "\" & runName
    Do While fso.FolderExists(backupDir)
        suffix = suffix + 1
        backupDir = backupRoot & "\" & runName & "_" & CStr(suffix)
    Loop
    fso.CreateFolder backupDir

    ' ==================== Main loop over target files ====================
    ' Every file is processed in isolation: a problem in one file is reported
    ' and the loop continues with the next one.
    For Each Key In Map.Keys

        Note = vbNullString

        If ProcessOneFile(SourcePath, CStr(Key), SectionText, backupDir, Note) Then
            ClearedCount = ClearedCount + 1
        Else
            SkippedCount = SkippedCount + 1
            SkippedFiles = SkippedFiles & "  - " & CStr(Key) & " (" & Note & ")" & vbCrLf
        End If

    Next Key

    ' ==================== Final report ====================
    Report = "Data clearing completed (Semester 01)." & vbCrLf & vbCrLf & _
             "Target files cleared successfully: " & ClearedCount & vbCrLf & _
             "Skipped files: " & SkippedCount & vbCrLf & "Backup: " & backupDir

    If SkippedCount > 0 Then
        Report = Report & vbCrLf & vbCrLf & "Skipped files details:" & vbCrLf & SkippedFiles
    End If

    MsgBox Report, vbInformation, "Clear Data Report (01)"
    GoTo CleanExit

UnexpectedError:
    MsgBox "An unexpected error occurred during clearing:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"

CleanExit:
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' Processes ONE target workbook. Returns True only if the data was cleared AND
' the workbook was saved. Otherwise returns False and puts the reason in Note.
' The workbook is never saved unless every step succeeded.
' ------------------------------------------------------------------------------
Private Function ProcessOneFile(ByVal FolderPath As String, _
                                ByVal FileName As String, _
                                ByVal SectionHeader As String, _
                                ByVal BackupDir As String, _
                                ByRef Note As String) As Boolean

    Dim wb      As Workbook
    Dim ws      As Worksheet
    Dim Stage   As String

    ProcessOneFile = False
    On Error GoTo FileError

    Stage = "locating the file"
    If Len(Dir(FolderPath & FileName, vbReadOnly + vbHidden + vbSystem)) = 0 Then
        Note = "file not found"
        Exit Function
    End If

    ' A workbook that is already open is never touched (it may hold unsaved work)
    Stage = "checking open workbooks"
    If IsWorkbookOpen(FolderPath & FileName) Then
        Note = "already open - close it and run again"
        Exit Function
    End If

    Stage = "backing up"
    FileCopy FolderPath & FileName, BackupDir & "\" & FileName

    Stage = "opening"
    Set wb = Workbooks.Open(FolderPath & FileName, _
                            UpdateLinks:=False, _
                            ReadOnly:=False, _
                            AddToMRU:=False, Password:="")

    ' Opened read-only (e.g. locked by another user): do not modify it
    If wb.ReadOnly Then
        Note = "opened read-only (in use?) - not modified"
        wb.Close SaveChanges:=False
        Exit Function
    End If

    Stage = "looking for the Data sheet"
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Data")
    On Error GoTo FileError

    If ws Is Nothing Then
        Note = "missing 'Data' sheet"
        wb.Close SaveChanges:=False
        Exit Function
    End If

    ' ==================== Clear Data ====================
    Stage = "clearing data"
    ClearDataArea ws, HeaderRow:=6, SectionHeader:=SectionHeader

    ' Recalculate only worksheets belonging to this target workbook.
    Stage = "recalculating"
    Dim calcSheet As Worksheet
    For Each calcSheet In wb.Worksheets
        calcSheet.Calculate
    Next calcSheet

    ' ==================== Protected save ====================
    Stage = "saving"
    wb.Save

    Stage = "closing"
    wb.Close SaveChanges:=False

    ProcessOneFile = True
    Exit Function

FileError:
    Note = "failed while " & Stage & " - error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ProcessOneFile = False
End Function

' ------------------------------------------------------------------------------
' True if a workbook with this full path is already open in this Excel session
' ------------------------------------------------------------------------------
Private Function IsWorkbookOpen(ByVal FullPath As String) As Boolean
    Dim wbOpen As Workbook
    For Each wbOpen In Application.Workbooks
        If StrComp(wbOpen.FullName, FullPath, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wbOpen
End Function

' ------------------------------------------------------------------------------
' Clearing logic of this module (self-contained, deliberately not shared)
' ------------------------------------------------------------------------------
Private Sub ClearDataArea(ByVal ws As Worksheet, _
                          ByVal HeaderRow As Long, _
                          ByVal SectionHeader As String)

    Dim colIndex As Long, lastHeaderCol As Long, c As Long
    Dim lastRow As Long, hdr As Variant
    Dim lo As ListObject, dataRange As Range, lastCell As Range

    ' Reset worksheet and table filter criteria without removing filter buttons.
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    For Each lo In ws.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
        End If
    Next lo
    Err.Clear
    On Error GoTo 0

    lastHeaderCol = ws.Cells(HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastHeaderCol
        hdr = ws.Cells(HeaderRow, c).Value
        If Not IsError(hdr) Then
            If StrComp(NormalizeHeader(CStr(hdr)), SectionHeader, vbTextCompare) = 0 Then
                If colIndex <> 0 Then Err.Raise vbObjectError + 701, "ClearDataArea", "Duplicate section headers; file not modified."
                colIndex = c
            End If
        End If
    Next c
    If colIndex = 0 Then Err.Raise vbObjectError + 702, "ClearDataArea", "Section header not found; file not modified."

    Set dataRange = ws.Range(ws.Cells(HeaderRow + 1, 1), ws.Cells(ws.Rows.Count, colIndex))
    Set lastCell = dataRange.Find(What:="*", After:=dataRange.Cells(1, 1), _
                                 LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not lastCell Is Nothing Then
        lastRow = lastCell.Row
        ws.Range(ws.Cells(HeaderRow + 1, 1), ws.Cells(lastRow, colIndex)).ClearContents
    End If
End Sub

Private Function NormalizeHeader(ByVal txt As String) As String
    Dim code As Variant
    txt = Replace(txt, ChrW(160), " ")
    For Each code In Array(&H200B, &H200E, &H200F, &H202A, &H202B, &H202C, &H640)
        txt = Replace(txt, ChrW(CLng(code)), "")
    Next code
    NormalizeHeader = Trim$(txt)
End Function
