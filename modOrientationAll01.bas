Attribute VB_Name = "modOrientationAll01"
Option Explicit

'==============================================================================
' Module  : modOrientationAll01
' Purpose : ONE entry (Run_OrientationAll01) for the whole school-orientation block
'           of Term 01. It is the action of the button "Term 01 Orientation".
'
'   Run_OrientationAll01
'    0. checks that every input file needed by the steps exists, in the order of the steps
'    1. modTransfer01            Data Transfer       Source01 -> guidance workbook BUILT IN MEMORY (not saved)
'    2. modStreamPreferences01   Stream Preferences  wishes file in Source01 (found by structure) -> wishes (still in memory;
'                                                    if it stops, a copy is saved)
'    3. modOrientation01         Orientation         decisions + general report, then the workbook is SAVED
'                                                    -> \معالجة بيانات الفصل 01\Term01Guidance[_NN].xlsx
'    4. modOrientationResults01  Export results      -> \معالجة بيانات الفصل 01\نتائج التوجيه التدريجي للفصل 01[_NN].xlsx
'
' Runs all steps silently and shows ONE consolidated message at the end
' (or ONE diagnostic message if a step stops).
' Handles user interruption (Esc / Ctrl+Break) cleanly so the button never freezes.
'==============================================================================

Private Const TITLE_TXT As String = "Term 01 Orientation"
Private Const SRC_FOLDER As String = "Source01"

Private m_busy As Boolean

Public Sub ResetOrientationState()
    m_busy = False
    On Error Resume Next
    modTransfer01.SetSilentMode False
    modStreamPreferences01.SetSilentMode False
    modOrientation01.SetSilentMode False
    modOrientationResults01.SetSilentMode False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo 0
End Sub

Public Sub Run_OrientationAll01()
    Dim base As String, problem As String, t0 As Date, outPath As String, secs As Long
    Dim finalMsg As String

    ' Always clear any stale lock from a previously interrupted session
    m_busy = False

    base = ThisWorkbook.Path
    If Len(base) = 0 Then
        MsgBox "Save the tool workbook first: the files are looked for next to it.", vbExclamation, TITLE_TXT
        Exit Sub
    End If
    If InStr(1, base, "://") > 0 Then
        MsgBox "The tool is opened from a cloud/web location:" & vbCrLf & base & vbCrLf & vbCrLf & _
               "Run it from a normal local folder that contains the files.", vbExclamation, TITLE_TXT
        Exit Sub
    End If
    If Right$(base, 1) <> Application.PathSeparator Then base = base & Application.PathSeparator

    ' ---- 0. every file the four steps need, in the order of the steps
    problem = CheckFiles(base)
    If Len(problem) > 0 Then
        MsgBox problem, vbExclamation, TITLE_TXT
        Exit Sub
    End If

    m_busy = True
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    t0 = Now

    ' ---- 1. data transfer (silent)
    modTransfer01.SetSilentMode True
    modTransfer01.Run_Transfer01
    If Not modTransfer01.Transfer01Succeeded() Then
        StopAt 1, "Data Transfer", modTransfer01.LastReport()
        GoTo Finished
    End If

    ' ---- 2. stream preferences / wishes (silent)
    modStreamPreferences01.SetSilentMode True
    modStreamPreferences01.Run_StreamPreferences01
    If Not modStreamPreferences01.Preferences01Succeeded() Then
        StopAt 2, "Stream Preferences", modStreamPreferences01.LastReport()
        GoTo Finished
    End If

    ' ---- 3. orientation (silent)
    modOrientation01.SetSilentMode True
    modOrientation01.Run_Orientation01
    If Not modOrientation01.Orientation01Succeeded() Then
        StopAt 3, "Orientation", modOrientation01.LastReport()
        GoTo Finished
    End If

    ' ---- 4. export results (silent)
    modOrientationResults01.SetSilentMode True
    modOrientationResults01.Run_OrientationResults01
    outPath = modOrientationResults01.Results01OutputPath()
    If Not modOrientationResults01.OrientationResults01Succeeded() Or Len(outPath) = 0 Or Not FileThere(outPath) Then
        StopAt 4, "Export Orientation Results", modOrientationResults01.LastReport()
        GoTo Finished
    End If

    secs = CLng((Now - t0) * 86400)
    finalMsg = "All 4 steps completed (" & secs & " s):" & vbCrLf & vbCrLf & _
               "[1] Data Transfer:" & vbCrLf & modTransfer01.LastReport() & vbCrLf & vbCrLf & _
               "[2] Stream Preferences:" & vbCrLf & modStreamPreferences01.LastReport() & vbCrLf & vbCrLf & _
               "[3] Orientation:" & vbCrLf & modOrientation01.LastReport() & vbCrLf & vbCrLf & _
               "[4] Export Results:" & vbCrLf & modOrientationResults01.LastReport()

    MsgBox finalMsg, vbInformation, TITLE_TXT

Finished:
    On Error Resume Next
    ThisWorkbook.Activate   ' the new guidance workbook is left open (already saved); return to the tool
    m_busy = False
    modTransfer01.SetSilentMode False
    modStreamPreferences01.SetSilentMode False
    modOrientation01.SetSilentMode False
    modOrientationResults01.SetSilentMode False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo 0
    Exit Sub

Failed:
    If Err.Number = 18 Then
        MsgBox "Execution was stopped by the user. Ready to run again.", vbExclamation, TITLE_TXT
    Else
        MsgBox "The orientation sequence stopped: " & Err.Number & " : " & Err.Description, vbCritical, TITLE_TXT
    End If
    Resume Finished
End Sub

Private Sub StopAt(ByVal stepNo As Long, ByVal stepName As String, ByVal stepReport As String)
    Dim msg As String
    msg = "The sequence stopped at step " & stepNo & " (" & stepName & ")." & vbCrLf & _
          "The steps after it were not run."
    If Len(Trim$(stepReport)) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Details:" & vbCrLf & stepReport
    End If
    MsgBox msg, vbExclamation, TITLE_TXT
End Sub

' Returns "" when everything is there, otherwise the text for the FIRST missing item.
Private Function CheckFiles(ByVal base As String) As String
    Dim problem As String, wishPath As String
    If Not FolderHasExcelFiles(base & SRC_FOLDER) Then
        CheckFiles = "Step 1 (Data Transfer) needs the folder '" & SRC_FOLDER & "' with at least one Excel file in it:" & _
                     vbCrLf & base & SRC_FOLDER & vbCrLf & vbCrLf & "Put the term result files there, then run again."
        Exit Function
    End If
    problem = modStreamPreferences01.FindWishesFile(wishPath)
    If Len(problem) > 0 Then
        CheckFiles = "Step 2 (Stream Preferences) cannot start:" & vbCrLf & problem & vbCrLf & vbCrLf & _
                     "Nothing has been changed."
        Exit Function
    End If
End Function

Private Function FileThere(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileThere = fso.FileExists(p)
    If Err.Number <> 0 Then FileThere = False: Err.Clear
End Function

' At least one .xls / .xlsx / .xlsm / .xlsb file that is not an Excel lock file (~$...)
Private Function FolderHasExcelFiles(ByVal p As String) As Boolean
    Dim fso As Object, f As Object, ext As String
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then Exit Function
    For Each f In fso.GetFolder(p).Files
        If Left$(f.Name, 2) <> "~$" Then
            ext = LCase$(fso.GetExtensionName(f.Name))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Then
                FolderHasExcelFiles = True
                Exit Function
            End If
        End If
    Next f
Done:
    Err.Clear
End Function