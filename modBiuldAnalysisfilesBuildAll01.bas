Attribute VB_Name = "modBiuldAnalysisfilesBuildAll01"
Option Explicit

'==============================================================================
' Module  : modBiuldAnalysisfilesBuildAll01
' Purpose : ONE macro (Build_All) that creates all the Term 01 analysis files.
'
'   Build_All
'    |- modBiuldAnalysisfilesData01       -> Data (collects \Source01\ + model columns)
'    |- modBiuldAnalysisfilesAnalysis01   -> Analysis01 (term 1, one block per section)
'    |- modBiuldAnalysisfilesCompaire01   -> AnalyseCompaire (only files with several sections)
'    |- modBiuldAnalysisfilesCartTec01    -> Cart_Tec (student card)
'==============================================================================

Public Sub Build_All()
    RunBuild ""
End Sub

Public Sub Build_One()
    Dim s As String
    s = InputBox("File to build (stem of the name), for example: 3SEGlobalAnalysis01", "Build_One", "3SEGlobalAnalysis01")
    If Len(Trim$(s)) = 0 Then Exit Sub
    RunBuild Trim$(s)
End Sub

Private Sub RunBuild(ByVal onlyStem As String)
    Dim wbGlobal As Workbook
    Dim sheetIndex As Object, dupKeys As Object, map As Object
    Dim key As Variant, stem As String, titles As Variant
    Dim srcReport As String, why As String, errMsg As String, savedPath As String
    Dim outDir As String, logText As String, failBrief As String, logPath As String, msg As String
    Dim okCount As Long, failCount As Long, total As Long, idx As Long
    Dim oldSU As Boolean, oldEV As Boolean, oldDA As Boolean, oldSB As Variant, oldCalc As Long
    Dim t0 As Double, t1 As Double, secs As Double, secsAll As Double
    Dim found As Boolean, firstErr As String, retried As Boolean

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the tool workbook first: the Source01 folder and the output folder are located next to it.", _
               vbExclamation, "Term 01 Analysis"
        Exit Sub
    End If
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "The tool is on OneDrive/SharePoint (URL path). Move it to a local folder first.", _
               vbExclamation, "Term 01 Analysis"
        Exit Sub
    End If

    oldSU = Application.ScreenUpdating
    oldEV = Application.EnableEvents
    oldDA = Application.DisplayAlerts
    oldSB = Application.StatusBar
    oldCalc = Application.Calculation
    t0 = Timer

    On Error GoTo Fatal
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set map = GetFileMap()
    If onlyStem <> "" Then
        If Not map.Exists(onlyStem) Then
            MsgBox "Unknown file name: " & onlyStem, vbExclamation, "Term 01 Analysis"
            GoTo Done
        End If
        total = 1
    Else
        total = map.Count
    End If

    Panel_Reset map

    Application.StatusBar = "Term 01 Analysis : checking fonts ..."
    EnsureFonts

    Application.StatusBar = "Term 01 Analysis : reading Source01 ..."
    If Not LoadSource(wbGlobal, sheetIndex, dupKeys, srcReport) Then
        Panel_Finish 0, total, 0, "Source01: " & Replace(Trim$(srcReport), vbCrLf, " | ")
        MsgBox "No usable source file was found." & vbCrLf & vbCrLf & Left$(srcReport, 700), vbCritical, "Term 01 Analysis"
        GoTo Done
    End If

    outDir = ThisWorkbook.Path & "\" & OutputFolderName() & "\"
    If Not EnsureFolderP(outDir) Then
        MsgBox "The output folder could not be created:" & vbCrLf & outDir, vbCritical, "Term 01 Analysis"
        GoTo Done
    End If

    logText = "Term 01 Analysis - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
              "Tool   : " & ThisWorkbook.FullName & vbCrLf & _
              "Output : " & outDir & vbCrLf & String$(70, "-") & vbCrLf

    For Each key In map.Keys
        idx = idx + 1
        stem = CStr(key)
        If onlyStem = "" Or StrComp(stem, onlyStem, vbTextCompare) = 0 Then
            found = True
            titles = map(key)
            Application.StatusBar = "Term 01 Analysis : " & idx & " / " & map.Count & "  " & stem
            Panel_Update idx, 1, "", 0
            t1 = Timer

            savedPath = ""
            retried = False
            If GroupReady(sheetIndex, dupKeys, titles, why) Then
                errMsg = BuildOneFile(stem, titles, sheetIndex, outDir, savedPath)
                If errMsg <> "" Then
                    firstErr = errMsg
                    DoEvents
                    savedPath = ""
                    errMsg = BuildOneFile(stem, titles, sheetIndex, outDir, savedPath)
                    If errMsg = "" Then retried = True
                End If
            Else
                errMsg = "SKIPPED - the source sections are not complete:" & vbCrLf & why
            End If

            secs = Timer - t1
            If secs < 0 Then secs = secs + 86400

            If errMsg = "" Then
                okCount = okCount + 1
                If retried Then
                    Panel_Update idx, 2, Mid$(savedPath, InStrRev(savedPath, "\") + 1) & "   [retry]", secs
                    logText = logText & "OK     " & stem & "  ->  " & Mid$(savedPath, InStrRev(savedPath, "\") + 1) & _
                              "  (" & Format$(secs, "0.0") & " s)  after ONE RETRY" & vbCrLf & "       first error: " & _
                              Replace(Trim$(firstErr), vbCrLf, " | ") & vbCrLf
                Else
                    Panel_Update idx, 2, Mid$(savedPath, InStrRev(savedPath, "\") + 1), secs
                    logText = logText & "OK     " & stem & "  ->  " & Mid$(savedPath, InStrRev(savedPath, "\") + 1) & _
                              "  (" & Format$(secs, "0.0") & " s)" & vbCrLf
                End If
            Else
                failCount = failCount + 1
                Panel_Update idx, 3, Replace(Trim$(errMsg), vbCrLf, " | "), secs
                logText = logText & "FAILED " & stem & "  (" & Format$(secs, "0.0") & " s)" & vbCrLf & "       " & _
                          Replace(Trim$(errMsg), vbCrLf, vbCrLf & "       ") & vbCrLf
                failBrief = failBrief & "  - " & stem & " : " & Left$(Replace(Trim$(errMsg), vbCrLf, " "), 90) & vbCrLf
            End If
        End If
    Next key

    secsAll = Timer - t0
    If secsAll < 0 Then secsAll = secsAll + 86400

    If Trim$(srcReport) <> "" Then logText = logText & String$(70, "-") & vbCrLf & "Source notes:" & vbCrLf & srcReport
    If gFontNotes <> "" Then logText = logText & String$(70, "-") & vbCrLf & "Fonts replaced:" & gFontNotes & vbCrLf
    logText = logText & String$(70, "-") & vbCrLf & "Created " & okCount & " / " & total & " in " & Format$(secsAll, "0") & " s" & vbCrLf

    logPath = NextFreePath(outDir, "BuildLog_" & Format$(Now, "yyyymmdd_hhnnss"), ".txt")
    WriteLog logPath, logText

    Panel_Finish okCount, total, secsAll, ""

    msg = "Created : " & okCount & " / " & total & "   (" & Format$(secsAll, "0") & " s)" & vbCrLf & "Folder  : " & outDir
    If failCount > 0 Then msg = msg & vbCrLf & vbCrLf & "NOT created (" & failCount & "):" & vbCrLf & failBrief
    msg = msg & vbCrLf & "Full details: panel sheet + " & Mid$(logPath, InStrRev(logPath, "\") + 1)
    MsgBox Left$(msg, 1000), IIf(failCount = 0, vbInformation, vbExclamation), "Term 01 Analysis"

Done:
    On Error Resume Next
    If Not wbGlobal Is Nothing Then wbGlobal.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.StatusBar = oldSB
    Application.DisplayAlerts = oldDA
    Application.EnableEvents = oldEV
    Application.ScreenUpdating = oldSU
    Panel_Show
    Exit Sub

Fatal:
    Panel_Finish 0, total, 0, "Stopped: " & Err.Number & " : " & Err.Description
    MsgBox "Build stopped: " & Err.Number & " : " & Err.Description, vbCritical, "Term 01 Analysis"
    Resume Done
End Sub

Private Function BuildOneFile(ByVal stem As String, ByVal titles As Variant, ByVal sheetIndex As Object, _
                              ByVal outDir As String, ByRef savedPath As String) As String
    Dim wbNew As Workbook, errText As String

    On Error GoTo EH
    Application.Calculation = xlCalculationManual
    Set wbNew = Workbooks.Add(xlWBATWorksheet)

    If Not BuildDataSheet(wbNew, sheetIndex, titles, stem, errText) Then Err.Raise vbObjectError + 901, "Build_All", errText
    If Not BuildAnalysisSheets(wbNew, CoreKeysFor(stem), errText) Then Err.Raise vbObjectError + 902, "Build_All", errText
    If Not BuildCompaireSheet(wbNew, errText) Then Err.Raise vbObjectError + 906, "Build_All", errText
    If Not BuildCartTecSheet(wbNew, errText) Then Err.Raise vbObjectError + 903, "Build_All", errText

    On Error Resume Next
    wbNew.Worksheets(DATA_SHEET).Activate
    wbNew.Worksheets(DATA_SHEET).Range("A1").Select
    On Error GoTo EH

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    savedPath = NextFreePath(outDir, stem, ".xlsx")
    wbNew.SaveAs Filename:=savedPath, FileFormat:=xlOpenXMLWorkbook
    If Not FileExistsP(savedPath) Then Err.Raise vbObjectError + 904, "Build_All", "The file was not found after saving: " & savedPath

    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing
    BuildOneFile = ""
    Exit Function
EH:
    BuildOneFile = Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
End Function

Private Sub WriteLog(ByVal path As String, ByVal txt As String)
    Dim fso As Object, ts As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, True)
    ts.Write txt
    ts.Close
    Err.Clear
End Sub
