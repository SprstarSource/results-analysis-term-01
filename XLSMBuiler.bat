@echo off
set "BUILD_DIR=%~dp0"
set "BAT_FILE=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:BAT_FILE,[Text.Encoding]::Default); $i=$s.IndexOf('#'+'PS_START'); Invoke-Expression $s.Substring($i)"
set "RC=%errorlevel%"
if not "%RC%"=="0" pause
exit /b %RC%

#PS_START
$ErrorActionPreference = "Stop"

$ProjectFolder = $env:BUILD_DIR.TrimEnd('\')
$OutputFile    = Join-Path $ProjectFolder "GeneratedWorkbook.xlsm"
$xlOpenXMLWorkbookMacroEnabled = 52

# Fixed root of the panel module name; the suffix varies per project (01, 02, ...)
$PanelRoot = "modBuildPanel"
$PanelProc = "RefreshPanelState"
$BuildPanelProc = "BuildControlPanel"

# Panel sheet detection: any sheet whose name contains one of these (Arabic built from char codes)
$ArabicPanel = -join ([char]0x0644,[char]0x0648,[char]0x062D,[char]0x0629)   # لوحة
$PanelSheetPatterns = @("*Panel*", "*$ArabicPanel*")

# Fallback ThisWorkbook code (used only if ThisWorkbook.cls is not found)
$DefaultThisWorkbookCode = @'
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo EH
    modBuildPanel.RefreshPanelState
    Exit Sub
EH:
    MsgBox "Panel error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Workbook_Activate()
    On Error Resume Next
    modBuildPanel.RefreshPanelState
    On Error GoTo 0
End Sub
'@

$Excel = $null; $Workbook = $null; $VBProject = $null
$Success = $false

# Read a document-module .cls and strip the export header (ANSI code page)
function Get-DocModuleCode([string]$Path) {
    $lines = [IO.File]::ReadAllLines($Path, [Text.Encoding]::Default)
    $i = 0
    while ($i -lt $lines.Count -and
           $lines[$i] -cmatch '^(VERSION |BEGIN|END\s*$|\s+MultiUse|Attribute VB_)') { $i++ }
    if ($i -ge $lines.Count) { return "" }
    return ($lines[$i..($lines.Count - 1)] -join "`r`n")
}

try {
    Write-Host "Starting Excel build..." -ForegroundColor Cyan

    $AllFiles = @(
        Get-ChildItem $ProjectFolder -Filter "*.bas" -File | Sort-Object Name
        Get-ChildItem $ProjectFolder -Filter "*.cls" -File | Sort-Object Name
        Get-ChildItem $ProjectFolder -Filter "*.frm" -File | Sort-Object Name
    )

    $HasThisWorkbookFile = [bool]($AllFiles | Where-Object { $_.Name -ieq "ThisWorkbook.cls" })
    $OtherFiles = @($AllFiles | Where-Object { $_.Name -ine "ThisWorkbook.cls" })

    Write-Host "Found $($OtherFiles.Count) module files."

    foreach ($f in $OtherFiles | Where-Object Extension -eq ".frm") {
        $frx = [IO.Path]::ChangeExtension($f.FullName, ".frx")
        if (-not (Test-Path $frx)) { Write-Warning "Missing FRX for $($f.Name)" }
    }

    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.EnableEvents = $false

    $Workbook = $Excel.Workbooks.Add()

    try { $VBProject = $Workbook.VBProject }
    catch { throw "Enable 'Trust access to the VBA project object model' in Excel Trust Center > Macro Settings." }

    # 1) Import modules
    foreach ($File in $OtherFiles) {
        Write-Host "Importing: $($File.Name)"
        try { [void]$VBProject.VBComponents.Import($File.FullName) }
        catch { throw "Failed importing $($File.Name): $_" }
    }

    # 2) Detect the panel module by its fixed root
    $PanelName = $null
    $PanelMatches = @()
    foreach ($comp in $VBProject.VBComponents) {
        if ($comp.Name -like "$PanelRoot*") { $PanelMatches += $comp.Name }
    }
    $PanelMatches = @($PanelMatches | Sort-Object)

    if ($PanelMatches.Count -eq 0) {
        Write-Warning "No module starting with '$PanelRoot' was found."
    } else {
        $PanelName = $PanelMatches[0]
        Write-Host "Panel module: $PanelName" -ForegroundColor Yellow
        if ($PanelMatches.Count -gt 1) {
            Write-Warning "Multiple panel modules found ($($PanelMatches -join ', ')). Using: $PanelName"
        }
        $pm = $VBProject.VBComponents.Item($PanelName).CodeModule
        $pText = $pm.Lines(1, $pm.CountOfLines)
        if ($pText -notmatch "(?i)\b(Sub|Function)\s+$PanelProc\b") { throw "Procedure '$PanelProc' not found in $PanelName." }
        if ($pText -notmatch "(?i)\bPublic\s+Sub\s+$BuildPanelProc\b") { throw "Procedure '$BuildPanelProc' not found in $PanelName." }
    }

    # 3) Inject ThisWorkbook code into the existing document module
    $cm = $VBProject.VBComponents.Item($Workbook.CodeName).CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }

    if ($HasThisWorkbookFile) {
        Write-Host "ThisWorkbook: from ThisWorkbook.cls"
        $code = Get-DocModuleCode (Join-Path $ProjectFolder "ThisWorkbook.cls")
    } else {
        Write-Host "ThisWorkbook: built-in default code"
        $code = $DefaultThisWorkbookCode
    }

    # Replace the root (or any suffixed name) with the real module name
    if ($PanelName) {
        $code = [regex]::Replace($code, "\b$PanelRoot\w*(?=\s*\.)", $PanelName)
    }
    $cm.AddFromString($code)

    # 4) Save
    $Workbook.SaveAs($OutputFile, $xlOpenXMLWorkbookMacroEnabled)
    $Success = $true

    Write-Host ""
    Write-Host "BUILD SUCCESSFUL" -ForegroundColor Green
    Write-Host "Output: $OutputFile"
}
catch {
    Write-Host "BUILD FAILED: $_" -ForegroundColor Red
    $global:LASTEXITCODE = 1
}
finally {
    if ($Workbook) { try { $Workbook.Close($false) } catch {} }
    if ($Excel)    { try { $Excel.Quit() } catch {} }

    foreach ($o in @($VBProject, $Workbook, $Excel)) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($Success) {
    # Launch: visible Excel, macros enabled for this session only, events ON
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.UserControl = $true           # keep Excel open after the script ends
    $previousAutomationSecurity = $xl.AutomationSecurity
    try {
        $xl.AutomationSecurity = 1
        $xl.EnableEvents = $true
        $wb = $xl.Workbooks.Open($OutputFile)   # Workbook_Open runs here
        $buildMacro = "'GeneratedWorkbook.xlsm'!$PanelName.$BuildPanelProc"
        Write-Host "Running: $buildMacro" -ForegroundColor Yellow
        [void]$xl.Run($buildMacro)
        $wb.Save()
    } finally {
        $xl.AutomationSecurity = $previousAutomationSecurity
    }

    # Go to the panel sheet
    $target = $null
    foreach ($ws in $wb.Worksheets) {
        foreach ($p in $PanelSheetPatterns) {
            if ($ws.Name -like $p) { $target = $ws; break }
        }
        if ($target) { break }
    }
    if ($target) {
        $target.Activate()
        [void]$target.Range("A1").Select()
        Write-Host "Panel sheet: $($target.Name)" -ForegroundColor Yellow
    } else {
        Write-Warning "No panel sheet found (patterns: $($PanelSheetPatterns -join ', ')). Staying on the active sheet."
    }
    exit 0
} else {
    exit 1
}