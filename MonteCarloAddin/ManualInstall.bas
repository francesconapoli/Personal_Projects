Attribute VB_Name = "ManualInstall"
'===============================================================================
' ManualInstall.bas
' Purpose: Run this macro from Excel VBA editor to install the addin manually
'
' INSTRUCTIONS:
' 1. Open Excel
' 2. Press Alt+F11 to open VBA Editor
' 3. Insert > Module
' 4. Paste this entire file
' 5. Change the SOURCE_PATH constant below to your actual path
' 6. Run the InstallMCSimAddin macro
'
' PREREQUISITES:
' - File > Options > Trust Center > Trust Center Settings >
'   Macro Settings > Check "Trust access to the VBA project object model"
'===============================================================================
Option Explicit

' *** CHANGE THIS PATH to where you extracted the project ***
Private Const SOURCE_PATH As String = "C:\Users\YourName\MonteCarloAddin\src"

' *** VBA project password (set to "" to skip locking) ***
Private Const VBA_PASSWORD As String = "MCSimAddin2024"

Private m_FormCount As Long
Private m_ModuleCount As Long
Private m_Errors As String

Public Sub InstallMCSimAddin()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(SOURCE_PATH) Then
        MsgBox "Source path not found: " & SOURCE_PATH & vbCrLf & _
               "Edit the SOURCE_PATH constant in the code.", vbCritical
        Exit Sub
    End If

    m_FormCount = 0
    m_ModuleCount = 0
    m_Errors = ""

    ' Create new workbook for the addin
    Dim wb As Workbook
    Set wb = Workbooks.Add

    Dim vbProj As Object
    Set vbProj = wb.VBProject

    ' ---- Step 1: Import standard modules (.bas files) ----
    Dim modulesPath As String
    modulesPath = SOURCE_PATH & "\Modules"

    If fso.FolderExists(modulesPath) Then
        Dim f As Object
        For Each f In fso.GetFolder(modulesPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
                On Error Resume Next
                vbProj.VBComponents.Import f.Path
                If Err.Number <> 0 Then
                    m_Errors = m_Errors & "Module " & f.Name & ": " & Err.Description & vbCrLf
                    Err.Clear
                Else
                    m_ModuleCount = m_ModuleCount + 1
                End If
                On Error GoTo 0
            End If
        Next
    End If

    ' ---- Step 2: Build UserForms from code-only .frm files ----
    ' The .frm files contain pure VBA code (no designer headers).
    ' All controls are created dynamically in UserForm_Initialize.
    Dim formsPath As String
    formsPath = SOURCE_PATH & "\Forms"

    If fso.FolderExists(formsPath) Then
        For Each f In fso.GetFolder(formsPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = "frm" Then
                BuildFormFromCode vbProj, f.Path, fso
            End If
        Next
    End If

    ' ---- Step 3: Update ThisWorkbook code ----
    Dim twbPath As String
    twbPath = SOURCE_PATH & "\ThisWorkbook.cls"
    If fso.FileExists(twbPath) Then
        Dim twbCode As String
        twbCode = ReadFileText(fso, twbPath)
        twbCode = StripHeaderLines(twbCode)

        Dim twbModule As Object
        Set twbModule = vbProj.VBComponents("ThisWorkbook")
        Dim cm As Object
        Set cm = twbModule.CodeModule
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        If Len(Trim(twbCode)) > 0 Then cm.AddFromString twbCode
    End If

    ' ---- Step 4: Save as XLAM ----
    Dim savePath As String
    savePath = Application.DefaultFilePath & "\MCSimAddin.xlam"

    Application.DisplayAlerts = False
    wb.SaveAs savePath, 55 ' xlOpenXMLAddIn
    Application.DisplayAlerts = True

    ' ---- Step 5: Lock VBA project (after save) ----
    If Len(VBA_PASSWORD) > 0 Then
        LockVBAProject wb
    End If

    wb.Close True  ' Save again after locking

    ' ---- Report results ----
    Dim msg As String
    msg = "Installation complete!" & vbCrLf & vbCrLf & _
          "Modules imported: " & m_ModuleCount & vbCrLf & _
          "UserForms created: " & m_FormCount & vbCrLf & _
          "Saved to: " & savePath

    If Len(m_Errors) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "ERRORS:" & vbCrLf & m_Errors
    End If

    If m_FormCount = 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "WARNING: No forms were created! Check that:" & vbCrLf & _
              "- 'Trust access to the VBA project object model' is enabled" & vbCrLf & _
              "- The Forms folder exists at: " & formsPath
    End If

    msg = msg & vbCrLf & vbCrLf & _
          "To install:" & vbCrLf & _
          "1. File > Options > Add-ins" & vbCrLf & _
          "2. Manage: Excel Add-ins > Go..." & vbCrLf & _
          "3. Browse... > select MCSimAddin.xlam" & vbCrLf & _
          "4. Check it and click OK"

    MsgBox msg, IIf(Len(m_Errors) > 0, vbExclamation, vbInformation), "MCSimAddin Installer"
End Sub

'===============================================================================
' BuildFormFromCode - Creates a UserForm and injects VBA code
'===============================================================================
Private Sub BuildFormFromCode(vbProj As Object, filePath As String, fso As Object)
    Dim formName As String
    formName = fso.GetBaseName(filePath)

    ' Read and normalize the code
    Dim codeText As String
    codeText = ReadFileText(fso, filePath)

    If Len(Trim(codeText)) = 0 Then
        m_Errors = m_Errors & "Form " & formName & ": File is empty" & vbCrLf
        Exit Sub
    End If

    ' Create a blank UserForm (vbext_ct_MSForm = 3)
    Dim formComp As Object
    On Error Resume Next
    Set formComp = vbProj.VBComponents.Add(3)
    If Err.Number <> 0 Then
        m_Errors = m_Errors & "Form " & formName & " (Add): " & Err.Description & vbCrLf
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Rename the form
    On Error Resume Next
    formComp.Name = formName
    If Err.Number <> 0 Then
        m_Errors = m_Errors & "Form " & formName & " (Name): " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0

    ' Inject the VBA code
    On Error Resume Next
    Dim cm As Object
    Set cm = formComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
    If Err.Number <> 0 Then
        m_Errors = m_Errors & "Form " & formName & " (Code): " & Err.Description & vbCrLf
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    m_FormCount = m_FormCount + 1
End Sub

'===============================================================================
' ReadFileText - Reads a text file and normalizes line endings to vbCrLf
'===============================================================================
Private Function ReadFileText(fso As Object, filePath As String) As String
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)
    Dim raw As String
    If Not ts.AtEndOfStream Then
        raw = ts.ReadAll
    End If
    ts.Close

    ' Normalize line endings: convert any LF-only or CR-only to CrLf
    ' (files may have been edited on Linux/Mac with LF endings)
    raw = Replace(raw, vbCrLf, vbLf)   ' First normalize CrLf to Lf
    raw = Replace(raw, vbCr, vbLf)      ' Then normalize lone Cr to Lf
    raw = Replace(raw, vbLf, vbCrLf)    ' Finally convert all Lf to CrLf

    ReadFileText = raw
End Function

'===============================================================================
' StripHeaderLines - Removes Attribute, VERSION, Begin/End header lines
'===============================================================================
Private Function StripHeaderLines(ByVal txt As String) As String
    Dim lines() As String
    lines = Split(txt, vbCrLf)
    Dim result As String
    Dim inCode As Boolean
    inCode = False
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim tl As String
        tl = Trim(lines(i))
        If Not inCode Then
            If Left(tl, 9) = "Attribute" Then GoTo SkipLine
            If Left(tl, 7) = "VERSION" Then GoTo SkipLine
            If Left(tl, 5) = "BEGIN" Or Left(tl, 5) = "Begin" Then GoTo SkipLine
            If tl = "END" Or tl = "End" Then GoTo SkipLine
            If tl = "" Then GoTo SkipLine
            inCode = True
        End If
        result = result & lines(i) & vbCrLf
SkipLine:
    Next i
    StripHeaderLines = result
End Function

'===============================================================================
' LockVBAProject - Password-protects the VBA project
' Uses SendKeys to automate the Protection dialog (no direct API exists).
' The VBE window will flash briefly during this process.
'===============================================================================
Private Sub LockVBAProject(wb As Workbook)
    On Error GoTo LockError

    ' Make sure VBE is ready
    Application.VBE.MainWindow.Visible = True
    DoEvents

    ' Select the target project in the Project Explorer
    Dim vbc As Object
    Set vbc = wb.VBProject.VBComponents(1)
    vbc.Activate
    DoEvents

    ' Send keystrokes: Tools > VBAProject Properties
    SendKeys "%{F11}", True   ' Ensure VBE is focused
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)

    SendKeys "%(T)E", True    ' Tools menu > Project Properties
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents

    ' Protection tab
    SendKeys "^{TAB}", True
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)

    ' Check "Lock project for viewing"
    SendKeys " ", True
    DoEvents

    ' Tab to password, type it, tab to confirm, type again, Enter
    SendKeys "{TAB}" & VBA_PASSWORD & "{TAB}" & VBA_PASSWORD & "{ENTER}", True
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)

    Application.VBE.MainWindow.Visible = False
    Exit Sub

LockError:
    On Error Resume Next
    Application.VBE.MainWindow.Visible = False
    m_Errors = m_Errors & "Password lock: " & Err.Description & vbCrLf & _
               "You can lock manually: VBA Editor > Tools > VBAProject Properties > Protection" & vbCrLf
End Sub
