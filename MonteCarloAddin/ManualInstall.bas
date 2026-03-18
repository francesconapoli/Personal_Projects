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
'===============================================================================
Option Explicit

' *** CHANGE THIS PATH to where you extracted the project ***
Private Const SOURCE_PATH As String = "C:\Users\YourName\MonteCarloAddin\src"

Public Sub InstallMCSimAddin()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(SOURCE_PATH) Then
        MsgBox "Source path not found: " & SOURCE_PATH & vbCrLf & _
               "Edit the SOURCE_PATH constant in the code.", vbCritical
        Exit Sub
    End If

    ' Create new workbook for the addin
    Dim wb As Workbook
    Set wb = Workbooks.Add

    Dim vbProj As Object
    Set vbProj = wb.VBProject

    ' Import standard modules (.bas files)
    Dim modulesPath As String
    modulesPath = SOURCE_PATH & "\Modules"

    If fso.FolderExists(modulesPath) Then
        Dim f As Object
        For Each f In fso.GetFolder(modulesPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
                Debug.Print "Importing module: " & f.Name
                On Error Resume Next
                vbProj.VBComponents.Import f.Path
                If Err.Number <> 0 Then
                    Debug.Print "  ERROR: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next
    End If

    ' Build UserForms from code-only .frm files
    ' (The .frm files contain pure VBA code - no designer headers.
    '  All controls are created dynamically in UserForm_Initialize.)
    Dim formsPath As String
    formsPath = SOURCE_PATH & "\Forms"

    If fso.FolderExists(formsPath) Then
        For Each f In fso.GetFolder(formsPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = "frm" Then
                Debug.Print "Building form: " & f.Name
                BuildFormFromCode vbProj, f.Path, fso
            End If
        Next
    End If

    ' Update ThisWorkbook code
    Dim twbPath As String
    twbPath = SOURCE_PATH & "\ThisWorkbook.cls"
    If fso.FileExists(twbPath) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(twbPath, 1)
        Dim content As String
        content = ts.ReadAll
        ts.Close

        Dim twbModule As Object
        Set twbModule = vbProj.VBComponents("ThisWorkbook")
        Dim cm As Object
        Set cm = twbModule.CodeModule

        ' Extract code portion (skip VERSION, class header, Attribute lines)
        Dim lines() As String
        lines = Split(content, vbCrLf)
        If UBound(lines) < 1 Then lines = Split(content, vbLf)
        Dim codePart As String
        Dim inCode As Boolean
        inCode = False
        Dim i As Long
        For i = 0 To UBound(lines)
            Dim tl As String
            tl = Trim(lines(i))
            ' Skip Attribute lines and blank lines before code starts
            If Not inCode Then
                If Left(tl, 9) = "Attribute" Then GoTo NextLine
                If Left(tl, 7) = "VERSION" Then GoTo NextLine
                If Left(tl, 5) = "BEGIN" Or Left(tl, 5) = "Begin" Then GoTo NextLine
                If tl = "END" Or tl = "End" Then GoTo NextLine
                If tl = "" Then GoTo NextLine
                ' Found first real code line
                inCode = True
            End If
            codePart = codePart & lines(i) & vbCrLf
NextLine:
        Next

        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        If Len(Trim(codePart)) > 0 Then cm.AddFromString codePart
    End If

    ' Save as XLAM
    Dim savePath As String
    savePath = Application.DefaultFilePath & "\MCSimAddin.xlam"

    Application.DisplayAlerts = False
    wb.SaveAs savePath, 55 ' xlOpenXMLAddIn
    Application.DisplayAlerts = True

    MsgBox "Add-in created successfully!" & vbCrLf & vbCrLf & _
           "Saved to: " & savePath & vbCrLf & vbCrLf & _
           "To install:" & vbCrLf & _
           "1. File > Options > Add-ins" & vbCrLf & _
           "2. Manage: Excel Add-ins > Go..." & vbCrLf & _
           "3. Browse... > select MCSimAddin.xlam" & vbCrLf & _
           "4. Check it and click OK", vbInformation, "MCSimAddin"

    wb.Close False
End Sub

Private Sub BuildFormFromCode(vbProj As Object, filePath As String, fso As Object)
    ' Creates a blank UserForm and injects VBA code from a .frm file.
    ' The .frm files are code-only (no VERSION/Begin/End header, no .frx needed).
    ' All form controls are created dynamically in UserForm_Initialize.
    On Error GoTo ErrHandler

    ' Read the code file
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1)
    Dim codeText As String
    codeText = ts.ReadAll
    ts.Close

    ' Get form name from filename (e.g. "frmAddInput" from "frmAddInput.frm")
    Dim formName As String
    formName = fso.GetBaseName(filePath)

    ' Create a blank UserForm (vbext_ct_MSForm = 3)
    Dim formComp As Object
    Set formComp = vbProj.VBComponents.Add(3)
    formComp.Name = formName

    ' Set the caption from the form name (will be overridden by Me.Width/Height in Initialize)
    On Error Resume Next
    formComp.Properties("Caption") = formName
    On Error GoTo ErrHandler

    ' Inject the VBA code into the form's code module
    Dim cm As Object
    Set cm = formComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText

    Debug.Print "  Created form: " & formName
    Exit Sub

ErrHandler:
    Debug.Print "  ERROR building " & formName & ": " & Err.Description
End Sub
