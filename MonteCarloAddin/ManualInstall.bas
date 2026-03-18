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

    ' Import modules
    Dim modulesPath As String
    modulesPath = SOURCE_PATH & "\Modules"

    If fso.FolderExists(modulesPath) Then
        Dim f As Object
        For Each f In fso.GetFolder(modulesPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
                Debug.Print "Importing: " & f.Name
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

    ' Build forms programmatically (direct .frm import requires .frx binary files)
    Dim formsPath As String
    formsPath = SOURCE_PATH & "\Forms"

    If fso.FolderExists(formsPath) Then
        For Each f In fso.GetFolder(formsPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = "frm" Then
                Debug.Print "Building form: " & f.Name
                BuildFormFromFile vbProj, f.Path, fso
            End If
        Next
    End If

    ' Update ThisWorkbook
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

        ' Extract code portion
        Dim lines() As String
        lines = Split(content, vbCrLf)
        Dim codePart As String
        Dim inCode As Boolean
        inCode = False
        Dim i As Long
        For i = 0 To UBound(lines)
            If Left(Trim(lines(i)), 6) = "Option" Or _
               Left(Trim(lines(i)), 7) = "Private" Or _
               Left(Trim(lines(i)), 6) = "Public" Or inCode Then
                inCode = True
                codePart = codePart & lines(i) & vbCrLf
            End If
        Next

        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        cm.AddFromString codePart
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

Private Sub BuildFormFromFile(vbProj As Object, filePath As String, fso As Object)
    ' Reads a .frm file, creates a blank UserForm, and injects the VBA code.
    ' This avoids the need for .frx companion files.
    On Error GoTo ErrHandler

    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1)
    Dim content As String
    content = ts.ReadAll
    ts.Close

    Dim allLines() As String
    allLines = Split(content, vbCrLf)
    If UBound(allLines) < 1 Then
        allLines = Split(content, vbLf)
    End If

    ' Get form name from filename
    Dim formName As String
    formName = fso.GetBaseName(filePath)

    ' Extract caption from the header block
    Dim caption As String
    caption = ""
    Dim i As Long
    For i = 0 To UBound(allLines)
        Dim trimLine As String
        trimLine = Trim(allLines(i))
        If Left(trimLine, 7) = "Caption" Then
            Dim qStart As Long, qEnd As Long
            qStart = InStr(trimLine, """")
            qEnd = InStrRev(trimLine, """")
            If qStart > 0 And qEnd > qStart Then
                caption = Mid(trimLine, qStart + 1, qEnd - qStart - 1)
            End If
        End If
        If trimLine = "End" Then Exit For
    Next i

    ' Create blank UserForm (vbext_ct_MSForm = 3)
    Dim formComp As Object
    Set formComp = vbProj.VBComponents.Add(3)
    formComp.Name = formName

    ' Set caption
    If Len(caption) > 0 Then
        On Error Resume Next
        formComp.Designer.caption = caption
        On Error GoTo ErrHandler
    End If

    ' Extract VBA code: skip Begin...End header block and Attribute lines
    Dim codeText As String
    codeText = ""
    Dim pastHeader As Boolean
    Dim pastAttributes As Boolean
    pastHeader = False
    pastAttributes = False

    For i = 0 To UBound(allLines)
        trimLine = Trim(allLines(i))
        If Not pastHeader Then
            If trimLine = "End" Then pastHeader = True
        ElseIf Not pastAttributes Then
            If Left(trimLine, 9) = "Attribute" Then
                ' skip
            Else
                pastAttributes = True
                codeText = codeText & allLines(i) & vbCrLf
            End If
        Else
            codeText = codeText & allLines(i) & vbCrLf
        End If
    Next i

    ' Trim leading/trailing whitespace
    codeText = Trim(codeText)

    If Len(codeText) > 0 Then
        Dim cm As Object
        Set cm = formComp.CodeModule
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        cm.AddFromString codeText
    End If

    Debug.Print "  Created: " & formName
    Exit Sub

ErrHandler:
    Debug.Print "  ERROR building " & formName & ": " & Err.Description
End Sub
