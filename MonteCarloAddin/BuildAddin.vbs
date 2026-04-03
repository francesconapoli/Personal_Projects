'===============================================================================
' BuildAddin.vbs
' Purpose: Automated build script to create MCSimAddin.xlam
'          Run this from command line: cscript BuildAddin.vbs
'          Or double-click in Windows Explorer
'===============================================================================
Option Explicit

Dim xlApp, wb, vbProj
Dim fso, scriptDir, srcDir

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
srcDir = fso.BuildPath(scriptDir, "src")

WScript.Echo "=============================================="
WScript.Echo " MCSimAddin Build Script"
WScript.Echo "=============================================="
WScript.Echo ""

' Check if source files exist
If Not fso.FolderExists(srcDir) Then
    WScript.Echo "ERROR: src folder not found at " & srcDir
    WScript.Quit 1
End If

' Start Excel
WScript.Echo "Starting Excel..."
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False
xlApp.ScreenUpdating = False

' Check Trust Center settings
On Error Resume Next
xlApp.AutomationSecurity = 1 ' msoAutomationSecurityLow
On Error GoTo 0

' Create new workbook
WScript.Echo "Creating workbook..."
Set wb = xlApp.Workbooks.Add
Set vbProj = wb.VBProject

' Import modules
WScript.Echo "Importing VBA modules..."

Dim modulesDir, formsDir, classesDir
modulesDir = fso.BuildPath(srcDir, "Modules")
formsDir = fso.BuildPath(srcDir, "Forms")
classesDir = fso.BuildPath(srcDir, "Classes")

' Import all .bas files from Modules
If fso.FolderExists(modulesDir) Then
    Dim basFile
    For Each basFile In fso.GetFolder(modulesDir).Files
        If LCase(fso.GetExtensionName(basFile.Name)) = "bas" Then
            WScript.Echo "  + " & basFile.Name
            On Error Resume Next
            vbProj.VBComponents.Import basFile.Path
            If Err.Number <> 0 Then
                WScript.Echo "    WARNING: Could not import " & basFile.Name & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
End If

' Import all .frm files from Forms
If fso.FolderExists(formsDir) Then
    Dim frmFile
    For Each frmFile In fso.GetFolder(formsDir).Files
        If LCase(fso.GetExtensionName(frmFile.Name)) = "frm" Then
            WScript.Echo "  + " & frmFile.Name
            On Error Resume Next
            vbProj.VBComponents.Import frmFile.Path
            If Err.Number <> 0 Then
                WScript.Echo "    WARNING: Could not import " & frmFile.Name & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
End If

' Import class modules
If fso.FolderExists(classesDir) Then
    Dim clsFile
    For Each clsFile In fso.GetFolder(classesDir).Files
        If LCase(fso.GetExtensionName(clsFile.Name)) = "cls" Then
            WScript.Echo "  + " & clsFile.Name
            On Error Resume Next
            vbProj.VBComponents.Import clsFile.Path
            If Err.Number <> 0 Then
                WScript.Echo "    WARNING: Could not import " & clsFile.Name & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
End If

' Update ThisWorkbook code
WScript.Echo "Updating ThisWorkbook..."
Dim twbPath
twbPath = fso.BuildPath(srcDir, "ThisWorkbook.cls")
If fso.FileExists(twbPath) Then
    Dim twbCode
    Dim ts
    Set ts = fso.OpenTextFile(twbPath, 1)
    twbCode = ts.ReadAll
    ts.Close

    ' Remove header lines (VERSION, BEGIN, etc.)
    Dim lines, codePart, inCode, line
    lines = Split(twbCode, vbCrLf)
    codePart = ""
    inCode = False
    Dim idx
    For idx = 0 To UBound(lines)
        line = lines(idx)
        If Left(Trim(line), 6) = "Option" Or Left(Trim(line), 7) = "Private" Or _
           Left(Trim(line), 6) = "Public" Or inCode Then
            inCode = True
            codePart = codePart & line & vbCrLf
        End If
    Next

    On Error Resume Next
    Dim twbModule
    Set twbModule = vbProj.VBComponents("ThisWorkbook")
    twbModule.CodeModule.DeleteLines 1, twbModule.CodeModule.CountOfLines
    twbModule.CodeModule.AddFromString codePart
    If Err.Number <> 0 Then
        WScript.Echo "    WARNING: Could not update ThisWorkbook - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End If

' Save as XLAM
Dim outputPath
outputPath = fso.BuildPath(scriptDir, "MCSimAddin.xlam")
WScript.Echo ""
WScript.Echo "Saving as XLAM..."
WScript.Echo "  Output: " & outputPath

' Delete existing file
If fso.FileExists(outputPath) Then
    fso.DeleteFile outputPath
End If

On Error Resume Next
wb.SaveAs outputPath, 55 ' xlOpenXMLAddIn = 55
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not save XLAM - " & Err.Description
    WScript.Echo "Trying alternate save format..."
    Err.Clear
    wb.SaveAs outputPath, 18 ' xlAddIn = 18 (for .xla)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Save failed - " & Err.Description
        Err.Clear
    End If
End If
On Error GoTo 0

wb.Close False

' Add Custom UI XML to the XLAM (it's a ZIP file)
WScript.Echo "Adding ribbon XML..."
Dim cuiPath
cuiPath = fso.BuildPath(srcDir, "CustomUI\customUI14.xml")
If fso.FileExists(cuiPath) And fso.FileExists(outputPath) Then
    WScript.Echo "  NOTE: Use the Custom UI Editor tool to inject customUI14.xml"
    WScript.Echo "  Download from: https://github.com/fernandreu/office-ribbonx-editor"
    WScript.Echo "  Or use the OOXML Packager included in this project."
End If

xlApp.Quit

WScript.Echo ""
WScript.Echo "=============================================="
WScript.Echo " Build Complete!"
WScript.Echo "=============================================="
WScript.Echo ""
WScript.Echo "Output: " & outputPath
WScript.Echo ""
WScript.Echo "INSTALLATION:"
WScript.Echo "  1. Open Excel"
WScript.Echo "  2. File > Options > Add-ins"
WScript.Echo "  3. Manage: Excel Add-ins > Go..."
WScript.Echo "  4. Browse... and select MCSimAddin.xlam"
WScript.Echo "  5. Check the add-in and click OK"
WScript.Echo ""
WScript.Echo "RIBBON XML (for custom tab):"
WScript.Echo "  Use Office RibbonX Editor to inject"
WScript.Echo "  src\CustomUI\customUI14.xml into the XLAM"
WScript.Echo ""

Set wb = Nothing
Set xlApp = Nothing
Set fso = Nothing
