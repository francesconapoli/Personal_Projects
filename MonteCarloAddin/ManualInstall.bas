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
' Can point to the MonteCarloAddin folder OR the src subfolder - both work.
Private Const SOURCE_PATH As String = "C:\Users\YourName\MonteCarloAddin"

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

    ' Auto-detect the correct base path
    ' Works whether SOURCE_PATH points to MonteCarloAddin or MonteCarloAddin\src
    Dim basePath As String
    basePath = SOURCE_PATH
    If fso.FolderExists(basePath & "\src\Modules") Then
        basePath = basePath & "\src"
    ElseIf Not fso.FolderExists(basePath & "\Modules") Then
        ' Try one more: maybe they pointed to the repo root containing MonteCarloAddin
        If fso.FolderExists(basePath & "\MonteCarloAddin\src\Modules") Then
            basePath = basePath & "\MonteCarloAddin\src"
        Else
            MsgBox "Cannot find Modules folder." & vbCrLf & vbCrLf & _
                   "Looked in:" & vbCrLf & _
                   "  " & SOURCE_PATH & "\Modules" & vbCrLf & _
                   "  " & SOURCE_PATH & "\src\Modules" & vbCrLf & vbCrLf & _
                   "Set SOURCE_PATH to the folder containing the Modules and Forms subfolders.", _
                   vbCritical, "MCSimAddin Installer"
            Exit Sub
        End If
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
    modulesPath = basePath & "\Modules"

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
                    Debug.Print "  Imported: " & f.Name
                End If
                On Error GoTo 0
            End If
        Next
    Else
        m_Errors = m_Errors & "Modules folder not found: " & modulesPath & vbCrLf
    End If

    ' ---- Step 2: Build UserForms from .frm files ----
    ' The .frm files use @CTRL/@SET directives to define design-time controls.
    ' The installer adds controls via formComp.Designer so events bind correctly.
    Dim formsPath As String
    formsPath = basePath & "\Forms"

    If fso.FolderExists(formsPath) Then
        For Each f In fso.GetFolder(formsPath).Files
            If LCase(fso.GetExtensionName(f.Name)) = "frm" Then
                BuildFormFromCode vbProj, f.Path, fso
            End If
        Next
    Else
        m_Errors = m_Errors & "Forms folder not found: " & formsPath & vbCrLf
    End If

    ' ---- Step 3: Update ThisWorkbook code ----
    Dim twbPath As String
    twbPath = basePath & "\ThisWorkbook.cls"
    If fso.FileExists(twbPath) Then
        Dim twbCode As String
        twbCode = ReadFileText(fso, twbPath)
        twbCode = StripHeaderLines(twbCode)

        Dim twbModule As Object
        Set twbModule = vbProj.VBComponents("ThisWorkbook")
        Dim cm As Object
        Set cm = twbModule.CodeModule
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        If Len(Trim(twbCode)) > 0 Then
            On Error Resume Next
            cm.AddFromString twbCode
            If Err.Number <> 0 Then
                m_Errors = m_Errors & "ThisWorkbook code: " & Err.Description & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Else
        m_Errors = m_Errors & "ThisWorkbook.cls not found" & vbCrLf
    End If

    ' ---- Step 4: Verify components ----
    Dim verifyMsg As String
    verifyMsg = VerifyComponents(vbProj)

    ' ---- Step 5: Save as XLAM ----
    Dim savePath As String
    savePath = Application.DefaultFilePath & "\MCSimAddin.xlam"

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.SaveAs savePath, 55 ' xlOpenXMLAddIn
    If Err.Number <> 0 Then
        m_Errors = m_Errors & "Save failed: " & Err.Description & vbCrLf
        Err.Clear
        ' Try alternate location
        savePath = Environ("USERPROFILE") & "\Documents\MCSimAddin.xlam"
        wb.SaveAs savePath, 55
        If Err.Number <> 0 Then
            m_Errors = m_Errors & "Save to Documents also failed: " & Err.Description & vbCrLf
            Err.Clear
            On Error GoTo 0
            Application.DisplayAlerts = True
            MsgBox "CRITICAL: Could not save the add-in file." & vbCrLf & m_Errors, vbCritical
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True

    wb.Close False

    ' ---- Report results ----
    Dim msg As String
    msg = "Installation complete!" & vbCrLf & vbCrLf & _
          "Modules imported: " & m_ModuleCount & vbCrLf & _
          "UserForms created: " & m_FormCount & vbCrLf & _
          "Saved to: " & savePath

    If Len(verifyMsg) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "VERIFICATION:" & vbCrLf & verifyMsg
    End If

    If Len(m_Errors) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "ERRORS:" & vbCrLf & m_Errors
    End If

    If m_FormCount = 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "WARNING: No forms were created!" & vbCrLf & _
              "Check that 'Trust access to the VBA project object model' is enabled:" & vbCrLf & _
              "File > Options > Trust Center > Trust Center Settings > Macro Settings"
    End If

    msg = msg & vbCrLf & vbCrLf & _
          "TO INSTALL THE ADD-IN:" & vbCrLf & _
          "1. File > Options > Add-ins" & vbCrLf & _
          "2. Manage: Excel Add-ins > Go..." & vbCrLf & _
          "3. Browse... > select MCSimAddin.xlam" & vbCrLf & _
          "4. Check it and click OK" & vbCrLf & vbCrLf & _
          "TO PASSWORD-PROTECT (optional):" & vbCrLf & _
          "1. Open the .xlam, press Alt+F11" & vbCrLf & _
          "2. Tools > VBAProject Properties > Protection tab" & vbCrLf & _
          "3. Check 'Lock project for viewing'" & vbCrLf & _
          "4. Set password, click OK, save"

    MsgBox msg, IIf(Len(m_Errors) > 0, vbExclamation, vbInformation), "MCSimAddin Installer"
End Sub

'===============================================================================
' VerifyComponents - Checks that all expected components exist
'===============================================================================
Private Function VerifyComponents(vbProj As Object) As String
    Dim missing As String
    Dim found As String

    ' Expected modules
    Dim expectedModules As Variant
    expectedModules = Array("modGlobals", "modSimEngine", "modDistributions", _
        "modRandomGenerators", "modCharting", "modSensitivity", _
        "modOutputFunctions", "modDistFitting", "modToolbar", "modRibbonCallbacks")

    ' Expected forms
    Dim expectedForms As Variant
    expectedForms = Array("frmAddInput", "frmCorrelation", "frmModelWindow", _
        "frmOutputSelect", "frmResults", "frmSensitivity", "frmSettings")

    Dim i As Long
    Dim comp As Object

    ' Check modules
    For i = 0 To UBound(expectedModules)
        On Error Resume Next
        Set comp = vbProj.VBComponents(CStr(expectedModules(i)))
        If Err.Number <> 0 Then
            missing = missing & "  MISSING: " & expectedModules(i) & vbCrLf
            Err.Clear
        Else
            found = found & "  OK: " & expectedModules(i) & vbCrLf
        End If
        On Error GoTo 0
    Next i

    ' Check forms
    For i = 0 To UBound(expectedForms)
        On Error Resume Next
        Set comp = vbProj.VBComponents(CStr(expectedForms(i)))
        If Err.Number <> 0 Then
            missing = missing & "  MISSING: " & expectedForms(i) & vbCrLf
            Err.Clear
        Else
            found = found & "  OK: " & expectedForms(i) & vbCrLf
        End If
        On Error GoTo 0
    Next i

    If Len(missing) > 0 Then
        VerifyComponents = "MISSING components:" & vbCrLf & missing
    Else
        VerifyComponents = "All 17 components verified OK"
    End If
End Function

'===============================================================================
' BuildFormFromCode - Creates a UserForm with design-time controls and VBA code
'
' The .frm files use special comment directives to define controls:
'   '@FORM:Width,Height,Caption
'   '@CTRL:Type,Name,Parent,Left,Top,Width,Height[,Caption]
'   '@SET:ControlName,Property,Value
'   '@END
' Controls are added via formComp.Designer so they are design-time controls
' with working event handlers. The remaining code is injected normally.
'===============================================================================
Private Sub BuildFormFromCode(vbProj As Object, filePath As String, fso As Object)
    Dim formName As String
    formName = fso.GetBaseName(filePath)
    Debug.Print "Building form: " & formName

    ' Read and normalize the code
    Dim fullText As String
    fullText = ReadFileText(fso, filePath)

    If Len(Trim(fullText)) = 0 Then
        m_Errors = m_Errors & "Form " & formName & ": File is empty" & vbCrLf
        Exit Sub
    End If

    ' Create a blank UserForm (vbext_ct_MSForm = 3)
    Dim formComp As Object
    On Error Resume Next
    Set formComp = vbProj.VBComponents.Add(3)
    If Err.Number <> 0 Or formComp Is Nothing Then
        m_Errors = m_Errors & "Form " & formName & " (Add): " & _
            IIf(Err.Number <> 0, Err.Description, "VBComponents.Add returned Nothing") & vbCrLf
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

    ' Add design-time controls from @CTRL directives and extract code
    Dim codeText As String
    codeText = SetupDesignerControls(formComp, fullText, formName)

    ' Inject the VBA code into the form's code module
    Dim cm As Object
    Set cm = formComp.CodeModule

    On Error Resume Next
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    cm.AddFromString codeText
    If Err.Number <> 0 Then
        m_Errors = m_Errors & "Form " & formName & " (Code): " & Err.Description & vbCrLf & _
            "  Code length: " & Len(codeText) & " chars, " & _
            UBound(Split(codeText, vbCrLf)) + 1 & " lines" & vbCrLf
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    m_FormCount = m_FormCount + 1
    Debug.Print "  Created form: " & formName & " (" & cm.CountOfLines & " lines)"
End Sub

'===============================================================================
' SetupDesignerControls - Parses @CTRL/@SET/@FORM directives and creates
'   design-time controls on the form Designer surface.
'   Returns the remaining code text (with directives stripped).
'===============================================================================
Private Function SetupDesignerControls(formComp As Object, ByVal txt As String, _
    ByVal formName As String) As String

    Dim lines() As String
    lines = Split(txt, vbCrLf)
    Dim codeLines As String
    Dim designer As Object
    Dim hasDirectives As Boolean
    hasDirectives = False

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim tl As String
        tl = Trim(lines(i))

        If Left(tl, 7) = "'@FORM:" Then
            hasDirectives = True
        ElseIf Left(tl, 7) = "'@CTRL:" Then
            hasDirectives = True
        ElseIf Left(tl, 6) = "'@SET:" Then
            hasDirectives = True
        ElseIf tl = "'@END" Then
            hasDirectives = True
        End If
    Next i

    ' If no directives found, return text as-is (backwards compatible)
    If Not hasDirectives Then
        SetupDesignerControls = txt
        Exit Function
    End If

    ' Get the Designer surface for adding design-time controls
    On Error Resume Next
    Set designer = formComp.designer
    If designer Is Nothing Then
        m_Errors = m_Errors & "Form " & formName & ": Cannot access Designer" & vbCrLf
        On Error GoTo 0
        SetupDesignerControls = txt
        Exit Function
    End If
    On Error GoTo 0

    ' Process directives
    For i = 0 To UBound(lines)
        tl = Trim(lines(i))

        If Left(tl, 7) = "'@FORM:" Then
            ParseFormDirective designer, formComp, Mid(tl, 8)
        ElseIf Left(tl, 7) = "'@CTRL:" Then
            ParseCtrlDirective designer, Mid(tl, 8), formName
        ElseIf Left(tl, 6) = "'@SET:" Then
            ParseSetDirective designer, Mid(tl, 7), formName
        ElseIf tl = "'@END" Then
            ' Skip
        Else
            codeLines = codeLines & lines(i) & vbCrLf
        End If
    Next i

    SetupDesignerControls = codeLines
End Function

'===============================================================================
' ParseFormDirective - Sets form Width, Height, Caption
'   Format: Width,Height,Caption
'===============================================================================
Private Sub ParseFormDirective(designer As Object, formComp As Object, ByVal directive As String)
    Dim parts() As String
    parts = Split(directive, ",")

    On Error Resume Next
    If UBound(parts) >= 0 Then designer.Width = CSng(Trim(parts(0)))
    If UBound(parts) >= 1 Then designer.Height = CSng(Trim(parts(1)))
    If UBound(parts) >= 2 Then
        ' Caption may contain commas, rejoin from index 2 onward
        Dim cap As String
        Dim j As Long
        cap = Trim(parts(2))
        For j = 3 To UBound(parts)
            cap = cap & "," & parts(j)
        Next j
        formComp.Properties("Caption") = cap
    End If
    Err.Clear
    On Error GoTo 0
End Sub

'===============================================================================
' ParseCtrlDirective - Adds a control to the Designer
'   Format: Type,Name,Parent,Left,Top,Width,Height[,Caption]
'   Type: Label|TextBox|ComboBox|CommandButton|ListBox|Frame|OptionButton|CheckBox
'   Parent: empty for form-level, or name of a Frame
'===============================================================================
Private Sub ParseCtrlDirective(designer As Object, ByVal directive As String, ByVal formName As String)
    Dim parts() As String
    parts = Split(directive, ",")
    If UBound(parts) < 6 Then Exit Sub ' Need at least 7 fields

    Dim ctrlType As String, ctrlName As String, parentName As String
    ctrlType = Trim(parts(0))
    ctrlName = Trim(parts(1))
    parentName = Trim(parts(2))

    Dim l As Single, t As Single, w As Single, h As Single
    l = CSng(Trim(parts(3)))
    t = CSng(Trim(parts(4)))
    w = CSng(Trim(parts(5)))
    h = CSng(Trim(parts(6)))

    ' Caption may contain commas
    Dim cap As String
    If UBound(parts) >= 7 Then
        cap = Trim(parts(7))
        Dim j As Long
        For j = 8 To UBound(parts)
            cap = cap & "," & parts(j)
        Next j
    End If

    ' Map type name to ProgID
    Dim progId As String
    progId = GetControlProgId(ctrlType)
    If progId = "" Then
        m_Errors = m_Errors & "Form " & formName & ": Unknown control type '" & ctrlType & "'" & vbCrLf
        Exit Sub
    End If

    ' Determine parent container
    Dim container As Object
    If parentName = "" Then
        Set container = designer
    Else
        On Error Resume Next
        Set container = designer.Controls(parentName)
        If container Is Nothing Then
            m_Errors = m_Errors & "Form " & formName & ": Parent '" & parentName & _
                "' not found for " & ctrlName & vbCrLf
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If

    ' Add the control
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = container.Controls.Add(progId, ctrlName)
    If Err.Number <> 0 Or ctrl Is Nothing Then
        m_Errors = m_Errors & "Form " & formName & ": Failed to add " & ctrlName & _
            " - " & Err.Description & vbCrLf
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Set position and size
    ctrl.Left = l
    ctrl.Top = t
    ctrl.Width = w
    ctrl.Height = h

    ' Set caption if provided
    If Len(cap) > 0 Then
        On Error Resume Next
        ctrl.Caption = cap
        Err.Clear
        On Error GoTo 0
    End If
End Sub

'===============================================================================
' ParseSetDirective - Sets a property on a design-time control
'   Format: ControlName,Property,Value
'   Supports: Font.Bold, Font.Name, Font.Size, ForeColor, Style, Default,
'             Cancel, WordWrap, ColumnCount, ColumnWidths, TextAlign, Visible,
'             Value, Enabled
'===============================================================================
Private Sub ParseSetDirective(designer As Object, ByVal directive As String, ByVal formName As String)
    Dim parts() As String
    parts = Split(directive, ",")
    If UBound(parts) < 2 Then Exit Sub

    Dim ctrlName As String, propName As String, propVal As String
    ctrlName = Trim(parts(0))
    propName = Trim(parts(1))
    ' Value may contain commas (e.g. ColumnWidths)
    propVal = Trim(parts(2))
    Dim j As Long
    For j = 3 To UBound(parts)
        propVal = propVal & "," & parts(j)
    Next j

    ' Find the control
    Dim ctrl As Object
    Set ctrl = FindDesignerControl(designer, ctrlName)
    If ctrl Is Nothing Then
        m_Errors = m_Errors & "Form " & formName & ": @SET control '" & ctrlName & "' not found" & vbCrLf
        Exit Sub
    End If

    ' Set the property
    On Error Resume Next
    Select Case LCase(propName)
        Case "font.bold"
            ctrl.Font.Bold = CBool(propVal)
        Case "font.name"
            ctrl.Font.Name = propVal
        Case "font.size"
            ctrl.Font.Size = CSng(propVal)
        Case "forecolor"
            ctrl.ForeColor = CLng(propVal)
        Case "style"
            ctrl.Style = CLng(propVal)
        Case "default"
            ctrl.Default = CBool(propVal)
        Case "cancel"
            ctrl.Cancel = CBool(propVal)
        Case "wordwrap"
            ctrl.WordWrap = CBool(propVal)
        Case "columncount"
            ctrl.ColumnCount = CLng(propVal)
        Case "columnwidths"
            ctrl.ColumnWidths = propVal
        Case "textalign"
            ctrl.TextAlign = CLng(propVal)
        Case "visible"
            ctrl.Visible = CBool(propVal)
        Case "value"
            ctrl.Value = CBool(propVal)
        Case "enabled"
            ctrl.Enabled = CBool(propVal)
        Case Else
            m_Errors = m_Errors & "Form " & formName & ": Unknown property '" & propName & _
                "' for " & ctrlName & vbCrLf
    End Select
    If Err.Number <> 0 Then
        m_Errors = m_Errors & "Form " & formName & ": @SET " & ctrlName & "." & propName & _
            " failed - " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' FindDesignerControl - Finds a control by name in the Designer or nested frames
'===============================================================================
Private Function FindDesignerControl(designer As Object, ByVal ctrlName As String) As Object
    On Error Resume Next
    ' Try form level first
    Set FindDesignerControl = designer.Controls(ctrlName)
    If Not FindDesignerControl Is Nothing Then Exit Function

    ' Search inside frames
    Dim ctrl As Object
    For Each ctrl In designer.Controls
        If TypeName(ctrl) = "Frame" Then
            Set FindDesignerControl = ctrl.Controls(ctrlName)
            If Not FindDesignerControl Is Nothing Then Exit Function
        End If
    Next ctrl
    On Error GoTo 0
End Function

'===============================================================================
' GetControlProgId - Maps type name to MSForms ProgID
'===============================================================================
Private Function GetControlProgId(ByVal typeName As String) As String
    Select Case LCase(typeName)
        Case "label":         GetControlProgId = "Forms.Label.1"
        Case "textbox":       GetControlProgId = "Forms.TextBox.1"
        Case "combobox":      GetControlProgId = "Forms.ComboBox.1"
        Case "commandbutton": GetControlProgId = "Forms.CommandButton.1"
        Case "listbox":       GetControlProgId = "Forms.ListBox.1"
        Case "frame":         GetControlProgId = "Forms.Frame.1"
        Case "optionbutton":  GetControlProgId = "Forms.OptionButton.1"
        Case "checkbox":      GetControlProgId = "Forms.CheckBox.1"
        Case Else:            GetControlProgId = ""
    End Select
End Function

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
    Dim inBeginBlock As Boolean
    inCode = False
    inBeginBlock = False
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim tl As String
        tl = Trim(lines(i))
        If Not inCode Then
            ' Track BEGIN...END blocks and skip everything inside
            If Left(tl, 5) = "BEGIN" Or Left(tl, 5) = "Begin" Then
                inBeginBlock = True
                GoTo SkipLine
            End If
            If inBeginBlock Then
                If tl = "END" Or tl = "End" Then inBeginBlock = False
                GoTo SkipLine
            End If
            If Left(tl, 9) = "Attribute" Then GoTo SkipLine
            If Left(tl, 7) = "VERSION" Then GoTo SkipLine
            If tl = "" Then GoTo SkipLine
            inCode = True
        End If
        result = result & lines(i) & vbCrLf
SkipLine:
    Next i
    StripHeaderLines = result
End Function
