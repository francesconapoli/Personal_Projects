VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddInput
   Caption         =   "Add Distribution Input"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   OleObjectBlob   =   "frmAddInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Form: frmAddInput
' Purpose: Add distribution function to selected cell
'          Mirrors @RISK's "Define Distribution" dialog
'===============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    Me.Width = 480
    Me.Height = 440

    ' Target cell
    Dim lblCell As MSForms.Label
    Set lblCell = Me.Controls.Add("Forms.Label.1", "lblCell")
    With lblCell
        .Caption = "Target Cell:"
        .Left = 12: .Top = 12: .Width = 70: .Height = 15
        .Font.Bold = True
    End With

    Dim txtCell As MSForms.TextBox
    Set txtCell = Me.Controls.Add("Forms.TextBox.1", "txtCell")
    With txtCell
        .Left = 90: .Top = 10: .Width = 120: .Height = 20
        If Not ActiveCell Is Nothing Then .Text = ActiveCell.Address
    End With

    ' Distribution Category
    Dim lblCat As MSForms.Label
    Set lblCat = Me.Controls.Add("Forms.Label.1", "lblCategory")
    With lblCat
        .Caption = "Category:"
        .Left = 12: .Top = 42: .Width = 70: .Height = 15
    End With

    Dim cboCat As MSForms.ComboBox
    Set cboCat = Me.Controls.Add("Forms.ComboBox.1", "cboCategory")
    With cboCat
        .Left = 90: .Top = 40: .Width = 160: .Height = 20
        .Style = fmStyleDropDownList
        .AddItem "Continuous"
        .AddItem "Discrete"
        .AddItem "Empirical"
        .ListIndex = 0
    End With

    ' Distribution Type
    Dim lblDist As MSForms.Label
    Set lblDist = Me.Controls.Add("Forms.Label.1", "lblDistType")
    With lblDist
        .Caption = "Distribution:"
        .Left = 12: .Top = 72: .Width = 70: .Height = 15
    End With

    Dim cboDist As MSForms.ComboBox
    Set cboDist = Me.Controls.Add("Forms.ComboBox.1", "cboDistribution")
    With cboDist
        .Left = 90: .Top = 70: .Width = 160: .Height = 20
        .Style = fmStyleDropDownList
    End With

    ' Parameter frame
    Dim fraParams As MSForms.Frame
    Set fraParams = Me.Controls.Add("Forms.Frame.1", "fraParams")
    With fraParams
        .Caption = "Parameters"
        .Left = 12: .Top = 100: .Width = 440: .Height = 180
    End With

    ' Parameter labels and textboxes (up to 4 parameters)
    Dim paramLabels As Variant
    paramLabels = Array("Parameter 1:", "Parameter 2:", "Parameter 3:", "Parameter 4:")
    Dim j As Long
    For j = 0 To 3
        Dim pl As MSForms.Label
        Set pl = fraParams.Controls.Add("Forms.Label.1", "lblParam" & j)
        With pl
            .Caption = CStr(paramLabels(j))
            .Left = 12: .Top = 20 + j * 35: .Width = 100: .Height = 15
        End With

        Dim pt As MSForms.TextBox
        Set pt = fraParams.Controls.Add("Forms.TextBox.1", "txtParam" & j)
        With pt
            .Left = 120: .Top = 18 + j * 35: .Width = 100: .Height = 20
            .Text = ""
        End With

        Dim pd As MSForms.Label
        Set pd = fraParams.Controls.Add("Forms.Label.1", "lblParamDesc" & j)
        With pd
            .Caption = ""
            .Left = 230: .Top = 20 + j * 35: .Width = 200: .Height = 15
            .ForeColor = RGB(100, 100, 100)
        End With
    Next j

    ' Name
    Dim lblName As MSForms.Label
    Set lblName = Me.Controls.Add("Forms.Label.1", "lblInputName")
    With lblName
        .Caption = "Input Name:"
        .Left = 12: .Top = 290: .Width = 70: .Height = 15
    End With

    Dim txtName As MSForms.TextBox
    Set txtName = Me.Controls.Add("Forms.TextBox.1", "txtInputName")
    With txtName
        .Left = 90: .Top = 288: .Width = 160: .Height = 20
        .Text = ""
    End With

    ' Preview
    Dim lblPreview As MSForms.Label
    Set lblPreview = Me.Controls.Add("Forms.Label.1", "lblPreview")
    With lblPreview
        .Caption = "Formula Preview:"
        .Left = 12: .Top = 320: .Width = 100: .Height = 15
        .Font.Bold = True
    End With

    Dim lblFormula As MSForms.Label
    Set lblFormula = Me.Controls.Add("Forms.Label.1", "lblFormulaPreview")
    With lblFormula
        .Caption = "=MCNormal(0, 1)"
        .Left = 12: .Top = 338: .Width = 440: .Height = 20
        .Font.Name = "Consolas"
        .Font.Size = 11
        .ForeColor = RGB(0, 0, 180)
    End With

    ' Buttons
    Dim btnInsert As MSForms.CommandButton
    Set btnInsert = Me.Controls.Add("Forms.CommandButton.1", "btnInsert")
    With btnInsert
        .Caption = "Insert"
        .Left = 250: .Top = 370: .Width = 90: .Height = 28
        .Default = True
    End With

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Caption = "Cancel"
        .Left = 350: .Top = 370: .Width = 90: .Height = 28
        .Cancel = True
    End With

    ' Load distributions
    LoadDistributions
End Sub

Private Sub LoadDistributions()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboDistribution")
    cbo.Clear

    Select Case Me.Controls("cboCategory").ListIndex
        Case 0 ' Continuous
            cbo.AddItem "Normal"
            cbo.AddItem "Uniform"
            cbo.AddItem "Triangular"
            cbo.AddItem "Lognormal"
            cbo.AddItem "Exponential"
            cbo.AddItem "Beta"
            cbo.AddItem "Beta General"
            cbo.AddItem "Gamma"
            cbo.AddItem "Weibull"
            cbo.AddItem "PERT"
            cbo.AddItem "Pareto"
            cbo.AddItem "Logistic"
            cbo.AddItem "Rayleigh"
            cbo.AddItem "Extreme Value"
            cbo.AddItem "Extreme Value Min"
            cbo.AddItem "Chi-Squared"
            cbo.AddItem "Student's t"
        Case 1 ' Discrete
            cbo.AddItem "Binomial"
            cbo.AddItem "Poisson"
            cbo.AddItem "Negative Binomial"
            cbo.AddItem "Discrete"
            cbo.AddItem "Discrete Uniform"
            cbo.AddItem "Integer Uniform"
        Case 2 ' Empirical
            cbo.AddItem "Cumulative"
            cbo.AddItem "Cumulative Descending"
            cbo.AddItem "Histogram"
    End Select

    If cbo.ListCount > 0 Then cbo.ListIndex = 0
End Sub

Private Sub cboCategory_Change()
    LoadDistributions
End Sub

Private Sub cboDistribution_Change()
    UpdateParamLabels
    UpdateFormulaPreview
End Sub

Private Sub UpdateParamLabels()
    Dim fra As MSForms.Frame
    Set fra = Me.Controls("fraParams")
    Dim distName As String
    distName = Me.Controls("cboDistribution").Text

    ' Reset all
    Dim j As Long
    For j = 0 To 3
        fra.Controls("lblParam" & j).Visible = False
        fra.Controls("txtParam" & j).Visible = False
        fra.Controls("lblParamDesc" & j).Visible = False
    Next j

    Select Case distName
        Case "Normal"
            ShowParam fra, 0, "Mean:", "0", "Center of distribution"
            ShowParam fra, 1, "Std Dev:", "1", "Spread of distribution"
        Case "Uniform"
            ShowParam fra, 0, "Minimum:", "0", "Lower bound"
            ShowParam fra, 1, "Maximum:", "1", "Upper bound"
        Case "Triangular"
            ShowParam fra, 0, "Minimum:", "0", "Lower bound"
            ShowParam fra, 1, "Most Likely:", "0.5", "Peak/mode"
            ShowParam fra, 2, "Maximum:", "1", "Upper bound"
        Case "Lognormal"
            ShowParam fra, 0, "Mean:", "1", "Arithmetic mean"
            ShowParam fra, 1, "Std Dev:", "0.5", "Arithmetic std dev"
        Case "Exponential"
            ShowParam fra, 0, "Beta (mean):", "1", "Mean/scale parameter"
        Case "Beta"
            ShowParam fra, 0, "Alpha 1:", "2", "Shape parameter 1"
            ShowParam fra, 1, "Alpha 2:", "5", "Shape parameter 2"
        Case "Beta General"
            ShowParam fra, 0, "Alpha 1:", "2", "Shape parameter 1"
            ShowParam fra, 1, "Alpha 2:", "5", "Shape parameter 2"
            ShowParam fra, 2, "Minimum:", "0", "Lower bound"
            ShowParam fra, 3, "Maximum:", "1", "Upper bound"
        Case "Gamma"
            ShowParam fra, 0, "Alpha (shape):", "2", "Shape parameter"
            ShowParam fra, 1, "Beta (scale):", "1", "Scale parameter"
        Case "Weibull"
            ShowParam fra, 0, "Shape:", "2", "Shape parameter"
            ShowParam fra, 1, "Scale:", "1", "Scale parameter"
        Case "PERT"
            ShowParam fra, 0, "Minimum:", "0", "Lower bound"
            ShowParam fra, 1, "Most Likely:", "0.5", "Most likely value"
            ShowParam fra, 2, "Maximum:", "1", "Upper bound"
        Case "Pareto"
            ShowParam fra, 0, "Shape (theta):", "2", "Shape parameter"
            ShowParam fra, 1, "Scale (xm):", "1", "Scale/location"
        Case "Logistic"
            ShowParam fra, 0, "Mean:", "0", "Location parameter"
            ShowParam fra, 1, "Scale:", "1", "Scale parameter"
        Case "Rayleigh"
            ShowParam fra, 0, "Scale:", "1", "Scale parameter"
        Case "Extreme Value", "Extreme Value Min"
            ShowParam fra, 0, "Alpha (location):", "0", "Location parameter"
            ShowParam fra, 1, "Beta (scale):", "1", "Scale parameter"
        Case "Chi-Squared"
            ShowParam fra, 0, "Degrees of Freedom:", "5", "DF > 0"
        Case "Student's t"
            ShowParam fra, 0, "Degrees of Freedom:", "10", "DF > 0"
        Case "Binomial"
            ShowParam fra, 0, "N (trials):", "10", "Number of trials"
            ShowParam fra, 1, "P (probability):", "0.5", "Success probability"
        Case "Poisson"
            ShowParam fra, 0, "Lambda (mean):", "5", "Expected count"
        Case "Negative Binomial"
            ShowParam fra, 0, "S (successes):", "5", "Target successes"
            ShowParam fra, 1, "P (probability):", "0.5", "Success probability"
        Case "Integer Uniform"
            ShowParam fra, 0, "Minimum:", "1", "Lower integer"
            ShowParam fra, 1, "Maximum:", "10", "Upper integer"
        Case "Discrete", "Discrete Uniform", "Cumulative", "Cumulative Descending", "Histogram"
            ShowParam fra, 0, "Values Range:", "", "Select cell range"
            ShowParam fra, 1, "Probabilities:", "", "Select cell range"
    End Select
End Sub

Private Sub ShowParam(ByRef fra As MSForms.Frame, ByVal idx As Long, _
    ByVal labelText As String, ByVal defaultVal As String, ByVal descText As String)
    fra.Controls("lblParam" & idx).Caption = labelText
    fra.Controls("lblParam" & idx).Visible = True
    If fra.Controls("txtParam" & idx).Text = "" Then
        fra.Controls("txtParam" & idx).Text = defaultVal
    End If
    fra.Controls("txtParam" & idx).Visible = True
    fra.Controls("lblParamDesc" & idx).Caption = descText
    fra.Controls("lblParamDesc" & idx).Visible = True
End Sub

Private Sub UpdateFormulaPreview()
    Dim distName As String
    distName = Me.Controls("cboDistribution").Text
    Dim fra As MSForms.Frame
    Set fra = Me.Controls("fraParams")
    Dim formula As String

    Dim p0 As String, p1 As String, p2 As String, p3 As String
    p0 = fra.Controls("txtParam0").Text
    p1 = fra.Controls("txtParam1").Text
    p2 = fra.Controls("txtParam2").Text
    p3 = fra.Controls("txtParam3").Text

    Select Case distName
        Case "Normal": formula = "=MCNormal(" & p0 & ", " & p1 & ")"
        Case "Uniform": formula = "=MCUniform(" & p0 & ", " & p1 & ")"
        Case "Triangular": formula = "=MCTriang(" & p0 & ", " & p1 & ", " & p2 & ")"
        Case "Lognormal": formula = "=MCLognorm(" & p0 & ", " & p1 & ")"
        Case "Exponential": formula = "=MCExpon(" & p0 & ")"
        Case "Beta": formula = "=MCBeta(" & p0 & ", " & p1 & ")"
        Case "Beta General": formula = "=MCBetaGeneral(" & p0 & ", " & p1 & ", " & p2 & ", " & p3 & ")"
        Case "Gamma": formula = "=MCGamma(" & p0 & ", " & p1 & ")"
        Case "Weibull": formula = "=MCWeibull(" & p0 & ", " & p1 & ")"
        Case "PERT": formula = "=MCPERT(" & p0 & ", " & p1 & ", " & p2 & ")"
        Case "Pareto": formula = "=MCPareto(" & p0 & ", " & p1 & ")"
        Case "Logistic": formula = "=MCLogistic(" & p0 & ", " & p1 & ")"
        Case "Rayleigh": formula = "=MCRayleigh(" & p0 & ")"
        Case "Extreme Value": formula = "=MCExtValue(" & p0 & ", " & p1 & ")"
        Case "Extreme Value Min": formula = "=MCExtValueMin(" & p0 & ", " & p1 & ")"
        Case "Chi-Squared": formula = "=MCChiSq(" & p0 & ")"
        Case "Student's t": formula = "=MCStudentT(" & p0 & ")"
        Case "Binomial": formula = "=MCBinomial(" & p0 & ", " & p1 & ")"
        Case "Poisson": formula = "=MCPoisson(" & p0 & ")"
        Case "Negative Binomial": formula = "=MCNegBinom(" & p0 & ", " & p1 & ")"
        Case "Integer Uniform": formula = "=MCIntUniform(" & p0 & ", " & p1 & ")"
        Case "Discrete": formula = "=MCDiscrete(" & p0 & ", " & p1 & ")"
        Case "Discrete Uniform": formula = "=MCDUniform(" & p0 & ")"
        Case Else: formula = ""
    End Select

    Me.Controls("lblFormulaPreview").Caption = formula
End Sub

Private Sub txtParam0_Change(): UpdateFormulaPreview: End Sub
Private Sub txtParam1_Change(): UpdateFormulaPreview: End Sub
Private Sub txtParam2_Change(): UpdateFormulaPreview: End Sub
Private Sub txtParam3_Change(): UpdateFormulaPreview: End Sub

Private Sub btnInsert_Click()
    On Error GoTo InsertError

    Dim targetAddr As String
    targetAddr = Me.Controls("txtCell").Text
    If targetAddr = "" Then
        MsgBox "Specify a target cell.", vbExclamation
        Exit Sub
    End If

    Dim formula As String
    formula = Me.Controls("lblFormulaPreview").Caption
    If formula = "" Then
        MsgBox "Select a distribution.", vbExclamation
        Exit Sub
    End If

    Dim rng As Range
    Set rng = ActiveSheet.Range(targetAddr)
    rng.formula = formula

    Me.Hide
    Exit Sub

InsertError:
    MsgBox "Error inserting formula: " & Err.Description, vbCritical
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub
