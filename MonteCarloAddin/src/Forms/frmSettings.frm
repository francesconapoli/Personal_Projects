VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings
   Caption         =   "Simulation Settings"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Form: frmSettings
' Purpose: Simulation settings dialog (mirrors @RISK Settings)
'===============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    ' ---- Create controls programmatically ----
    Me.Width = 400
    Me.Height = 460

    ' Iterations
    Dim lblIter As MSForms.Label
    Set lblIter = Me.Controls.Add("Forms.Label.1", "lblIterations")
    With lblIter
        .Caption = "Number of Iterations:"
        .Left = 12: .Top = 18: .Width = 120: .Height = 15
    End With

    Dim txtIter As MSForms.TextBox
    Set txtIter = Me.Controls.Add("Forms.TextBox.1", "txtIterations")
    With txtIter
        .Left = 180: .Top = 15: .Width = 100: .Height = 20
        .Text = CStr(g_Settings.Iterations)
    End With

    ' Simulations
    Dim lblSim As MSForms.Label
    Set lblSim = Me.Controls.Add("Forms.Label.1", "lblSimulations")
    With lblSim
        .Caption = "Number of Simulations:"
        .Left = 12: .Top = 48: .Width = 120: .Height = 15
    End With

    Dim txtSim As MSForms.TextBox
    Set txtSim = Me.Controls.Add("Forms.TextBox.1", "txtSimulations")
    With txtSim
        .Left = 180: .Top = 45: .Width = 100: .Height = 20
        .Text = CStr(g_Settings.Simulations)
    End With

    ' Sampling Method
    Dim fraSampling As MSForms.Frame
    Set fraSampling = Me.Controls.Add("Forms.Frame.1", "fraSampling")
    With fraSampling
        .Caption = "Sampling Method"
        .Left = 12: .Top = 78: .Width = 360: .Height = 60
    End With

    Dim optMC As MSForms.OptionButton
    Set optMC = fraSampling.Controls.Add("Forms.OptionButton.1", "optMonteCarlo")
    With optMC
        .Caption = "Monte Carlo"
        .Left = 12: .Top = 18: .Width = 120: .Height = 18
        .Value = (g_Settings.SampMethod = smMonteCarlo)
    End With

    Dim optLHS As MSForms.OptionButton
    Set optLHS = fraSampling.Controls.Add("Forms.OptionButton.1", "optLatinHypercube")
    With optLHS
        .Caption = "Latin Hypercube (Recommended)"
        .Left = 150: .Top = 18: .Width = 200: .Height = 18
        .Value = (g_Settings.SampMethod = smLatinHypercube)
    End With

    ' Random Seed
    Dim fraRandom As MSForms.Frame
    Set fraRandom = Me.Controls.Add("Forms.Frame.1", "fraRandom")
    With fraRandom
        .Caption = "Random Number Generator"
        .Left = 12: .Top = 150: .Width = 360: .Height = 70
    End With

    Dim chkFixed As MSForms.CheckBox
    Set chkFixed = fraRandom.Controls.Add("Forms.CheckBox.1", "chkFixedSeed")
    With chkFixed
        .Caption = "Use Fixed Seed"
        .Left = 12: .Top = 18: .Width = 120: .Height = 18
        .Value = g_Settings.UseFixedSeed
    End With

    Dim txtSeed As MSForms.TextBox
    Set txtSeed = fraRandom.Controls.Add("Forms.TextBox.1", "txtSeed")
    With txtSeed
        .Left = 180: .Top = 16: .Width = 100: .Height = 20
        .Text = CStr(g_Settings.RandomSeed)
        .Enabled = g_Settings.UseFixedSeed
    End With

    ' Convergence
    Dim fraConv As MSForms.Frame
    Set fraConv = Me.Controls.Add("Forms.Frame.1", "fraConvergence")
    With fraConv
        .Caption = "Convergence Monitoring"
        .Left = 12: .Top = 232: .Width = 360: .Height = 120
    End With

    Dim chkConv As MSForms.CheckBox
    Set chkConv = fraConv.Controls.Add("Forms.CheckBox.1", "chkConvergence")
    With chkConv
        .Caption = "Enable Convergence Testing"
        .Left = 12: .Top = 18: .Width = 180: .Height = 18
        .Value = g_Settings.ConvergenceEnabled
    End With

    Dim lblTol As MSForms.Label
    Set lblTol = fraConv.Controls.Add("Forms.Label.1", "lblTolerance")
    With lblTol
        .Caption = "Tolerance (%):"
        .Left = 12: .Top = 44: .Width = 80: .Height = 15
    End With

    Dim txtTol As MSForms.TextBox
    Set txtTol = fraConv.Controls.Add("Forms.TextBox.1", "txtTolerance")
    With txtTol
        .Left = 130: .Top = 42: .Width = 60: .Height = 20
        .Text = CStr(g_Settings.ConvergenceTolerance * 100)
    End With

    Dim lblConf As MSForms.Label
    Set lblConf = fraConv.Controls.Add("Forms.Label.1", "lblConfidence")
    With lblConf
        .Caption = "Confidence (%):"
        .Left = 12: .Top = 70: .Width = 80: .Height = 15
    End With

    Dim txtConf As MSForms.TextBox
    Set txtConf = fraConv.Controls.Add("Forms.TextBox.1", "txtConfidence")
    With txtConf
        .Left = 130: .Top = 68: .Width = 60: .Height = 20
        .Text = CStr(g_Settings.ConvergenceConfidence * 100)
    End With

    Dim lblEvery As MSForms.Label
    Set lblEvery = fraConv.Controls.Add("Forms.Label.1", "lblTestEvery")
    With lblEvery
        .Caption = "Test Every N iter:"
        .Left = 210: .Top = 44: .Width = 80: .Height = 15
    End With

    Dim txtEvery As MSForms.TextBox
    Set txtEvery = fraConv.Controls.Add("Forms.TextBox.1", "txtTestEvery")
    With txtEvery
        .Left = 300: .Top = 42: .Width = 50: .Height = 20
        .Text = CStr(g_Settings.ConvergenceTestEvery)
    End With

    ' Update Frequency
    Dim lblUpdate As MSForms.Label
    Set lblUpdate = Me.Controls.Add("Forms.Label.1", "lblUpdateFreq")
    With lblUpdate
        .Caption = "Status Update Every N iter:"
        .Left = 12: .Top = 365: .Width = 150: .Height = 15
    End With

    Dim txtUpdate As MSForms.TextBox
    Set txtUpdate = Me.Controls.Add("Forms.TextBox.1", "txtUpdateFreq")
    With txtUpdate
        .Left = 180: .Top = 363: .Width = 60: .Height = 20
        .Text = CStr(g_Settings.UpdateFrequency)
    End With

    ' Buttons
    Dim btnOK As MSForms.CommandButton
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With btnOK
        .Caption = "OK"
        .Left = 180: .Top = 400: .Width = 80: .Height = 28
        .Default = True
    End With

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Caption = "Cancel"
        .Left = 270: .Top = 400: .Width = 80: .Height = 28
        .Cancel = True
    End With
End Sub

Private Sub btnOK_Click()
    On Error GoTo ValidationError

    g_Settings.Iterations = CLng(Me.Controls("txtIterations").Text)
    g_Settings.Simulations = CLng(Me.Controls("txtSimulations").Text)

    If Me.Controls("fraSampling").Controls("optMonteCarlo").Value Then
        g_Settings.SampMethod = smMonteCarlo
    Else
        g_Settings.SampMethod = smLatinHypercube
    End If

    g_Settings.UseFixedSeed = Me.Controls("fraRandom").Controls("chkFixedSeed").Value
    g_Settings.RandomSeed = CLng(Me.Controls("fraRandom").Controls("txtSeed").Text)

    g_Settings.ConvergenceEnabled = Me.Controls("fraConvergence").Controls("chkConvergence").Value
    g_Settings.ConvergenceTolerance = CDbl(Me.Controls("fraConvergence").Controls("txtTolerance").Text) / 100#
    g_Settings.ConvergenceConfidence = CDbl(Me.Controls("fraConvergence").Controls("txtConfidence").Text) / 100#
    g_Settings.ConvergenceTestEvery = CLng(Me.Controls("fraConvergence").Controls("txtTestEvery").Text)

    g_Settings.UpdateFrequency = CLng(Me.Controls("txtUpdateFreq").Text)

    ' Validation
    If g_Settings.Iterations < 1 Or g_Settings.Iterations > 1000000 Then
        MsgBox "Iterations must be between 1 and 1,000,000.", vbExclamation
        Exit Sub
    End If

    Me.Hide
    Exit Sub

ValidationError:
    MsgBox "Invalid input: " & Err.Description, vbExclamation
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub chkFixedSeed_Click()
    Me.Controls("fraRandom").Controls("txtSeed").Enabled = _
        Me.Controls("fraRandom").Controls("chkFixedSeed").Value
End Sub
