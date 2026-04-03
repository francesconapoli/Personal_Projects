'===============================================================================
' Form: frmSettings
' Purpose: Simulation settings dialog (mirrors @RISK Settings)
'===============================================================================
'@FORM:400,460,Simulation Settings
'@CTRL:Label,lblIterations,,12,18,120,15,Number of Iterations:
'@CTRL:TextBox,txtIterations,,180,15,100,20
'@CTRL:Label,lblSimulations,,12,48,120,15,Number of Simulations:
'@CTRL:TextBox,txtSimulations,,180,45,100,20
'@CTRL:Frame,fraSampling,,12,78,360,60,Sampling Method
'@CTRL:OptionButton,optMonteCarlo,fraSampling,12,18,120,18,Monte Carlo
'@CTRL:OptionButton,optLatinHypercube,fraSampling,150,18,200,18,Latin Hypercube (Recommended)
'@CTRL:Frame,fraRandom,,12,150,360,70,Random Number Generator
'@CTRL:CheckBox,chkFixedSeed,fraRandom,12,18,120,18,Use Fixed Seed
'@CTRL:TextBox,txtSeed,fraRandom,180,16,100,20
'@CTRL:Frame,fraConvergence,,12,232,360,120,Convergence Monitoring
'@CTRL:CheckBox,chkConvergence,fraConvergence,12,18,180,18,Enable Convergence Testing
'@CTRL:Label,lblTolerance,fraConvergence,12,44,80,15,Tolerance (%):
'@CTRL:TextBox,txtTolerance,fraConvergence,130,42,60,20
'@CTRL:Label,lblConfidence,fraConvergence,12,70,80,15,Confidence (%):
'@CTRL:TextBox,txtConfidence,fraConvergence,130,68,60,20
'@CTRL:Label,lblTestEvery,fraConvergence,210,44,80,15,Test Every N iter:
'@CTRL:TextBox,txtTestEvery,fraConvergence,300,42,50,20
'@CTRL:Label,lblUpdateFreq,,12,365,150,15,Status Update Every N iter:
'@CTRL:TextBox,txtUpdateFreq,,180,363,60,20
'@CTRL:CommandButton,btnOK,,180,400,80,28,OK
'@SET:btnOK,Default,True
'@CTRL:CommandButton,btnCancel,,270,400,80,28,Cancel
'@SET:btnCancel,Cancel,True
'@END
Option Explicit

Private Sub UserForm_Initialize()
    Me.Controls("txtIterations").Text = CStr(g_Settings.Iterations)
    Me.Controls("txtSimulations").Text = CStr(g_Settings.Simulations)

    Me.Controls("fraSampling").Controls("optMonteCarlo").Value = (g_Settings.SampMethod = smMonteCarlo)
    Me.Controls("fraSampling").Controls("optLatinHypercube").Value = (g_Settings.SampMethod = smLatinHypercube)

    Me.Controls("fraRandom").Controls("chkFixedSeed").Value = g_Settings.UseFixedSeed
    Me.Controls("fraRandom").Controls("txtSeed").Text = CStr(g_Settings.RandomSeed)
    Me.Controls("fraRandom").Controls("txtSeed").Enabled = g_Settings.UseFixedSeed

    Me.Controls("fraConvergence").Controls("chkConvergence").Value = g_Settings.ConvergenceEnabled
    Me.Controls("fraConvergence").Controls("txtTolerance").Text = CStr(g_Settings.ConvergenceTolerance * 100)
    Me.Controls("fraConvergence").Controls("txtConfidence").Text = CStr(g_Settings.ConvergenceConfidence * 100)
    Me.Controls("fraConvergence").Controls("txtTestEvery").Text = CStr(g_Settings.ConvergenceTestEvery)

    Me.Controls("txtUpdateFreq").Text = CStr(g_Settings.UpdateFrequency)
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
