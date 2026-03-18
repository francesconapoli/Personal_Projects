'===============================================================================
' Form: frmModelWindow
' Purpose: Shows all defined inputs, outputs, and correlations
'          Mirrors @RISK's Model Window
'===============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Model Window"
    Me.Width = 450
    Me.Height = 500

    ' Scan workbook
    ScanWorkbookForFunctions

    ' Inputs section
    Dim lblInputs As MSForms.Label
    Set lblInputs = Me.Controls.Add("Forms.Label.1", "lblInputs")
    With lblInputs
        .Caption = "Inputs (" & g_InputCount & " found):"
        .Left = 12: .Top = 8: .Width = 200: .Height = 15
        .Font.Bold = True
        .Font.Size = 10
    End With

    Dim lstInputs As MSForms.ListBox
    Set lstInputs = Me.Controls.Add("Forms.ListBox.1", "lstInputs")
    With lstInputs
        .Left = 12: .Top = 26: .Width = 410: .Height = 140
        .ColumnCount = 3
        .ColumnWidths = "100;100;180"
    End With

    ' Populate inputs
    Dim i As Long
    For i = 0 To g_InputCount - 1
        lstInputs.AddItem
        lstInputs.List(lstInputs.ListCount - 1, 0) = g_Inputs(i).SheetName & "!" & g_Inputs(i).CellAddress
        lstInputs.List(lstInputs.ListCount - 1, 1) = GetDistName(g_Inputs(i).DistType)
        lstInputs.List(lstInputs.ListCount - 1, 2) = g_Inputs(i).InputName
    Next i

    ' Outputs section
    Dim lblOutputs As MSForms.Label
    Set lblOutputs = Me.Controls.Add("Forms.Label.1", "lblOutputs")
    With lblOutputs
        .Caption = "Outputs (" & g_OutputCount & " found):"
        .Left = 12: .Top = 176: .Width = 200: .Height = 15
        .Font.Bold = True
        .Font.Size = 10
    End With

    Dim lstOutputs As MSForms.ListBox
    Set lstOutputs = Me.Controls.Add("Forms.ListBox.1", "lstOutputs")
    With lstOutputs
        .Left = 12: .Top = 194: .Width = 410: .Height = 100
        .ColumnCount = 2
        .ColumnWidths = "150;250"
    End With

    For i = 0 To g_OutputCount - 1
        lstOutputs.AddItem
        lstOutputs.List(lstOutputs.ListCount - 1, 0) = g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress
        lstOutputs.List(lstOutputs.ListCount - 1, 1) = g_Outputs(i).OutputName
    Next i

    ' Correlations section
    Dim lblCorr As MSForms.Label
    Set lblCorr = Me.Controls.Add("Forms.Label.1", "lblCorrelations")
    With lblCorr
        .Caption = "Correlations (" & g_CorrelationCount & " defined):"
        .Left = 12: .Top = 305: .Width = 200: .Height = 15
        .Font.Bold = True
        .Font.Size = 10
    End With

    Dim lstCorr As MSForms.ListBox
    Set lstCorr = Me.Controls.Add("Forms.ListBox.1", "lstCorrelations")
    With lstCorr
        .Left = 12: .Top = 323: .Width = 410: .Height = 80
        .ColumnCount = 3
        .ColumnWidths = "150;150;60"
    End With

    For i = 0 To g_CorrelationCount - 1
        lstCorr.AddItem
        lstCorr.List(lstCorr.ListCount - 1, 0) = g_Inputs(g_Correlations(i).Input1Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 1) = g_Inputs(g_Correlations(i).Input2Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 2) = Format(g_Correlations(i).Coefficient, "0.00")
    Next i

    ' Refresh button
    Dim btnRefresh As MSForms.CommandButton
    Set btnRefresh = Me.Controls.Add("Forms.CommandButton.1", "btnRefresh")
    With btnRefresh
        .Caption = "Refresh"
        .Left = 250: .Top = 420: .Width = 80: .Height = 28
    End With

    ' Close button
    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = 340: .Top = 420: .Width = 80: .Height = 28
        .Cancel = True
    End With
End Sub

Private Function GetDistName(ByVal dt As DistributionType) As String
    Select Case dt
        Case dtNormal: GetDistName = "Normal"
        Case dtUniform: GetDistName = "Uniform"
        Case dtTriangular: GetDistName = "Triangular"
        Case dtLognormal: GetDistName = "Lognormal"
        Case dtExponential: GetDistName = "Exponential"
        Case dtBeta: GetDistName = "Beta"
        Case dtBetaGeneral: GetDistName = "Beta General"
        Case dtGamma: GetDistName = "Gamma"
        Case dtWeibull: GetDistName = "Weibull"
        Case dtPERT: GetDistName = "PERT"
        Case dtPareto: GetDistName = "Pareto"
        Case dtLogistic: GetDistName = "Logistic"
        Case dtRayleigh: GetDistName = "Rayleigh"
        Case dtExtValue: GetDistName = "Ext Value"
        Case dtExtValueMin: GetDistName = "Ext Value Min"
        Case dtChiSquared: GetDistName = "Chi-Squared"
        Case dtStudentT: GetDistName = "Student-t"
        Case dtBernoulli: GetDistName = "Bernoulli"
        Case dtBinomial: GetDistName = "Binomial"
        Case dtPoisson: GetDistName = "Poisson"
        Case dtNegBinomial: GetDistName = "Neg Binomial"
        Case dtDiscrete: GetDistName = "Discrete"
        Case dtDUniform: GetDistName = "D Uniform"
        Case dtIntUniform: GetDistName = "Int Uniform"
        Case dtCumul: GetDistName = "Cumulative"
        Case dtHistogrm: GetDistName = "Histogram"
        Case Else: GetDistName = "Unknown"
    End Select
End Function

Private Sub btnRefresh_Click()
    Unload Me
    frmModelWindow.Show vbModeless
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub
