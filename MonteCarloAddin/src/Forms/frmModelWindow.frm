'===============================================================================
' Form: frmModelWindow
' Purpose: Shows all defined inputs, outputs, and correlations
'          Mirrors @RISK's Model Window
'===============================================================================
'@FORM:450,500,Model Window
'@CTRL:Label,lblInputs,,12,8,200,15,Inputs:
'@SET:lblInputs,Font.Bold,True
'@SET:lblInputs,Font.Size,10
'@CTRL:ListBox,lstInputs,,12,26,410,140
'@SET:lstInputs,ColumnCount,3
'@SET:lstInputs,ColumnWidths,100;100;180
'@CTRL:Label,lblOutputs,,12,176,200,15,Outputs:
'@SET:lblOutputs,Font.Bold,True
'@SET:lblOutputs,Font.Size,10
'@CTRL:ListBox,lstOutputs,,12,194,410,100
'@SET:lstOutputs,ColumnCount,2
'@SET:lstOutputs,ColumnWidths,150;250
'@CTRL:Label,lblCorrelations,,12,305,200,15,Correlations:
'@SET:lblCorrelations,Font.Bold,True
'@SET:lblCorrelations,Font.Size,10
'@CTRL:ListBox,lstCorrelations,,12,323,410,80
'@SET:lstCorrelations,ColumnCount,3
'@SET:lstCorrelations,ColumnWidths,150;150;60
'@CTRL:CommandButton,btnRefresh,,250,420,80,28,Refresh
'@CTRL:CommandButton,btnClose,,340,420,80,28,Close
'@SET:btnClose,Cancel,True
'@END
Option Explicit

Private Sub UserForm_Initialize()
    ' Scan workbook
    ScanWorkbookForFunctions

    ' Update section headers with counts
    Me.Controls("lblInputs").Caption = "Inputs (" & g_InputCount & " found):"
    Me.Controls("lblOutputs").Caption = "Outputs (" & g_OutputCount & " found):"
    Me.Controls("lblCorrelations").Caption = "Correlations (" & g_CorrelationCount & " defined):"

    ' Populate inputs
    Dim lstInputs As MSForms.ListBox
    Set lstInputs = Me.Controls("lstInputs")
    Dim i As Long
    For i = 0 To g_InputCount - 1
        lstInputs.AddItem
        lstInputs.List(lstInputs.ListCount - 1, 0) = g_Inputs(i).SheetName & "!" & g_Inputs(i).CellAddress
        lstInputs.List(lstInputs.ListCount - 1, 1) = GetDistName(g_Inputs(i).DistType)
        lstInputs.List(lstInputs.ListCount - 1, 2) = g_Inputs(i).InputName
    Next i

    ' Populate outputs
    Dim lstOutputs As MSForms.ListBox
    Set lstOutputs = Me.Controls("lstOutputs")
    For i = 0 To g_OutputCount - 1
        lstOutputs.AddItem
        lstOutputs.List(lstOutputs.ListCount - 1, 0) = g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress
        lstOutputs.List(lstOutputs.ListCount - 1, 1) = g_Outputs(i).OutputName
    Next i

    ' Populate correlations
    Dim lstCorr As MSForms.ListBox
    Set lstCorr = Me.Controls("lstCorrelations")
    For i = 0 To g_CorrelationCount - 1
        lstCorr.AddItem
        lstCorr.List(lstCorr.ListCount - 1, 0) = g_Inputs(g_Correlations(i).Input1Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 1) = g_Inputs(g_Correlations(i).Input2Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 2) = Format(g_Correlations(i).Coefficient, "0.00")
    Next i
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
