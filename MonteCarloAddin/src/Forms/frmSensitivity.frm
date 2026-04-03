'===============================================================================
' Form: frmSensitivity
' Purpose: Sensitivity analysis configuration and results
'===============================================================================
'@FORM:520,450,Sensitivity Analysis
'@CTRL:Label,lblOutput,,12,12,90,15,Analyze Output:
'@SET:lblOutput,Font.Bold,True
'@CTRL:ComboBox,cboOutput,,110,10,250,20
'@SET:cboOutput,Style,2
'@CTRL:Frame,fraMethod,,12,40,480,90,Sensitivity Method
'@CTRL:OptionButton,optMethod0,fraMethod,12,18,220,18,Regression Coefficients
'@CTRL:OptionButton,optMethod1,fraMethod,242,18,220,18,Rank Correlation (Spearman)
'@SET:optMethod1,Value,True
'@CTRL:OptionButton,optMethod2,fraMethod,12,46,220,18,Contribution to Variance
'@CTRL:OptionButton,optMethod3,fraMethod,242,46,220,18,Change in Output Statistic
'@CTRL:CommandButton,btnRun,,370,140,120,28,Run Analysis
'@CTRL:Label,lblResults,,12,145,60,15,Results:
'@SET:lblResults,Font.Bold,True
'@CTRL:Label,lblRSquared,,80,145,200,15
'@CTRL:ListBox,lstResults,,12,170,480,160
'@SET:lstResults,ColumnCount,4
'@SET:lstResults,ColumnWidths,40;150;100;100
'@CTRL:CommandButton,btnTornado,,12,340,110,28,Tornado Chart
'@CTRL:CommandButton,btnSpider,,130,340,110,28,Spider Plot
'@CTRL:CommandButton,btnExport,,248,340,110,28,Export to Sheet
'@CTRL:CommandButton,btnClose,,410,380,80,28,Close
'@SET:btnClose,Cancel,True
'@END
Option Explicit

Private Sub UserForm_Initialize()
    ' Populate output combo
    Dim cboOutput As MSForms.ComboBox
    Set cboOutput = Me.Controls("cboOutput")
    Dim i As Long
    For i = 0 To g_OutputCount - 1
        cboOutput.AddItem g_Outputs(i).OutputName
    Next i
    If g_OutputCount > 0 Then cboOutput.ListIndex = 0
End Sub

Private Sub btnRun_Click()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboOutput")
    If cbo.ListIndex < 0 Then
        MsgBox "Select an output.", vbExclamation
        Exit Sub
    End If

    ' Determine method
    Dim method As SensitivityMethod
    Dim fra As MSForms.Frame
    Set fra = Me.Controls("fraMethod")
    If fra.Controls("optMethod0").Value Then method = snRegression
    If fra.Controls("optMethod1").Value Then method = snRankCorrelation
    If fra.Controls("optMethod2").Value Then method = snContributionToVariance
    If fra.Controls("optMethod3").Value Then method = snChangeInStatistic

    ' Run analysis
    RunSensitivityAnalysis cbo.ListIndex, method

    ' Display results
    Me.Controls("lblRSquared").Caption = "R-Squared = " & Format(g_SensRSquared, "0.0000")

    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstResults")
    lst.Clear

    ' Header
    lst.AddItem
    lst.List(0, 0) = "Rank"
    lst.List(0, 1) = "Input"
    lst.List(0, 2) = "Coefficient"
    lst.List(0, 3) = "Contribution %"

    Dim i As Long
    For i = 0 To g_SensResultCount - 1
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = CStr(i + 1)
        lst.List(lst.ListCount - 1, 1) = g_SensResults(i).InputName
        lst.List(lst.ListCount - 1, 2) = Format(g_SensResults(i).Coefficient, "0.0000")
        lst.List(lst.ListCount - 1, 3) = Format(g_SensResults(i).ContributionPct, "0.00") & "%"
    Next i
End Sub

Private Sub btnTornado_Click()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboOutput")
    If cbo.ListIndex >= 0 Then CreateTornadoDiagram cbo.ListIndex
End Sub

Private Sub btnSpider_Click()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboOutput")
    If cbo.ListIndex >= 0 Then CreateSpiderPlot cbo.ListIndex
End Sub

Private Sub btnExport_Click()
    If g_SensResultCount = 0 Then
        MsgBox "Run analysis first.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateResultsSheet()
    Dim startRow As Long
    startRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 3

    ws.Cells(startRow, 1).Value = "Sensitivity Analysis Results"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 12
    ws.Cells(startRow + 1, 1).Value = "R-Squared:"
    ws.Cells(startRow + 1, 2).Value = g_SensRSquared

    ws.Cells(startRow + 3, 1).Value = "Rank"
    ws.Cells(startRow + 3, 2).Value = "Input"
    ws.Cells(startRow + 3, 3).Value = "Coefficient"
    ws.Cells(startRow + 3, 4).Value = "Contribution %"
    ws.Range(ws.Cells(startRow + 3, 1), ws.Cells(startRow + 3, 4)).Font.Bold = True

    Dim i As Long
    For i = 0 To g_SensResultCount - 1
        ws.Cells(startRow + 4 + i, 1).Value = i + 1
        ws.Cells(startRow + 4 + i, 2).Value = g_SensResults(i).InputName
        ws.Cells(startRow + 4 + i, 3).Value = g_SensResults(i).Coefficient
        ws.Cells(startRow + 4 + i, 4).Value = g_SensResults(i).ContributionPct
    Next i

    ws.Columns.AutoFit
    MsgBox "Results exported to 'MC Results' sheet.", vbInformation
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub
