'===============================================================================
' Form: frmResults
' Purpose: Results browser - shows statistics and charts for selected output
'===============================================================================
'@FORM:680,620,Simulation Results
'@CTRL:Label,lblOutput,,12,12,50,15,Output:
'@SET:lblOutput,Font.Bold,True
'@CTRL:ComboBox,cboOutput,,65,10,250,20
'@SET:cboOutput,Style,2
'@CTRL:Label,lblSimInfo,,330,12,320,15
'@CTRL:Frame,fraStats,,12,40,640,220,Statistics
'@CTRL:Label,lblStatMean,fraStats,12,18,85,15,Mean:
'@SET:lblStatMean,Font.Bold,True
'@CTRL:Label,lblValMean,fraStats,100,18,100,15
'@SET:lblValMean,TextAlign,3
'@CTRL:Label,lblStatStdDev,fraStats,12,40,85,15,Std Deviation:
'@SET:lblStatStdDev,Font.Bold,True
'@CTRL:Label,lblValStdDev,fraStats,100,40,100,15
'@SET:lblValStdDev,TextAlign,3
'@CTRL:Label,lblStatVariance,fraStats,12,62,85,15,Variance:
'@SET:lblStatVariance,Font.Bold,True
'@CTRL:Label,lblValVariance,fraStats,100,62,100,15
'@SET:lblValVariance,TextAlign,3
'@CTRL:Label,lblStatMin,fraStats,12,84,85,15,Minimum:
'@SET:lblStatMin,Font.Bold,True
'@CTRL:Label,lblValMin,fraStats,100,84,100,15
'@SET:lblValMin,TextAlign,3
'@CTRL:Label,lblStatMax,fraStats,12,106,85,15,Maximum:
'@SET:lblStatMax,Font.Bold,True
'@CTRL:Label,lblValMax,fraStats,100,106,100,15
'@SET:lblValMax,TextAlign,3
'@CTRL:Label,lblStatMedian,fraStats,220,18,85,15,Median:
'@SET:lblStatMedian,Font.Bold,True
'@CTRL:Label,lblValMedian,fraStats,310,18,100,15
'@SET:lblValMedian,TextAlign,3
'@CTRL:Label,lblStatMode,fraStats,220,40,85,15,Mode:
'@SET:lblStatMode,Font.Bold,True
'@CTRL:Label,lblValMode,fraStats,310,40,100,15
'@SET:lblValMode,TextAlign,3
'@CTRL:Label,lblStatSkewness,fraStats,220,62,85,15,Skewness:
'@SET:lblStatSkewness,Font.Bold,True
'@CTRL:Label,lblValSkewness,fraStats,310,62,100,15
'@SET:lblValSkewness,TextAlign,3
'@CTRL:Label,lblStatKurtosis,fraStats,220,84,85,15,Kurtosis:
'@SET:lblStatKurtosis,Font.Bold,True
'@CTRL:Label,lblValKurtosis,fraStats,310,84,100,15
'@SET:lblValKurtosis,TextAlign,3
'@CTRL:Label,lblPctL0,fraStats,340,18,100,15,5th Percentile:
'@SET:lblPctL0,Font.Bold,True
'@CTRL:Label,lblPctV0,fraStats,450,18,100,15
'@SET:lblPctV0,TextAlign,3
'@CTRL:Label,lblPctL1,fraStats,340,40,100,15,10th Percentile:
'@SET:lblPctL1,Font.Bold,True
'@CTRL:Label,lblPctV1,fraStats,450,40,100,15
'@SET:lblPctV1,TextAlign,3
'@CTRL:Label,lblPctL2,fraStats,340,62,100,15,25th Percentile:
'@SET:lblPctL2,Font.Bold,True
'@CTRL:Label,lblPctV2,fraStats,450,62,100,15
'@SET:lblPctV2,TextAlign,3
'@CTRL:Label,lblPctL3,fraStats,340,84,100,15,50th Percentile:
'@SET:lblPctL3,Font.Bold,True
'@CTRL:Label,lblPctV3,fraStats,450,84,100,15
'@SET:lblPctV3,TextAlign,3
'@CTRL:Label,lblPctL4,fraStats,340,106,100,15,75th Percentile:
'@SET:lblPctL4,Font.Bold,True
'@CTRL:Label,lblPctV4,fraStats,450,106,100,15
'@SET:lblPctV4,TextAlign,3
'@CTRL:Label,lblPctL5,fraStats,340,128,100,15,90th Percentile:
'@SET:lblPctL5,Font.Bold,True
'@CTRL:Label,lblPctV5,fraStats,450,128,100,15
'@SET:lblPctV5,TextAlign,3
'@CTRL:Label,lblPctL6,fraStats,340,150,100,15,95th Percentile:
'@SET:lblPctL6,Font.Bold,True
'@CTRL:Label,lblPctV6,fraStats,450,150,100,15
'@SET:lblPctV6,TextAlign,3
'@CTRL:CommandButton,btnHistogram,,12,270,95,28,Histogram
'@CTRL:CommandButton,btnCumulative,,115,270,95,28,Cumulative
'@CTRL:CommandButton,btnOverlay,,218,270,95,28,Overlay
'@CTRL:CommandButton,btnTornado,,321,270,95,28,Tornado
'@CTRL:CommandButton,btnSpider,,424,270,95,28,Spider Plot
'@CTRL:CommandButton,btnAllCharts,,527,270,95,28,All Charts
'@CTRL:CommandButton,btnExportData,,12,310,120,28,Export Raw Data
'@CTRL:CommandButton,btnStatsReport,,140,310,120,28,Statistics Report
'@CTRL:Label,lblPreview,,12,350,200,15,Iteration Data (first 100):
'@SET:lblPreview,Font.Bold,True
'@CTRL:ListBox,lstIterData,,12,368,640,200
'@SET:lstIterData,ColumnCount,2
'@SET:lstIterData,ColumnWidths,60;100
'@CTRL:CommandButton,btnClose,,560,310,80,28,Close
'@SET:btnClose,Cancel,True
'@END
Option Explicit

Private m_CurrentOutput As Long

Private Sub UserForm_Initialize()
    ' Populate output combo
    Dim cboOutput As MSForms.ComboBox
    Set cboOutput = Me.Controls("cboOutput")
    Dim i As Long
    For i = 0 To g_OutputCount - 1
        cboOutput.AddItem g_Outputs(i).OutputName & " (" & g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress & ")"
    Next i
    If g_OutputCount > 0 Then cboOutput.ListIndex = 0

    ' Load first output
    If g_OutputCount > 0 Then LoadOutputStats 0
End Sub

Private Sub LoadOutputStats(ByVal idx As Long)
    If idx < 0 Or idx >= g_OutputCount Then Exit Sub
    m_CurrentOutput = idx

    Dim f As MSForms.Frame
    Set f = Me.Controls("fraStats")

    ' Update stat values
    f.Controls("lblValMean").Caption = Format(g_Outputs(idx).StatMean, "#,##0.0000")
    f.Controls("lblValStdDev").Caption = Format(g_Outputs(idx).StatStdDev, "#,##0.0000")
    f.Controls("lblValVariance").Caption = Format(g_Outputs(idx).StatVariance, "#,##0.0000")
    f.Controls("lblValMin").Caption = Format(g_Outputs(idx).StatMin, "#,##0.0000")
    f.Controls("lblValMax").Caption = Format(g_Outputs(idx).StatMax, "#,##0.0000")
    f.Controls("lblValMedian").Caption = Format(g_Outputs(idx).StatMedian, "#,##0.0000")
    f.Controls("lblValMode").Caption = Format(g_Outputs(idx).StatMode, "#,##0.0000")
    f.Controls("lblValSkewness").Caption = Format(g_Outputs(idx).StatSkewness, "#,##0.0000")
    f.Controls("lblValKurtosis").Caption = Format(g_Outputs(idx).StatKurtosis, "#,##0.0000")

    ' Percentiles
    Dim pcts As Variant
    pcts = Array(5, 10, 25, 50, 75, 90, 95)
    Dim i As Long
    For i = 0 To UBound(pcts)
        f.Controls("lblPctV" & i).Caption = Format(g_Outputs(idx).Percentiles(CLng(pcts(i))), "#,##0.0000")
    Next i

    ' Sim info
    Dim n As Long
    n = UBound(g_Outputs(idx).IterationData)
    Me.Controls("lblSimInfo").Caption = n & " iterations | " & _
        Format(g_SimEndTime - g_SimStartTime, "0.00") & "s"

    ' Load iteration data preview
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstIterData")
    lst.Clear
    Dim nShow As Long
    nShow = n
    If nShow > 100 Then nShow = 100
    For i = 1 To nShow
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = CStr(i)
        lst.List(lst.ListCount - 1, 1) = Format(g_Outputs(idx).IterationData(i), "#,##0.0000")
    Next i
End Sub

Private Sub cboOutput_Change()
    LoadOutputStats Me.Controls("cboOutput").ListIndex
End Sub

Private Sub btnHistogram_Click()
    CreateHistogram m_CurrentOutput
End Sub

Private Sub btnCumulative_Click()
    CreateCumulativeChart m_CurrentOutput
End Sub

Private Sub btnOverlay_Click()
    CreateOverlayChart m_CurrentOutput
End Sub

Private Sub btnTornado_Click()
    CreateTornadoDiagram m_CurrentOutput
End Sub

Private Sub btnSpider_Click()
    CreateSpiderPlot m_CurrentOutput
End Sub

Private Sub btnAllCharts_Click()
    GenerateAllCharts m_CurrentOutput
End Sub

Private Sub btnExportData_Click()
    OnExportData Nothing
End Sub

Private Sub btnStatsReport_Click()
    OnStatisticsReport Nothing
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub
