'===============================================================================
' Form: frmResults
' Purpose: Results browser - shows statistics and charts for selected output
'===============================================================================
Option Explicit

Private m_CurrentOutput As Long

Private Sub UserForm_Initialize()
    Me.Width = 680
    Me.Height = 620

    ' Output selector
    Dim lblOutput As MSForms.Label
    Set lblOutput = Me.Controls.Add("Forms.Label.1", "lblOutput")
    With lblOutput
        .Caption = "Output:"
        .Left = 12: .Top = 12: .Width = 50: .Height = 15
        .Font.Bold = True
    End With

    Dim cboOutput As MSForms.ComboBox
    Set cboOutput = Me.Controls.Add("Forms.ComboBox.1", "cboOutput")
    With cboOutput
        .Left = 65: .Top = 10: .Width = 250: .Height = 20
        .Style = fmStyleDropDownList
    End With

    ' Populate outputs
    Dim i As Long
    For i = 0 To g_OutputCount - 1
        cboOutput.AddItem g_Outputs(i).OutputName & " (" & g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress & ")"
    Next i
    If g_OutputCount > 0 Then cboOutput.ListIndex = 0

    ' Statistics frame
    Dim fraStats As MSForms.Frame
    Set fraStats = Me.Controls.Add("Forms.Frame.1", "fraStats")
    With fraStats
        .Caption = "Statistics"
        .Left = 12: .Top = 40: .Width = 640: .Height = 220
    End With

    ' Stats labels (left column)
    Dim statLabels As Variant
    statLabels = Array("Mean:", "Std Deviation:", "Variance:", "Minimum:", "Maximum:", _
                       "Median:", "Mode:", "Skewness:", "Kurtosis:")
    Dim statNames As Variant
    statNames = Array("lblStatMean", "lblStatStdDev", "lblStatVariance", "lblStatMin", _
                     "lblStatMax", "lblStatMedian", "lblStatMode", "lblStatSkewness", "lblStatKurtosis")
    Dim valNames As Variant
    valNames = Array("lblValMean", "lblValStdDev", "lblValVariance", "lblValMin", _
                    "lblValMax", "lblValMedian", "lblValMode", "lblValSkewness", "lblValKurtosis")

    Dim y As Long
    y = 18
    For i = 0 To UBound(statLabels)
        Dim lbl As MSForms.Label
        Set lbl = fraStats.Controls.Add("Forms.Label.1", CStr(statNames(i)))
        With lbl
            .Caption = CStr(statLabels(i))
            .Left = 12: .Top = y: .Width = 80: .Height = 15
            .Font.Bold = True
        End With

        Dim val As MSForms.Label
        Set val = fraStats.Controls.Add("Forms.Label.1", CStr(valNames(i)))
        With val
            .Caption = ""
            .Left = 100: .Top = y: .Width = 100: .Height = 15
            .TextAlign = fmTextAlignRight
        End With

        If i = 4 Then
            y = 18 ' Reset for right column
        Else
            y = y + 22
        End If
    Next i

    ' Percentile labels (right column)
    Dim pctLabels As Variant
    pctLabels = Array("5th Percentile:", "10th Percentile:", "25th Percentile:", _
                     "50th Percentile:", "75th Percentile:", "90th Percentile:", "95th Percentile:")
    Dim pctValues As Variant
    pctValues = Array(5, 10, 25, 50, 75, 90, 95)

    y = 18
    For i = 0 To UBound(pctLabels)
        Set lbl = fraStats.Controls.Add("Forms.Label.1", "lblPctL" & i)
        With lbl
            .Caption = CStr(pctLabels(i))
            .Left = 340: .Top = y: .Width = 100: .Height = 15
            .Font.Bold = True
        End With

        Set val = fraStats.Controls.Add("Forms.Label.1", "lblPctV" & i)
        With val
            .Caption = ""
            .Left = 450: .Top = y: .Width = 100: .Height = 15
            .TextAlign = fmTextAlignRight
        End With
        y = y + 22
    Next i

    ' Simulation info
    Dim lblInfo As MSForms.Label
    Set lblInfo = Me.Controls.Add("Forms.Label.1", "lblSimInfo")
    With lblInfo
        .Caption = ""
        .Left = 330: .Top = 12: .Width = 320: .Height = 15
    End With

    ' Chart buttons
    Dim btnHistogram As MSForms.CommandButton
    Set btnHistogram = Me.Controls.Add("Forms.CommandButton.1", "btnHistogram")
    With btnHistogram
        .Caption = "Histogram"
        .Left = 12: .Top = 270: .Width = 95: .Height = 28
    End With

    Dim btnCumulative As MSForms.CommandButton
    Set btnCumulative = Me.Controls.Add("Forms.CommandButton.1", "btnCumulative")
    With btnCumulative
        .Caption = "Cumulative"
        .Left = 115: .Top = 270: .Width = 95: .Height = 28
    End With

    Dim btnOverlay As MSForms.CommandButton
    Set btnOverlay = Me.Controls.Add("Forms.CommandButton.1", "btnOverlay")
    With btnOverlay
        .Caption = "Overlay"
        .Left = 218: .Top = 270: .Width = 95: .Height = 28
    End With

    Dim btnTornado As MSForms.CommandButton
    Set btnTornado = Me.Controls.Add("Forms.CommandButton.1", "btnTornado")
    With btnTornado
        .Caption = "Tornado"
        .Left = 321: .Top = 270: .Width = 95: .Height = 28
    End With

    Dim btnSpider As MSForms.CommandButton
    Set btnSpider = Me.Controls.Add("Forms.CommandButton.1", "btnSpider")
    With btnSpider
        .Caption = "Spider Plot"
        .Left = 424: .Top = 270: .Width = 95: .Height = 28
    End With

    Dim btnAllCharts As MSForms.CommandButton
    Set btnAllCharts = Me.Controls.Add("Forms.CommandButton.1", "btnAllCharts")
    With btnAllCharts
        .Caption = "All Charts"
        .Left = 527: .Top = 270: .Width = 95: .Height = 28
    End With

    ' Export buttons
    Dim btnExport As MSForms.CommandButton
    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExportData")
    With btnExport
        .Caption = "Export Raw Data"
        .Left = 12: .Top = 310: .Width = 120: .Height = 28
    End With

    Dim btnStats As MSForms.CommandButton
    Set btnStats = Me.Controls.Add("Forms.CommandButton.1", "btnStatsReport")
    With btnStats
        .Caption = "Statistics Report"
        .Left = 140: .Top = 310: .Width = 120: .Height = 28
    End With

    ' Iteration data preview (ListBox)
    Dim lblPreview As MSForms.Label
    Set lblPreview = Me.Controls.Add("Forms.Label.1", "lblPreview")
    With lblPreview
        .Caption = "Iteration Data (first 100):"
        .Left = 12: .Top = 350: .Width = 200: .Height = 15
        .Font.Bold = True
    End With

    Dim lstData As MSForms.ListBox
    Set lstData = Me.Controls.Add("Forms.ListBox.1", "lstIterData")
    With lstData
        .Left = 12: .Top = 368: .Width = 640: .Height = 200
        .ColumnCount = 2
        .ColumnWidths = "60;100"
    End With

    ' Close button
    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = 560: .Top = 310: .Width = 80: .Height = 28
        .Cancel = True
    End With

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
