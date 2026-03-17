Attribute VB_Name = "modCharting"
'===============================================================================
' Module: modCharting
' Purpose: Chart generation - Histograms, Cumulative, Tornado, Scatter, Spider
'===============================================================================
Option Explicit

Private Const CHART_WIDTH As Long = 500
Private Const CHART_HEIGHT As Long = 350

'===============================================================================
' Create Histogram Chart
'===============================================================================
Public Sub CreateHistogram(ByVal outputIdx As Long, Optional ByVal targetSheet As Worksheet = Nothing, _
    Optional ByVal leftPos As Double = 10, Optional ByVal topPos As Double = 10)

    If g_SimState <> ssComplete Then Exit Sub
    If outputIdx < 0 Or outputIdx >= g_OutputCount Then Exit Sub

    Dim n As Long
    n = UBound(g_Outputs(outputIdx).IterationData)

    ' Compute histogram bins
    Dim nBins As Long
    nBins = CLng(Sqr(CDbl(n)))
    If nBins < 20 Then nBins = 20
    If nBins > 100 Then nBins = 100

    Dim minV As Double, maxV As Double, binWidth As Double
    minV = g_Outputs(outputIdx).StatMin
    maxV = g_Outputs(outputIdx).StatMax
    If maxV = minV Then maxV = minV + 1#
    binWidth = (maxV - minV) / CDbl(nBins)

    Dim bins() As Long, binMids() As Double
    ReDim bins(1 To nBins)
    ReDim binMids(1 To nBins)

    Dim i As Long
    For i = 1 To nBins
        binMids(i) = minV + (CDbl(i) - 0.5) * binWidth
    Next i

    For i = 1 To n
        Dim bIdx As Long
        bIdx = CLng(Int((g_Outputs(outputIdx).IterationData(i) - minV) / binWidth)) + 1
        If bIdx < 1 Then bIdx = 1
        If bIdx > nBins Then bIdx = nBins
        bins(bIdx) = bins(bIdx) + 1
    Next i

    ' Convert to relative frequency
    Dim relFreq() As Double
    ReDim relFreq(1 To nBins)
    For i = 1 To nBins
        relFreq(i) = CDbl(bins(i)) / CDbl(n)
    Next i

    ' Create chart data on a temp sheet or results sheet
    If targetSheet Is Nothing Then
        Set targetSheet = GetOrCreateResultsSheet()
    End If

    ' Write data to hidden area
    Dim dataStart As Long
    dataStart = FindNextEmptyRow(targetSheet)
    targetSheet.Cells(dataStart, 20).Value = "Bin"
    targetSheet.Cells(dataStart, 21).Value = "Frequency"
    For i = 1 To nBins
        targetSheet.Cells(dataStart + i, 20).Value = binMids(i)
        targetSheet.Cells(dataStart + i, 21).Value = relFreq(i)
    Next i

    ' Create chart
    Dim cht As ChartObject
    Set cht = targetSheet.ChartObjects.Add(leftPos, topPos, CHART_WIDTH, CHART_HEIGHT)
    With cht.Chart
        .ChartType = xlColumnClustered
        .SetSourceData targetSheet.Range(targetSheet.Cells(dataStart, 20), _
                       targetSheet.Cells(dataStart + nBins, 21))
        .HasTitle = True
        .ChartTitle.Text = "Distribution: " & g_Outputs(outputIdx).OutputName
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Value"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Relative Frequency"
        .HasLegend = False

        ' Color bars
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(70, 130, 180)
        .SeriesCollection(1).Format.Fill.Transparency = 0.2

        ' Add mean line annotation
        Dim meanStr As String
        meanStr = "Mean = " & Format(g_Outputs(outputIdx).StatMean, "#,##0.00") & vbCrLf & _
                 "Std Dev = " & Format(g_Outputs(outputIdx).StatStdDev, "#,##0.00") & vbCrLf & _
                 "N = " & Format(n, "#,##0")

        ' Add text box
        Dim tb As Shape
        Set tb = targetSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                leftPos + CHART_WIDTH - 160, topPos + 30, 150, 60)
        tb.TextFrame.Characters.Text = meanStr
        tb.TextFrame.Characters.Font.Size = 8
        tb.Fill.Visible = msoFalse
        tb.Line.Visible = msoFalse
    End With

    ' Hide data columns
    targetSheet.Columns(20).Hidden = True
    targetSheet.Columns(21).Hidden = True
End Sub

'===============================================================================
' Create Cumulative Distribution Chart
'===============================================================================
Public Sub CreateCumulativeChart(ByVal outputIdx As Long, Optional ByVal ascending As Boolean = True, _
    Optional ByVal targetSheet As Worksheet = Nothing, _
    Optional ByVal leftPos As Double = 10, Optional ByVal topPos As Double = 370)

    If g_SimState <> ssComplete Then Exit Sub
    If outputIdx < 0 Or outputIdx >= g_OutputCount Then Exit Sub

    If targetSheet Is Nothing Then
        Set targetSheet = GetOrCreateResultsSheet()
    End If

    ' Use percentile data
    Dim nPoints As Long
    nPoints = 101

    Dim dataStart As Long
    dataStart = FindNextEmptyRow(targetSheet)
    targetSheet.Cells(dataStart, 22).Value = "Value"
    targetSheet.Cells(dataStart, 23).Value = "Cumulative Probability"

    Dim i As Long
    For i = 0 To 100
        targetSheet.Cells(dataStart + i + 1, 22).Value = g_Outputs(outputIdx).Percentiles(i)
        If ascending Then
            targetSheet.Cells(dataStart + i + 1, 23).Value = CDbl(i) / 100#
        Else
            targetSheet.Cells(dataStart + i + 1, 23).Value = 1# - CDbl(i) / 100#
        End If
    Next i

    Dim cht As ChartObject
    Set cht = targetSheet.ChartObjects.Add(leftPos, topPos, CHART_WIDTH, CHART_HEIGHT)
    With cht.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        .SetSourceData targetSheet.Range(targetSheet.Cells(dataStart, 22), _
                       targetSheet.Cells(dataStart + 101, 23))
        .HasTitle = True
        If ascending Then
            .ChartTitle.Text = "Cumulative Ascending: " & g_Outputs(outputIdx).OutputName
        Else
            .ChartTitle.Text = "Cumulative Descending: " & g_Outputs(outputIdx).OutputName
        End If
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Value"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Probability"
        .HasLegend = False
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(220, 50, 50)
        .SeriesCollection(1).Format.Line.Weight = 2
    End With

    targetSheet.Columns(22).Hidden = True
    targetSheet.Columns(23).Hidden = True
End Sub

'===============================================================================
' Create Tornado Diagram
'===============================================================================
Public Sub CreateTornadoDiagram(ByVal outputIdx As Long, _
    Optional ByVal targetSheet As Worksheet = Nothing, _
    Optional ByVal leftPos As Double = 520, Optional ByVal topPos As Double = 10)

    If g_SimState <> ssComplete Then Exit Sub

    ' Run sensitivity analysis if not done
    If g_SensResultCount = 0 Then
        RunSensitivityAnalysis outputIdx, snRankCorrelation
    End If

    If g_SensResultCount = 0 Then Exit Sub

    If targetSheet Is Nothing Then
        Set targetSheet = GetOrCreateResultsSheet()
    End If

    Dim nShow As Long
    nShow = g_SensResultCount
    If nShow > 15 Then nShow = 15

    Dim dataStart As Long
    dataStart = FindNextEmptyRow(targetSheet)
    targetSheet.Cells(dataStart, 24).Value = "Input"
    targetSheet.Cells(dataStart, 25).Value = "Coefficient"

    Dim i As Long
    For i = 1 To nShow
        targetSheet.Cells(dataStart + i, 24).Value = g_SensResults(i - 1).InputName
        targetSheet.Cells(dataStart + i, 25).Value = g_SensResults(i - 1).Coefficient
    Next i

    Dim cht As ChartObject
    Set cht = targetSheet.ChartObjects.Add(leftPos, topPos, CHART_WIDTH, CHART_HEIGHT)
    With cht.Chart
        .ChartType = xlBarClustered
        .SetSourceData targetSheet.Range(targetSheet.Cells(dataStart, 24), _
                       targetSheet.Cells(dataStart + nShow, 25))
        .HasTitle = True
        .ChartTitle.Text = "Tornado: " & g_Outputs(outputIdx).OutputName
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Std Regression Coefficient"
        .HasLegend = False

        ' Color positive/negative differently
        Dim pt As Long
        For pt = 1 To .SeriesCollection(1).Points.count
            If g_SensResults(pt - 1).Coefficient >= 0 Then
                .SeriesCollection(1).Points(pt).Format.Fill.ForeColor.RGB = RGB(70, 130, 180)
            Else
                .SeriesCollection(1).Points(pt).Format.Fill.ForeColor.RGB = RGB(220, 80, 80)
            End If
        Next pt
    End With

    targetSheet.Columns(24).Hidden = True
    targetSheet.Columns(25).Hidden = True
End Sub

'===============================================================================
' Create Scatter Plot
'===============================================================================
Public Sub CreateScatterPlot(ByVal outputIdx1 As Long, ByVal outputIdx2 As Long, _
    Optional ByVal targetSheet As Worksheet = Nothing, _
    Optional ByVal leftPos As Double = 520, Optional ByVal topPos As Double = 370)

    If g_SimState <> ssComplete Then Exit Sub
    If outputIdx1 < 0 Or outputIdx1 >= g_OutputCount Then Exit Sub
    If outputIdx2 < 0 Or outputIdx2 >= g_OutputCount Then Exit Sub

    If targetSheet Is Nothing Then
        Set targetSheet = GetOrCreateResultsSheet()
    End If

    Dim n As Long
    n = UBound(g_Outputs(outputIdx1).IterationData)
    Dim nPlot As Long
    nPlot = n
    If nPlot > 2000 Then nPlot = 2000 ' Limit points for performance

    Dim dataStart As Long
    dataStart = FindNextEmptyRow(targetSheet)
    targetSheet.Cells(dataStart, 26).Value = g_Outputs(outputIdx1).OutputName
    targetSheet.Cells(dataStart, 27).Value = g_Outputs(outputIdx2).OutputName

    Dim i As Long, step As Long
    step = n \ nPlot
    If step < 1 Then step = 1
    Dim row As Long
    row = 1
    For i = 1 To n Step step
        targetSheet.Cells(dataStart + row, 26).Value = g_Outputs(outputIdx1).IterationData(i)
        targetSheet.Cells(dataStart + row, 27).Value = g_Outputs(outputIdx2).IterationData(i)
        row = row + 1
    Next i

    Dim cht As ChartObject
    Set cht = targetSheet.ChartObjects.Add(leftPos, topPos, CHART_WIDTH, CHART_HEIGHT)
    With cht.Chart
        .ChartType = xlXYScatter
        .SetSourceData targetSheet.Range(targetSheet.Cells(dataStart, 26), _
                       targetSheet.Cells(dataStart + row - 1, 27))
        .HasTitle = True
        .ChartTitle.Text = g_Outputs(outputIdx1).OutputName & " vs " & g_Outputs(outputIdx2).OutputName
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = g_Outputs(outputIdx1).OutputName
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = g_Outputs(outputIdx2).OutputName
        .HasLegend = False
        .SeriesCollection(1).MarkerSize = 3
        .SeriesCollection(1).MarkerForegroundColor = RGB(70, 130, 180)
        .SeriesCollection(1).MarkerBackgroundColor = RGB(70, 130, 180)
    End With

    targetSheet.Columns(26).Hidden = True
    targetSheet.Columns(27).Hidden = True
End Sub

'===============================================================================
' Create Spider Plot
'===============================================================================
Public Sub CreateSpiderPlot(ByVal outputIdx As Long, _
    Optional ByVal targetSheet As Worksheet = Nothing, _
    Optional ByVal leftPos As Double = 10, Optional ByVal topPos As Double = 740)

    If g_SimState <> ssComplete Then Exit Sub
    If g_SensResultCount = 0 Then
        RunSensitivityAnalysis outputIdx, snRankCorrelation
    End If

    If targetSheet Is Nothing Then
        Set targetSheet = GetOrCreateResultsSheet()
    End If

    Dim spiderData As Variant
    spiderData = GetSpiderData(outputIdx)

    Dim nRows As Long, nCols As Long
    nRows = UBound(spiderData, 1) + 1
    nCols = UBound(spiderData, 2) + 1

    Dim dataStart As Long
    dataStart = FindNextEmptyRow(targetSheet)

    Dim i As Long, j As Long
    For i = 0 To UBound(spiderData, 1)
        For j = 0 To UBound(spiderData, 2)
            targetSheet.Cells(dataStart + i, 28 + j).Value = spiderData(i, j)
        Next j
    Next i

    Dim cht As ChartObject
    Set cht = targetSheet.ChartObjects.Add(leftPos, topPos, CHART_WIDTH, CHART_HEIGHT)
    With cht.Chart
        .ChartType = xlXYScatterSmoothNoMarkers
        .SetSourceData targetSheet.Range(targetSheet.Cells(dataStart, 28), _
                       targetSheet.Cells(dataStart + UBound(spiderData, 1), 28 + UBound(spiderData, 2)))
        .HasTitle = True
        .ChartTitle.Text = "Spider Plot: " & g_Outputs(outputIdx).OutputName
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Input Percentile"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Output Value"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub

'===============================================================================
' Create Overlay Chart (Histogram + Cumulative)
'===============================================================================
Public Sub CreateOverlayChart(ByVal outputIdx As Long, _
    Optional ByVal targetSheet As Worksheet = Nothing, _
    Optional ByVal leftPos As Double = 10, Optional ByVal topPos As Double = 10)

    If g_SimState <> ssComplete Then Exit Sub

    If targetSheet Is Nothing Then
        Set targetSheet = GetOrCreateResultsSheet()
    End If

    Dim n As Long
    n = UBound(g_Outputs(outputIdx).IterationData)

    ' Compute histogram
    Dim nBins As Long
    nBins = CLng(Sqr(CDbl(n)))
    If nBins < 20 Then nBins = 20
    If nBins > 80 Then nBins = 80

    Dim minV As Double, maxV As Double, binWidth As Double
    minV = g_Outputs(outputIdx).StatMin
    maxV = g_Outputs(outputIdx).StatMax
    If maxV = minV Then maxV = minV + 1#
    binWidth = (maxV - minV) / CDbl(nBins)

    Dim bins() As Long, binMids() As Double
    ReDim bins(1 To nBins)
    ReDim binMids(1 To nBins)

    Dim i As Long
    For i = 1 To nBins
        binMids(i) = minV + (CDbl(i) - 0.5) * binWidth
    Next i
    For i = 1 To n
        Dim bIdx As Long
        bIdx = CLng(Int((g_Outputs(outputIdx).IterationData(i) - minV) / binWidth)) + 1
        If bIdx < 1 Then bIdx = 1
        If bIdx > nBins Then bIdx = nBins
        bins(bIdx) = bins(bIdx) + 1
    Next i

    Dim dataStart As Long
    dataStart = FindNextEmptyRow(targetSheet)
    targetSheet.Cells(dataStart, 30).Value = "Bin"
    targetSheet.Cells(dataStart, 31).Value = "Frequency"
    targetSheet.Cells(dataStart, 32).Value = "Cumulative"

    Dim cumSum As Double
    cumSum = 0#
    For i = 1 To nBins
        targetSheet.Cells(dataStart + i, 30).Value = binMids(i)
        targetSheet.Cells(dataStart + i, 31).Value = CDbl(bins(i)) / CDbl(n)
        cumSum = cumSum + CDbl(bins(i)) / CDbl(n)
        targetSheet.Cells(dataStart + i, 32).Value = cumSum
    Next i

    Dim cht As ChartObject
    Set cht = targetSheet.ChartObjects.Add(leftPos, topPos, CHART_WIDTH + 50, CHART_HEIGHT + 30)
    With cht.Chart
        .ChartType = xlColumnClustered

        ' Add histogram series
        .SetSourceData targetSheet.Range(targetSheet.Cells(dataStart, 30), _
                       targetSheet.Cells(dataStart + nBins, 31))

        ' Add cumulative line on secondary axis
        Dim s2 As Series
        Set s2 = .SeriesCollection.NewSeries
        s2.Values = targetSheet.Range(targetSheet.Cells(dataStart + 1, 32), _
                                     targetSheet.Cells(dataStart + nBins, 32))
        s2.XValues = targetSheet.Range(targetSheet.Cells(dataStart + 1, 30), _
                                      targetSheet.Cells(dataStart + nBins, 30))
        s2.AxisGroup = xlSecondary
        s2.ChartType = xlXYScatterSmoothNoMarkers
        s2.Name = "Cumulative"
        s2.Format.Line.ForeColor.RGB = RGB(220, 50, 50)
        s2.Format.Line.Weight = 2

        .SeriesCollection(1).Name = "Frequency"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(70, 130, 180)

        .HasTitle = True
        .ChartTitle.Text = "Distribution: " & g_Outputs(outputIdx).OutputName
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Relative Frequency"
        .Axes(xlValue, xlSecondary).HasTitle = True
        .Axes(xlValue, xlSecondary).AxisTitle.Text = "Cumulative Probability"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub

'===============================================================================
' Create Summary Statistics Table
'===============================================================================
Public Sub CreateStatisticsTable(Optional ByVal targetSheet As Worksheet = Nothing, _
    Optional ByVal startRow As Long = 1)

    If g_SimState <> ssComplete Then Exit Sub

    If targetSheet Is Nothing Then
        Set targetSheet = GetOrCreateResultsSheet()
    End If

    Dim row As Long
    row = startRow

    ' Header
    With targetSheet
        .Cells(row, 1).Value = "Output Name"
        .Cells(row, 2).Value = "Cell"
        .Cells(row, 3).Value = "Mean"
        .Cells(row, 4).Value = "Std Dev"
        .Cells(row, 5).Value = "Variance"
        .Cells(row, 6).Value = "Minimum"
        .Cells(row, 7).Value = "Maximum"
        .Cells(row, 8).Value = "Median"
        .Cells(row, 9).Value = "Mode"
        .Cells(row, 10).Value = "Skewness"
        .Cells(row, 11).Value = "Kurtosis"
        .Cells(row, 12).Value = "5th Pctl"
        .Cells(row, 13).Value = "10th Pctl"
        .Cells(row, 14).Value = "25th Pctl"
        .Cells(row, 15).Value = "75th Pctl"
        .Cells(row, 16).Value = "90th Pctl"
        .Cells(row, 17).Value = "95th Pctl"

        ' Bold header
        .Range(.Cells(row, 1), .Cells(row, 17)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row, 17)).Interior.Color = RGB(70, 130, 180)
        .Range(.Cells(row, 1), .Cells(row, 17)).Font.Color = RGB(255, 255, 255)

        Dim i As Long
        For i = 0 To g_OutputCount - 1
            row = row + 1
            .Cells(row, 1).Value = g_Outputs(i).OutputName
            .Cells(row, 2).Value = g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress
            .Cells(row, 3).Value = g_Outputs(i).StatMean
            .Cells(row, 4).Value = g_Outputs(i).StatStdDev
            .Cells(row, 5).Value = g_Outputs(i).StatVariance
            .Cells(row, 6).Value = g_Outputs(i).StatMin
            .Cells(row, 7).Value = g_Outputs(i).StatMax
            .Cells(row, 8).Value = g_Outputs(i).StatMedian
            .Cells(row, 9).Value = g_Outputs(i).StatMode
            .Cells(row, 10).Value = g_Outputs(i).StatSkewness
            .Cells(row, 11).Value = g_Outputs(i).StatKurtosis
            .Cells(row, 12).Value = g_Outputs(i).Percentiles(5)
            .Cells(row, 13).Value = g_Outputs(i).Percentiles(10)
            .Cells(row, 14).Value = g_Outputs(i).Percentiles(25)
            .Cells(row, 15).Value = g_Outputs(i).Percentiles(75)
            .Cells(row, 16).Value = g_Outputs(i).Percentiles(90)
            .Cells(row, 17).Value = g_Outputs(i).Percentiles(95)

            ' Number format
            .Range(.Cells(row, 3), .Cells(row, 17)).NumberFormat = "#,##0.00"
        Next i

        .Columns("A:Q").AutoFit
    End With
End Sub

'===============================================================================
' Helper: Get or create results sheet
'===============================================================================
Public Function GetOrCreateResultsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("MC Results")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
        ws.Name = "MC Results"
    End If
    Set GetOrCreateResultsSheet = ws
End Function

'===============================================================================
' Helper: Find next empty row
'===============================================================================
Private Function FindNextEmptyRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 20).End(xlUp).row
    If lastRow < 1 Then lastRow = 1
    FindNextEmptyRow = lastRow + 2
End Function

'===============================================================================
' Generate All Charts for an Output
'===============================================================================
Public Sub GenerateAllCharts(ByVal outputIdx As Long)
    Dim ws As Worksheet
    Set ws = GetOrCreateResultsSheet()
    ws.Cells.Clear

    ' Statistics table at top
    CreateStatisticsTable ws, 1

    ' Charts below
    Dim chartTop As Double
    chartTop = (g_OutputCount + 3) * 15 + 10

    CreateHistogram outputIdx, ws, 10, chartTop
    CreateCumulativeChart outputIdx, True, ws, 10, chartTop + CHART_HEIGHT + 20
    CreateTornadoDiagram outputIdx, ws, CHART_WIDTH + 30, chartTop
    If g_OutputCount > 1 Then
        CreateScatterPlot 0, IIf(outputIdx = 0, 1, 0), ws, CHART_WIDTH + 30, chartTop + CHART_HEIGHT + 20
    End If
    CreateSpiderPlot outputIdx, ws, 10, chartTop + (CHART_HEIGHT + 20) * 2
    CreateOverlayChart outputIdx, ws, CHART_WIDTH + 30, chartTop + (CHART_HEIGHT + 20) * 2

    ws.Activate
End Sub
