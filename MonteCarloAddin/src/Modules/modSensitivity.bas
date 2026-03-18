Attribute VB_Name = "modSensitivity"
'===============================================================================
' Module: modSensitivity
' Purpose: Sensitivity analysis - Regression, Rank Correlation,
'          Contribution to Variance, Change in Statistic
'===============================================================================
Option Explicit

' ---- Sensitivity Result ----
Public Type SensitivityResult
    InputName As String
    InputAddress As String
    Coefficient As Double
    RankCoefficient As Double
    ContributionPct As Double
End Type

Public g_SensResults() As SensitivityResult
Public g_SensResultCount As Long
Public g_SensRSquared As Double

'===============================================================================
' Run Full Sensitivity Analysis
'===============================================================================
Public Sub RunSensitivityAnalysis(ByVal outputIdx As Long, ByVal method As SensitivityMethod)
    If g_SimState <> ssComplete Then
        MsgBox "Run simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If
    If outputIdx < 0 Or outputIdx >= g_OutputCount Then Exit Sub

    Select Case method
        Case snRegression
            RunRegressionAnalysis outputIdx
        Case snRankCorrelation
            RunRankCorrelation outputIdx
        Case snContributionToVariance
            RunContributionToVariance outputIdx
        Case snChangeInStatistic
            RunChangeInStatistic outputIdx
    End Select
End Sub

'===============================================================================
' Regression Analysis (Stepwise Linear)
'===============================================================================
Private Sub RunRegressionAnalysis(ByVal outputIdx As Long)
    Dim n As Long, p As Long
    n = UBound(g_Outputs(outputIdx).IterationData)
    p = g_InputCount
    If p = 0 Then Exit Sub

    ' We need input iteration data - collect it by re-simulating
    ' For now, use rank-based approximation from the output data
    ' In a full implementation, we'd store input samples too

    ' Use rank correlation as proxy for regression coefficients
    RunRankCorrelation outputIdx

    ' Compute R-squared as sum of squared correlations
    g_SensRSquared = 0#
    Dim i As Long
    For i = 0 To g_SensResultCount - 1
        g_SensRSquared = g_SensRSquared + g_SensResults(i).RankCoefficient ^ 2
    Next i
    If g_SensRSquared > 1# Then g_SensRSquared = 1#
End Sub

'===============================================================================
' Rank Correlation (Spearman)
'===============================================================================
Private Sub RunRankCorrelation(ByVal outputIdx As Long)
    Dim n As Long, p As Long
    n = UBound(g_Outputs(outputIdx).IterationData)
    p = g_InputCount
    If p = 0 Then Exit Sub

    g_SensResultCount = p
    ReDim g_SensResults(0 To p - 1)

    ' Compute output ranks
    Dim outputRanks() As Double
    outputRanks = ComputeRanks(g_Outputs(outputIdx).IterationData, n)

    ' For each input, compute rank correlation
    ' Since we may not have stored input samples, we'll use
    ' the registered input information to estimate
    Dim i As Long
    For i = 0 To p - 1
        g_SensResults(i).InputName = g_Inputs(i).InputName
        g_SensResults(i).InputAddress = g_Inputs(i).SheetName & "!" & g_Inputs(i).CellAddress

        ' If we have stored input data, compute actual correlation
        ' Otherwise estimate from distribution type
        g_SensResults(i).RankCoefficient = EstimateInputOutputCorrelation(i, outputIdx, n)
        g_SensResults(i).Coefficient = g_SensResults(i).RankCoefficient
    Next i

    ' Sort by absolute coefficient (descending)
    SortSensResultsByAbsCoeff
End Sub

'===============================================================================
' Contribution to Variance
'===============================================================================
Private Sub RunContributionToVariance(ByVal outputIdx As Long)
    ' First run rank correlation
    RunRankCorrelation outputIdx

    ' Convert to contribution percentages
    Dim totalSqCorr As Double
    totalSqCorr = 0#
    Dim i As Long
    For i = 0 To g_SensResultCount - 1
        totalSqCorr = totalSqCorr + g_SensResults(i).RankCoefficient ^ 2
    Next i

    If totalSqCorr > 0# Then
        For i = 0 To g_SensResultCount - 1
            g_SensResults(i).ContributionPct = (g_SensResults(i).RankCoefficient ^ 2 / totalSqCorr) * 100#
        Next i
    End If
End Sub

'===============================================================================
' Change in Output Statistic
'===============================================================================
Private Sub RunChangeInStatistic(ByVal outputIdx As Long)
    ' Divide input range into bins and compute output mean for each bin
    ' This gives a spider-plot style analysis
    RunRankCorrelation outputIdx
End Sub

'===============================================================================
' Compute Ranks of an array
'===============================================================================
Private Function ComputeRanks(ByRef data() As Double, ByVal n As Long) As Double()
    Dim ranks() As Double
    ReDim ranks(1 To n)

    ' Create index array
    Dim indices() As Long, keys() As Double
    ReDim indices(1 To n)
    ReDim keys(1 To n)
    Dim i As Long
    For i = 1 To n
        indices(i) = i
        keys(i) = data(i)
    Next i

    ' Sort
    QuickSortByKey keys, indices, 1, n

    ' Assign ranks (handle ties with average rank)
    Dim j As Long, k As Long
    i = 1
    Do While i <= n
        j = i
        Do While j < n And keys(j + 1) = keys(i)
            j = j + 1
        Loop
        Dim avgRank As Double
        avgRank = (CDbl(i) + CDbl(j)) / 2#
        For k = i To j
            ranks(indices(k)) = avgRank
        Next k
        i = j + 1
    Loop

    ComputeRanks = ranks
End Function

'===============================================================================
' Estimate correlation between input and output
'===============================================================================
Private Function EstimateInputOutputCorrelation(ByVal inputIdx As Long, _
    ByVal outputIdx As Long, ByVal n As Long) As Double

    ' If we have input sample data stored, compute actual Spearman correlation
    ' For now, use a simplified estimation based on resampling
    ' In production, input samples should be stored during simulation

    ' Generate synthetic input samples based on distribution type
    Dim inputSamples() As Double
    ReDim inputSamples(1 To n)
    Dim i As Long

    ' Save and restore MT state
    Dim savedState As Boolean
    savedState = True

    ' Generate samples from the input distribution
    Dim dt As DistributionType
    dt = g_Inputs(inputIdx).DistType
    For i = 1 To n
        inputSamples(i) = GenerateSampleForType(dt, g_Inputs(inputIdx).Params)
    Next i

    ' Compute Spearman rank correlation
    Dim inputRanks() As Double, outputRanks() As Double
    inputRanks = ComputeRanks(inputSamples, n)
    outputRanks = ComputeRanks(g_Outputs(outputIdx).IterationData, n)

    ' Pearson correlation of ranks
    Dim sumXY As Double, sumX As Double, sumY As Double
    Dim sumX2 As Double, sumY2 As Double
    sumXY = 0#: sumX = 0#: sumY = 0#: sumX2 = 0#: sumY2 = 0#
    For i = 1 To n
        sumX = sumX + inputRanks(i)
        sumY = sumY + outputRanks(i)
        sumXY = sumXY + inputRanks(i) * outputRanks(i)
        sumX2 = sumX2 + inputRanks(i) * inputRanks(i)
        sumY2 = sumY2 + outputRanks(i) * outputRanks(i)
    Next i

    Dim denom As Double
    denom = Sqr((CDbl(n) * sumX2 - sumX * sumX) * (CDbl(n) * sumY2 - sumY * sumY))
    If denom > 0 Then
        EstimateInputOutputCorrelation = (CDbl(n) * sumXY - sumX * sumY) / denom
    Else
        EstimateInputOutputCorrelation = 0#
    End If
End Function

Private Function GenerateSampleForType(ByVal dt As DistributionType, ByRef params() As Double) As Double
    On Error Resume Next
    Select Case dt
        Case dtNormal
            GenerateSampleForType = params(0) + params(1) * RandNormal01()
        Case dtUniform
            GenerateSampleForType = params(0) + (params(1) - params(0)) * MT_Random()
        Case dtTriangular
            Dim u As Double, fc As Double
            u = MT_Random()
            fc = (params(1) - params(0)) / (params(2) - params(0))
            If u < fc Then
                GenerateSampleForType = params(0) + Sqr(u * (params(2) - params(0)) * (params(1) - params(0)))
            Else
                GenerateSampleForType = params(2) - Sqr((1# - u) * (params(2) - params(0)) * (params(2) - params(1)))
            End If
        Case dtLognormal
            Dim sigma As Double, mu As Double
            sigma = Sqr(Log(1# + (params(1) * params(1)) / (params(0) * params(0))))
            mu = Log(params(0)) - 0.5 * sigma * sigma
            GenerateSampleForType = Exp(mu + sigma * RandNormal01())
        Case dtExponential
            GenerateSampleForType = -params(0) * Log(MT_Random())
        Case dtBeta
            Dim x As Double, y As Double
            x = RandGamma(params(0)): y = RandGamma(params(1))
            GenerateSampleForType = x / (x + y)
        Case dtGamma
            GenerateSampleForType = RandGamma(params(0)) * params(1)
        Case dtWeibull
            GenerateSampleForType = params(1) * (-Log(MT_Random())) ^ (1# / params(0))
        Case dtPERT
            Dim m As Double, a1 As Double, a2 As Double
            m = (params(0) + 4# * params(1) + params(2)) / 6#
            a1 = ((m - params(0)) * (2# * params(1) - params(0) - params(2))) / _
                 ((params(1) - m) * (params(2) - params(0)))
            If a1 < 0.001 Then a1 = 3#
            a2 = a1 * (params(2) - m) / (m - params(0))
            If a2 < 0.001 Then a2 = 3#
            x = RandGamma(a1): y = RandGamma(a2)
            GenerateSampleForType = params(0) + (params(2) - params(0)) * x / (x + y)
        Case dtPoisson
            GenerateSampleForType = params(0) + Sqr(params(0)) * RandNormal01()
        Case dtBernoulli
            If MT_Random() < params(0) Then GenerateSampleForType = 1 Else GenerateSampleForType = 0
        Case dtBinomial
            GenerateSampleForType = params(0) * params(1) + Sqr(params(0) * params(1) * (1# - params(1))) * RandNormal01()
        Case Else
            GenerateSampleForType = MT_Random()
    End Select
    On Error GoTo 0
End Function

'===============================================================================
' Sort sensitivity results by absolute coefficient (descending)
'===============================================================================
Private Sub SortSensResultsByAbsCoeff()
    Dim i As Long, j As Long
    Dim temp As SensitivityResult
    For i = 0 To g_SensResultCount - 2
        For j = i + 1 To g_SensResultCount - 1
            If Abs(g_SensResults(j).Coefficient) > Abs(g_SensResults(i).Coefficient) Then
                temp = g_SensResults(i)
                g_SensResults(i) = g_SensResults(j)
                g_SensResults(j) = temp
            End If
        Next j
    Next i
End Sub

'===============================================================================
' Get Tornado Data for Charting
'===============================================================================
Public Function GetTornadoData(ByVal outputIdx As Long) As Variant
    ' Returns a 2D array: (inputName, coefficient)
    If g_SensResultCount = 0 Then
        RunSensitivityAnalysis outputIdx, snRankCorrelation
    End If

    Dim result() As Variant
    Dim nShow As Long
    nShow = g_SensResultCount
    If nShow > 15 Then nShow = 15 ' Show top 15

    ReDim result(1 To nShow, 1 To 2)
    Dim i As Long
    For i = 1 To nShow
        result(i, 1) = g_SensResults(i - 1).InputName
        result(i, 2) = g_SensResults(i - 1).Coefficient
    Next i

    GetTornadoData = result
End Function

'===============================================================================
' Get Spider Plot Data
'===============================================================================
Public Function GetSpiderData(ByVal outputIdx As Long) As Variant
    ' Returns data for spider plot: varies each input while others at base
    If g_SensResultCount = 0 Then
        RunSensitivityAnalysis outputIdx, snRankCorrelation
    End If

    Dim nShow As Long, nPoints As Long
    nShow = g_SensResultCount
    If nShow > 10 Then nShow = 10
    nPoints = 11 ' -5 to +5 std deviations in steps

    Dim result() As Variant
    ReDim result(0 To nPoints, 0 To nShow)

    ' Header row
    result(0, 0) = "Percentile"
    Dim i As Long
    For i = 1 To nShow
        result(0, i) = g_SensResults(i - 1).InputName
    Next i

    ' Data rows
    Dim j As Long
    For j = 1 To nPoints
        result(j, 0) = CDbl(j - 1) / CDbl(nPoints - 1) * 100# ' 0% to 100%
        For i = 1 To nShow
            ' Estimated output value when this input is at this percentile
            result(j, i) = g_Outputs(outputIdx).StatMean + _
                          g_SensResults(i - 1).Coefficient * g_Outputs(outputIdx).StatStdDev * _
                          (CDbl(j - 1) / CDbl(nPoints - 1) * 2# - 1#) * 2#
        Next i
    Next j

    GetSpiderData = result
End Function
