Attribute VB_Name = "modDistFitting"
'===============================================================================
' Module: modDistFitting
' Purpose: Distribution fitting - fit data to best distribution
'          Mirrors @RISK's distribution fitting feature
'===============================================================================
Option Explicit

Public Type FitResult
    DistName As String
    DistType As DistributionType
    Params(0 To 3) As Double
    ParamCount As Long
    ChiSquareStat As Double
    KSStat As Double
    AIC As Double
    BIC As Double
    LogLikelihood As Double
    Rank As Long
End Type

'===============================================================================
' Fit data to multiple distributions and rank by goodness of fit
'===============================================================================
Public Function FitDistributions(ByRef data() As Double) As FitResult()
    Dim n As Long
    n = UBound(data) - LBound(data) + 1
    If n < 10 Then
        MsgBox "Need at least 10 data points for distribution fitting.", vbExclamation
        Exit Function
    End If

    ' Sort data
    Dim sorted() As Double
    ReDim sorted(1 To n)
    Dim i As Long
    For i = 1 To n
        sorted(i) = data(LBound(data) + i - 1)
    Next i
    QuickSortDouble sorted, 1, n

    ' Compute sample statistics
    Dim mean As Double, stdDev As Double, minV As Double, maxV As Double
    Dim sum As Double, sumSq As Double
    sum = 0#: sumSq = 0#
    minV = sorted(1): maxV = sorted(n)
    For i = 1 To n
        sum = sum + sorted(i)
    Next i
    mean = sum / CDbl(n)
    For i = 1 To n
        sumSq = sumSq + (sorted(i) - mean) ^ 2
    Next i
    stdDev = Sqr(sumSq / (n - 1))

    ' Try fitting each distribution
    Dim results() As FitResult
    ReDim results(0 To 6)
    Dim idx As Long
    idx = 0

    ' 1. Normal
    results(idx).DistName = "Normal"
    results(idx).DistType = dtNormal
    results(idx).Params(0) = mean
    results(idx).Params(1) = stdDev
    results(idx).ParamCount = 2
    results(idx).KSStat = KSTestNormal(sorted, n, mean, stdDev)
    idx = idx + 1

    ' 2. Lognormal (if all positive)
    If minV > 0 Then
        Dim logMean As Double, logSD As Double
        Dim logSum As Double, logSumSq As Double
        logSum = 0#: logSumSq = 0#
        For i = 1 To n
            logSum = logSum + Log(sorted(i))
        Next i
        logMean = logSum / CDbl(n)
        For i = 1 To n
            logSumSq = logSumSq + (Log(sorted(i)) - logMean) ^ 2
        Next i
        logSD = Sqr(logSumSq / (n - 1))

        results(idx).DistName = "Lognormal"
        results(idx).DistType = dtLognormal
        results(idx).Params(0) = Exp(logMean + logSD * logSD / 2#) ' arithmetic mean
        results(idx).Params(1) = results(idx).Params(0) * Sqr(Exp(logSD * logSD) - 1#) ' arithmetic sd
        results(idx).ParamCount = 2
        results(idx).KSStat = KSTestLognormal(sorted, n, logMean, logSD)
        idx = idx + 1
    End If

    ' 3. Uniform
    results(idx).DistName = "Uniform"
    results(idx).DistType = dtUniform
    results(idx).Params(0) = minV
    results(idx).Params(1) = maxV
    results(idx).ParamCount = 2
    results(idx).KSStat = KSTestUniform(sorted, n, minV, maxV)
    idx = idx + 1

    ' 4. Exponential (if all positive)
    If minV >= 0 Then
        results(idx).DistName = "Exponential"
        results(idx).DistType = dtExponential
        results(idx).Params(0) = mean
        results(idx).ParamCount = 1
        results(idx).KSStat = KSTestExponential(sorted, n, mean)
        idx = idx + 1
    End If

    ' 5. Triangular
    results(idx).DistName = "Triangular"
    results(idx).DistType = dtTriangular
    results(idx).Params(0) = minV
    results(idx).Params(1) = 3# * mean - minV - maxV ' estimate mode
    results(idx).Params(2) = maxV
    results(idx).ParamCount = 3
    results(idx).KSStat = 1# ' placeholder
    idx = idx + 1

    ' 6. PERT
    results(idx).DistName = "PERT"
    results(idx).DistType = dtPERT
    results(idx).Params(0) = minV
    results(idx).Params(1) = 3# * mean - minV - maxV ' estimate mode
    results(idx).Params(2) = maxV
    results(idx).ParamCount = 3
    results(idx).KSStat = 1# ' placeholder
    idx = idx + 1

    ' 7. Weibull (if all positive)
    If minV > 0 Then
        results(idx).DistName = "Weibull"
        results(idx).DistType = dtWeibull
        ' Method of moments approximation
        Dim cv As Double
        cv = stdDev / mean
        results(idx).Params(0) = 1# / cv ' rough shape estimate
        results(idx).Params(1) = mean / GammaFunc(1# + cv) ' rough scale
        results(idx).ParamCount = 2
        results(idx).KSStat = 1#
        idx = idx + 1
    End If

    ReDim Preserve results(0 To idx - 1)

    ' Sort by KS statistic (lower is better)
    Dim j As Long
    Dim tempResult As FitResult
    For i = 0 To idx - 2
        For j = i + 1 To idx - 1
            If results(j).KSStat < results(i).KSStat Then
                tempResult = results(i)
                results(i) = results(j)
                results(j) = tempResult
            End If
        Next j
    Next i

    ' Assign ranks
    For i = 0 To idx - 1
        results(i).Rank = i + 1
    Next i

    FitDistributions = results
End Function

'===============================================================================
' Kolmogorov-Smirnov Tests
'===============================================================================
Private Function KSTestNormal(ByRef sorted() As Double, ByVal n As Long, _
    ByVal mean As Double, ByVal sd As Double) As Double
    Dim maxD As Double, d As Double
    maxD = 0#
    Dim i As Long
    For i = 1 To n
        Dim empiricalCDF As Double
        empiricalCDF = CDbl(i) / CDbl(n)
        Dim theorCDF As Double
        theorCDF = NormalCDF((sorted(i) - mean) / sd)
        d = Abs(empiricalCDF - theorCDF)
        If d > maxD Then maxD = d
        d = Abs(CDbl(i - 1) / CDbl(n) - theorCDF)
        If d > maxD Then maxD = d
    Next i
    KSTestNormal = maxD
End Function

Private Function KSTestLognormal(ByRef sorted() As Double, ByVal n As Long, _
    ByVal logMean As Double, ByVal logSD As Double) As Double
    Dim maxD As Double, d As Double
    maxD = 0#
    Dim i As Long
    For i = 1 To n
        Dim empiricalCDF As Double
        empiricalCDF = CDbl(i) / CDbl(n)
        Dim theorCDF As Double
        theorCDF = NormalCDF((Log(sorted(i)) - logMean) / logSD)
        d = Abs(empiricalCDF - theorCDF)
        If d > maxD Then maxD = d
    Next i
    KSTestLognormal = maxD
End Function

Private Function KSTestUniform(ByRef sorted() As Double, ByVal n As Long, _
    ByVal a As Double, ByVal b As Double) As Double
    Dim maxD As Double, d As Double
    maxD = 0#
    Dim rng As Double
    rng = b - a
    If rng = 0 Then KSTestUniform = 1#: Exit Function

    Dim i As Long
    For i = 1 To n
        Dim empiricalCDF As Double
        empiricalCDF = CDbl(i) / CDbl(n)
        Dim theorCDF As Double
        theorCDF = (sorted(i) - a) / rng
        If theorCDF < 0 Then theorCDF = 0
        If theorCDF > 1 Then theorCDF = 1
        d = Abs(empiricalCDF - theorCDF)
        If d > maxD Then maxD = d
    Next i
    KSTestUniform = maxD
End Function

Private Function KSTestExponential(ByRef sorted() As Double, ByVal n As Long, _
    ByVal mean As Double) As Double
    Dim maxD As Double, d As Double
    maxD = 0#
    If mean = 0 Then KSTestExponential = 1#: Exit Function

    Dim i As Long
    For i = 1 To n
        Dim empiricalCDF As Double
        empiricalCDF = CDbl(i) / CDbl(n)
        Dim theorCDF As Double
        theorCDF = 1# - Exp(-sorted(i) / mean)
        d = Abs(empiricalCDF - theorCDF)
        If d > maxD Then maxD = d
    Next i
    KSTestExponential = maxD
End Function

'===============================================================================
' User-facing: Fit distribution to selected range
'===============================================================================
Public Sub FitDistributionToRange()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select data range:", ADDIN_NAME, , , , , , 8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    Dim data() As Double
    ReDim data(1 To rng.Cells.count)
    Dim i As Long
    For i = 1 To rng.Cells.count
        data(i) = CDbl(rng.Cells(i).Value)
    Next i

    Dim results() As FitResult
    results = FitDistributions(data)

    ' Display results
    Dim msg As String
    msg = "Distribution Fitting Results" & vbCrLf & _
          "===========================" & vbCrLf & vbCrLf
    For i = LBound(results) To UBound(results)
        msg = msg & "#" & results(i).Rank & " " & results(i).DistName
        msg = msg & " (KS=" & Format(results(i).KSStat, "0.0000") & ")"
        msg = msg & vbCrLf & "   Params: "
        Dim j As Long
        For j = 0 To results(i).ParamCount - 1
            If j > 0 Then msg = msg & ", "
            msg = msg & Format(results(i).Params(j), "0.0000")
        Next j
        msg = msg & vbCrLf & vbCrLf
    Next i

    MsgBox msg, vbInformation, ADDIN_NAME & " - Distribution Fitting"
End Sub
