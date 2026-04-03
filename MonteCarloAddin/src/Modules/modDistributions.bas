Attribute VB_Name = "modDistributions"
'===============================================================================
' Module: modDistributions
' Purpose: All @RISK-compatible distribution worksheet functions
'          Each function can be called from cells: =MCNormal(100, 15)
'          During simulation, returns sampled value; otherwise returns mean/mode
'===============================================================================
Option Explicit

'===============================================================================
' CONTINUOUS DISTRIBUTIONS
'===============================================================================

'--- Normal Distribution ---
Public Function MCNormal(ByVal mean As Double, ByVal stdDev As Double) As Variant
Attribute MCNormal.VB_Description = "Normal distribution. Returns random sample during simulation."
    If g_SimState = ssRunning Then
        MCNormal = mean + stdDev * RandNormal01()
        RegisterInputAuto Application.Caller, dtNormal, mean, stdDev
    Else
        MCNormal = mean
    End If
End Function

'--- Uniform Distribution ---
Public Function MCUniform(ByVal minVal As Double, ByVal maxVal As Double) As Variant
Attribute MCUniform.VB_Description = "Uniform distribution between min and max."
    If g_SimState = ssRunning Then
        MCUniform = minVal + (maxVal - minVal) * MT_Random()
        RegisterInputAuto Application.Caller, dtUniform, minVal, maxVal
    Else
        MCUniform = (minVal + maxVal) / 2#
    End If
End Function

'--- Triangular Distribution ---
Public Function MCTriang(ByVal minVal As Double, ByVal mostLikely As Double, ByVal maxVal As Double) As Variant
Attribute MCTriang.VB_Description = "Triangular distribution with min, most likely, max."
    If g_SimState = ssRunning Then
        Dim u As Double, F_c As Double
        u = MT_Random()
        F_c = (mostLikely - minVal) / (maxVal - minVal)
        If u < F_c Then
            MCTriang = minVal + Sqr(u * (maxVal - minVal) * (mostLikely - minVal))
        Else
            MCTriang = maxVal - Sqr((1# - u) * (maxVal - minVal) * (maxVal - mostLikely))
        End If
        RegisterInputAuto Application.Caller, dtTriangular, minVal, mostLikely, maxVal
    Else
        MCTriang = mostLikely
    End If
End Function

'--- Lognormal Distribution ---
Public Function MCLognorm(ByVal mean As Double, ByVal stdDev As Double) As Variant
Attribute MCLognorm.VB_Description = "Lognormal distribution with arithmetic mean and std dev."
    If g_SimState = ssRunning Then
        ' Convert arithmetic mean/stddev to log-space params
        Dim variance As Double, mu As Double, sigma As Double
        variance = stdDev * stdDev
        sigma = Sqr(Log(1# + variance / (mean * mean)))
        mu = Log(mean) - 0.5 * sigma * sigma
        MCLognorm = Exp(mu + sigma * RandNormal01())
        RegisterInputAuto Application.Caller, dtLognormal, mean, stdDev
    Else
        MCLognorm = mean
    End If
End Function

'--- Exponential Distribution ---
Public Function MCExpon(ByVal beta As Double) As Variant
Attribute MCExpon.VB_Description = "Exponential distribution with rate parameter beta."
    If g_SimState = ssRunning Then
        MCExpon = -beta * Log(MT_Random())
        RegisterInputAuto Application.Caller, dtExponential, beta
    Else
        MCExpon = beta
    End If
End Function

'--- Beta Distribution ---
Public Function MCBeta(ByVal alpha1 As Double, ByVal alpha2 As Double) As Variant
Attribute MCBeta.VB_Description = "Beta distribution on [0,1] with shape parameters."
    If g_SimState = ssRunning Then
        Dim x As Double, y As Double
        x = RandGamma(alpha1)
        y = RandGamma(alpha2)
        MCBeta = x / (x + y)
        RegisterInputAuto Application.Caller, dtBeta, alpha1, alpha2
    Else
        MCBeta = alpha1 / (alpha1 + alpha2)
    End If
End Function

'--- Beta General Distribution ---
Public Function MCBetaGeneral(ByVal alpha1 As Double, ByVal alpha2 As Double, _
    ByVal minVal As Double, ByVal maxVal As Double) As Variant
Attribute MCBetaGeneral.VB_Description = "Beta distribution on [min, max] with shape parameters."
    If g_SimState = ssRunning Then
        Dim x As Double, y As Double, bv As Double
        x = RandGamma(alpha1)
        y = RandGamma(alpha2)
        bv = x / (x + y)
        MCBetaGeneral = minVal + (maxVal - minVal) * bv
        RegisterInputAuto Application.Caller, dtBetaGeneral, alpha1, alpha2, minVal, maxVal
    Else
        MCBetaGeneral = minVal + (maxVal - minVal) * alpha1 / (alpha1 + alpha2)
    End If
End Function

'--- Gamma Distribution ---
Public Function MCGamma(ByVal alpha As Double, ByVal beta As Double) As Variant
Attribute MCGamma.VB_Description = "Gamma distribution with shape alpha and scale beta."
    If g_SimState = ssRunning Then
        MCGamma = RandGamma(alpha) * beta
        RegisterInputAuto Application.Caller, dtGamma, alpha, beta
    Else
        MCGamma = alpha * beta
    End If
End Function

'--- Weibull Distribution ---
Public Function MCWeibull(ByVal shape As Double, ByVal scale As Double) As Variant
Attribute MCWeibull.VB_Description = "Weibull distribution with shape and scale parameters."
    If g_SimState = ssRunning Then
        MCWeibull = scale * (-Log(MT_Random())) ^ (1# / shape)
        RegisterInputAuto Application.Caller, dtWeibull, shape, scale
    Else
        MCWeibull = scale * GammaFunc(1# + 1# / shape)
    End If
End Function

'--- PERT Distribution ---
Public Function MCPERT(ByVal minVal As Double, ByVal mostLikely As Double, ByVal maxVal As Double) As Variant
Attribute MCPERT.VB_Description = "PERT distribution (modified beta) with min, most likely, max."
    If g_SimState = ssRunning Then
        Dim mean As Double, alpha1 As Double, alpha2 As Double
        mean = (minVal + 4# * mostLikely + maxVal) / 6#
        If mean = mostLikely Then
            alpha1 = 3#
        Else
            alpha1 = ((mean - minVal) * (2# * mostLikely - minVal - maxVal)) / _
                     ((mostLikely - mean) * (maxVal - minVal))
        End If
        If alpha1 < 0.001 Then alpha1 = 0.001
        alpha2 = alpha1 * (maxVal - mean) / (mean - minVal)
        If alpha2 < 0.001 Then alpha2 = 0.001

        Dim x As Double, y As Double, bv As Double
        x = RandGamma(alpha1)
        y = RandGamma(alpha2)
        bv = x / (x + y)
        MCPERT = minVal + (maxVal - minVal) * bv
        RegisterInputAuto Application.Caller, dtPERT, minVal, mostLikely, maxVal
    Else
        MCPERT = (minVal + 4# * mostLikely + maxVal) / 6#
    End If
End Function

'--- Pareto Distribution ---
Public Function MCPareto(ByVal shape As Double, ByVal scale As Double) As Variant
Attribute MCPareto.VB_Description = "Pareto distribution with shape (theta) and scale (xm)."
    If g_SimState = ssRunning Then
        MCPareto = scale / (MT_Random() ^ (1# / shape))
        RegisterInputAuto Application.Caller, dtPareto, shape, scale
    Else
        If shape > 1# Then
            MCPareto = shape * scale / (shape - 1#)
        Else
            MCPareto = scale
        End If
    End If
End Function

'--- Logistic Distribution ---
Public Function MCLogistic(ByVal mean As Double, ByVal scale As Double) As Variant
Attribute MCLogistic.VB_Description = "Logistic distribution with mean and scale."
    If g_SimState = ssRunning Then
        Dim u As Double
        u = MT_Random()
        MCLogistic = mean + scale * Log(u / (1# - u))
        RegisterInputAuto Application.Caller, dtLogistic, mean, scale
    Else
        MCLogistic = mean
    End If
End Function

'--- Rayleigh Distribution ---
Public Function MCRayleigh(ByVal scale As Double) As Variant
Attribute MCRayleigh.VB_Description = "Rayleigh distribution with scale parameter."
    If g_SimState = ssRunning Then
        MCRayleigh = scale * Sqr(-2# * Log(MT_Random()))
        RegisterInputAuto Application.Caller, dtRayleigh, scale
    Else
        MCRayleigh = scale * Sqr(PI / 2#)
    End If
End Function

'--- Extreme Value (Gumbel Max) ---
Public Function MCExtValue(ByVal alpha As Double, ByVal beta As Double) As Variant
Attribute MCExtValue.VB_Description = "Extreme Value (Gumbel Maximum) distribution."
    If g_SimState = ssRunning Then
        MCExtValue = alpha - beta * Log(-Log(MT_Random()))
        RegisterInputAuto Application.Caller, dtExtValue, alpha, beta
    Else
        MCExtValue = alpha + beta * 0.5772156649 ' Euler-Mascheroni
    End If
End Function

'--- Extreme Value Min ---
Public Function MCExtValueMin(ByVal alpha As Double, ByVal beta As Double) As Variant
Attribute MCExtValueMin.VB_Description = "Extreme Value Minimum distribution."
    If g_SimState = ssRunning Then
        MCExtValueMin = alpha + beta * Log(-Log(MT_Random()))
        RegisterInputAuto Application.Caller, dtExtValueMin, alpha, beta
    Else
        MCExtValueMin = alpha - beta * 0.5772156649
    End If
End Function

'--- Chi-Squared Distribution ---
Public Function MCChiSq(ByVal df As Double) As Variant
Attribute MCChiSq.VB_Description = "Chi-Squared distribution with df degrees of freedom."
    If g_SimState = ssRunning Then
        MCChiSq = RandGamma(df / 2#) * 2#
        RegisterInputAuto Application.Caller, dtChiSquared, df
    Else
        MCChiSq = df
    End If
End Function

'--- Student's t Distribution ---
Public Function MCStudentT(ByVal df As Double) As Variant
Attribute MCStudentT.VB_Description = "Student's t-distribution with df degrees of freedom."
    If g_SimState = ssRunning Then
        Dim z As Double, chi2 As Double
        z = RandNormal01()
        chi2 = RandGamma(df / 2#) * 2#
        MCStudentT = z / Sqr(chi2 / df)
        RegisterInputAuto Application.Caller, dtStudentT, df
    Else
        MCStudentT = 0#
    End If
End Function

'===============================================================================
' DISCRETE DISTRIBUTIONS
'===============================================================================

'--- Bernoulli Distribution ---
Public Function MCBernoulli(ByVal p As Double) As Variant
Attribute MCBernoulli.VB_Description = "Bernoulli distribution returning 1 with probability p, 0 otherwise."
    If g_SimState = ssRunning Then
        If MT_Random() < p Then
            MCBernoulli = 1
        Else
            MCBernoulli = 0
        End If
        RegisterInputAuto Application.Caller, dtBernoulli, p, 0#
    Else
        MCBernoulli = CLng(p >= 0.5)
    End If
End Function

'--- Binomial Distribution ---
Public Function MCBinomial(ByVal n As Long, ByVal p As Double) As Variant
Attribute MCBinomial.VB_Description = "Binomial distribution with n trials and probability p."
    If g_SimState = ssRunning Then
        Dim count As Long, i As Long
        count = 0
        ' For large n use normal approximation
        If n > 100 Then
            Dim x As Double
            x = n * p + Sqr(n * p * (1# - p)) * RandNormal01()
            MCBinomial = CLng(Application.WorksheetFunction.Max(0, Application.WorksheetFunction.Min(n, Int(x + 0.5))))
        Else
            For i = 1 To n
                If MT_Random() < p Then count = count + 1
            Next i
            MCBinomial = count
        End If
        RegisterInputAuto Application.Caller, dtBinomial, CDbl(n), p
    Else
        MCBinomial = CLng(n * p)
    End If
End Function

'--- Poisson Distribution ---
Public Function MCPoisson(ByVal lambda As Double) As Variant
Attribute MCPoisson.VB_Description = "Poisson distribution with mean lambda."
    If g_SimState = ssRunning Then
        If lambda < 30 Then
            ' Direct method
            Dim L As Double, k As Long, p As Double
            L = Exp(-lambda)
            k = 0
            p = 1#
            Do
                k = k + 1
                p = p * MT_Random()
            Loop While p > L
            MCPoisson = k - 1
        Else
            ' Normal approximation for large lambda
            Dim x As Double
            x = lambda + Sqr(lambda) * RandNormal01()
            MCPoisson = CLng(Application.WorksheetFunction.Max(0, Int(x + 0.5)))
        End If
        RegisterInputAuto Application.Caller, dtPoisson, lambda
    Else
        MCPoisson = CLng(lambda)
    End If
End Function

'--- Negative Binomial ---
Public Function MCNegBinom(ByVal s As Long, ByVal p As Double) As Variant
Attribute MCNegBinom.VB_Description = "Negative Binomial distribution."
    If g_SimState = ssRunning Then
        ' Gamma-Poisson mixture
        Dim g As Double
        g = RandGamma(CDbl(s)) * (1# - p) / p
        If g < 30 Then
            Dim L As Double, k As Long, pr As Double
            L = Exp(-g)
            k = 0
            pr = 1#
            Do
                k = k + 1
                pr = pr * MT_Random()
            Loop While pr > L
            MCNegBinom = k - 1
        Else
            MCNegBinom = CLng(Application.WorksheetFunction.Max(0, Int(g + Sqr(g) * RandNormal01() + 0.5)))
        End If
        RegisterInputAuto Application.Caller, dtNegBinomial, CDbl(s), p
    Else
        MCNegBinom = CLng(s * (1# - p) / p)
    End If
End Function

'--- Discrete Distribution ---
Public Function MCDiscrete(ByVal values As Variant, ByVal probabilities As Variant) As Variant
Attribute MCDiscrete.VB_Description = "Custom discrete distribution with values and probabilities."
    Dim vals() As Double, probs() As Double
    Dim n As Long, i As Long

    ' Convert from Range or Array
    vals = ConvertToDoubleArray(values)
    probs = ConvertToDoubleArray(probabilities)
    n = UBound(vals) - LBound(vals) + 1

    If g_SimState = ssRunning Then
        Dim u As Double, cumP As Double
        u = MT_Random()
        cumP = 0#
        For i = LBound(probs) To UBound(probs)
            cumP = cumP + probs(i)
            If u <= cumP Then
                MCDiscrete = vals(i)
                Exit Function
            End If
        Next i
        MCDiscrete = vals(UBound(vals))
    Else
        ' Return expected value
        Dim ev As Double
        ev = 0#
        For i = LBound(vals) To UBound(vals)
            ev = ev + vals(i) * probs(i)
        Next i
        MCDiscrete = ev
    End If
End Function

'--- Discrete Uniform ---
Public Function MCDUniform(ByVal values As Variant) As Variant
Attribute MCDUniform.VB_Description = "Discrete uniform distribution - equally likely values."
    Dim vals() As Double
    vals = ConvertToDoubleArray(values)
    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1

    If g_SimState = ssRunning Then
        Dim idx As Long
        idx = LBound(vals) + Int(MT_Random() * n)
        If idx > UBound(vals) Then idx = UBound(vals)
        MCDUniform = vals(idx)
    Else
        ' Return first value or mean
        Dim s As Double, i As Long
        s = 0
        For i = LBound(vals) To UBound(vals)
            s = s + vals(i)
        Next i
        MCDUniform = s / n
    End If
End Function

'--- Integer Uniform ---
Public Function MCIntUniform(ByVal minVal As Long, ByVal maxVal As Long) As Variant
Attribute MCIntUniform.VB_Description = "Integer uniform distribution from min to max."
    If g_SimState = ssRunning Then
        MCIntUniform = minVal + Int(MT_Random() * (maxVal - minVal + 1))
        RegisterInputAuto Application.Caller, dtIntUniform, CDbl(minVal), CDbl(maxVal)
    Else
        MCIntUniform = CLng((minVal + maxVal) / 2)
    End If
End Function

'===============================================================================
' EMPIRICAL DISTRIBUTIONS
'===============================================================================

'--- Cumulative Distribution ---
Public Function MCCumul(ByVal xValues As Variant, ByVal probValues As Variant) As Variant
Attribute MCCumul.VB_Description = "Cumulative ascending distribution from (x, cumProb) pairs."
    Dim xv() As Double, pv() As Double
    xv = ConvertToDoubleArray(xValues)
    pv = ConvertToDoubleArray(probValues)

    If g_SimState = ssRunning Then
        Dim u As Double
        u = MT_Random()
        Dim i As Long
        For i = LBound(pv) + 1 To UBound(pv)
            If u <= pv(i) Then
                ' Linear interpolation
                Dim frac As Double
                If pv(i) = pv(i - 1) Then
                    frac = 0.5
                Else
                    frac = (u - pv(i - 1)) / (pv(i) - pv(i - 1))
                End If
                MCCumul = xv(i - 1) + frac * (xv(i) - xv(i - 1))
                Exit Function
            End If
        Next i
        MCCumul = xv(UBound(xv))
    Else
        MCCumul = xv(LBound(xv) + (UBound(xv) - LBound(xv)) \ 2)
    End If
End Function

'--- Cumulative Descending ---
Public Function MCCumulD(ByVal xValues As Variant, ByVal probValues As Variant) As Variant
Attribute MCCumulD.VB_Description = "Cumulative descending distribution."
    Dim xv() As Double, pv() As Double
    xv = ConvertToDoubleArray(xValues)
    pv = ConvertToDoubleArray(probValues)

    ' Convert descending to ascending
    Dim n As Long, i As Long
    n = UBound(pv) - LBound(pv) + 1
    For i = LBound(pv) To UBound(pv)
        pv(i) = 1# - pv(i)
    Next i

    MCCumulD = MCCumul(xv, pv)
End Function

'--- Histogram Distribution ---
Public Function MCHistogrm(ByVal minVal As Double, ByVal maxVal As Double, ByVal binProbs As Variant) As Variant
Attribute MCHistogrm.VB_Description = "Histogram distribution with min, max, and bin probabilities."
    Dim probs() As Double
    probs = ConvertToDoubleArray(binProbs)
    Dim nBins As Long
    nBins = UBound(probs) - LBound(probs) + 1

    If g_SimState = ssRunning Then
        ' Normalize probabilities
        Dim total As Double, i As Long
        total = 0#
        For i = LBound(probs) To UBound(probs)
            total = total + probs(i)
        Next i

        Dim u As Double, cumP As Double, binIdx As Long
        u = MT_Random()
        cumP = 0#
        binIdx = 0
        For i = LBound(probs) To UBound(probs)
            cumP = cumP + probs(i) / total
            If u <= cumP Then
                binIdx = i - LBound(probs)
                Exit For
            End If
        Next i

        ' Uniform within selected bin
        Dim binWidth As Double
        binWidth = (maxVal - minVal) / nBins
        MCHistogrm = minVal + binWidth * (binIdx + MT_Random())
    Else
        MCHistogrm = (minVal + maxVal) / 2#
    End If
End Function

'--- Compound Distribution ---
Public Function MCCompound(ByVal freqCell As Range, ByVal sevCell As Range) As Variant
Attribute MCCompound.VB_Description = "Compound distribution: sum of N severity draws where N is from frequency."
    If g_SimState = ssRunning Then
        Dim n As Long, i As Long, total As Double
        ' Force recalc of frequency cell to get a count
        n = CLng(freqCell.Value)
        total = 0#
        For i = 1 To n
            Application.Calculate
            total = total + CDbl(sevCell.Value)
        Next i
        MCCompound = total
    Else
        MCCompound = CDbl(freqCell.Value) * CDbl(sevCell.Value)
    End If
End Function

'--- SimTable (for scenario analysis) ---
Public Function MCSimtable(ByVal values As Variant) As Variant
Attribute MCSimtable.VB_Description = "Returns value corresponding to current simulation number."
    Dim vals() As Double
    vals = ConvertToDoubleArray(values)

    If g_SimState = ssRunning Then
        Dim idx As Long
        idx = g_CurrentSimulation - 1 + LBound(vals)
        If idx > UBound(vals) Then idx = UBound(vals)
        If idx < LBound(vals) Then idx = LBound(vals)
        MCSimtable = vals(idx)
    Else
        MCSimtable = vals(LBound(vals))
    End If
End Function

'===============================================================================
' HELPER: Auto-register input during simulation
'===============================================================================
Private Sub RegisterInputAuto(ByVal caller As Variant, ByVal dType As DistributionType, ParamArray params() As Variant)
    On Error Resume Next
    If TypeName(caller) <> "Range" Then Exit Sub

    Dim rng As Range
    Set rng = caller
    Dim addr As String
    addr = rng.Worksheet.Name & "!" & rng.Address

    ' Check if already registered
    Dim i As Long
    For i = 0 To g_InputCount - 1
        If g_Inputs(i).SheetName & "!" & g_Inputs(i).CellAddress = addr Then Exit Sub
    Next i

    ' Register new input
    If g_InputCount = 0 Then
        ReDim g_Inputs(0 To 0)
    Else
        ReDim Preserve g_Inputs(0 To g_InputCount)
    End If
    With g_Inputs(g_InputCount)
        .CellAddress = rng.Address
        .SheetName = rng.Worksheet.Name
        .DistType = dType
        .InputName = rng.Address
        .CorrelationGroup = -1
        ReDim .Params(0 To UBound(params))
        Dim j As Long
        For j = 0 To UBound(params)
            .Params(j) = CDbl(params(j))
        Next j
    End With
    g_InputCount = g_InputCount + 1
    On Error GoTo 0
End Sub

'===============================================================================
' HELPER: Convert Range or Array to Double array
'===============================================================================
Public Function ConvertToDoubleArray(ByVal v As Variant) As Double()
    Dim result() As Double
    Dim i As Long, n As Long

    If TypeName(v) = "Range" Then
        Dim rng As Range
        Set rng = v
        n = rng.Cells.count
        ReDim result(1 To n)
        For i = 1 To n
            result(i) = CDbl(rng.Cells(i).Value)
        Next i
    ElseIf IsArray(v) Then
        If LBound(v) = 0 Then
            n = UBound(v) - LBound(v) + 1
            ReDim result(0 To UBound(v))
            For i = LBound(v) To UBound(v)
                result(i) = CDbl(v(i))
            Next i
        Else
            n = UBound(v) - LBound(v) + 1
            ReDim result(LBound(v) To UBound(v))
            For i = LBound(v) To UBound(v)
                result(i) = CDbl(v(i))
            Next i
        End If
    Else
        ReDim result(1 To 1)
        result(1) = CDbl(v)
    End If

    ConvertToDoubleArray = result
End Function
