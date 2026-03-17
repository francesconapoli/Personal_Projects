Attribute VB_Name = "modRandomGenerators"
'===============================================================================
' Module: modRandomGenerators
' Purpose: Core random number generation with Mersenne Twister and
'          inverse CDF methods for all distributions
'===============================================================================
Option Explicit

' ---- Mersenne Twister Constants ----
Private Const MT_N As Long = 624
Private Const MT_M As Long = 397
Private Const MATRIX_A As Long = &H9908B0DF
Private Const UPPER_MASK As Long = &H80000000
Private Const LOWER_MASK As Long = &H7FFFFFFF

Private mt(0 To MT_N - 1) As Long
Private mti As Long
Private mtInitialized As Boolean

' ---- Mathematical Constants ----
Private Const PI As Double = 3.14159265358979
Private Const E_CONST As Double = 2.71828182845905
Private Const LN2 As Double = 0.693147180559945
Private Const SQRT2PI As Double = 2.50662827463100

'===============================================================================
' Mersenne Twister Implementation
'===============================================================================
Public Sub MT_Initialize(ByVal seed As Long)
    mt(0) = seed And &H7FFFFFFF
    Dim i As Long
    For i = 1 To MT_N - 1
        ' Simplified seeding to avoid overflow
        mt(i) = (1812433253 * (mt(i - 1) Xor (mt(i - 1) \ 1073741824)) + i) And &H7FFFFFFF
    Next i
    mti = MT_N
    mtInitialized = True
End Sub

Public Function MT_Random() As Double
    ' Returns uniform random in (0, 1)
    If Not mtInitialized Then MT_Initialize CLng(Timer * 1000) Xor CLng(Now * 86400)

    Dim y As Long
    Dim kk As Long

    If mti >= MT_N Then
        For kk = 0 To MT_N - MT_M - 1
            y = (mt(kk) And UPPER_MASK) Or (mt(kk + 1) And LOWER_MASK)
            If (y And 1) = 1 Then
                mt(kk) = mt(kk + MT_M) Xor (y \ 2) Xor MATRIX_A
            Else
                mt(kk) = mt(kk + MT_M) Xor (y \ 2)
            End If
        Next kk
        For kk = MT_N - MT_M To MT_N - 2
            y = (mt(kk) And UPPER_MASK) Or (mt(kk + 1) And LOWER_MASK)
            If (y And 1) = 1 Then
                mt(kk) = mt(kk + (MT_M - MT_N)) Xor (y \ 2) Xor MATRIX_A
            Else
                mt(kk) = mt(kk + (MT_M - MT_N)) Xor (y \ 2)
            End If
        Next kk
        y = (mt(MT_N - 1) And UPPER_MASK) Or (mt(0) And LOWER_MASK)
        If (y And 1) = 1 Then
            mt(MT_N - 1) = mt(MT_M - 1) Xor (y \ 2) Xor MATRIX_A
        Else
            mt(MT_N - 1) = mt(MT_M - 1) Xor (y \ 2)
        End If
        mti = 0
    End If

    y = mt(mti)
    mti = mti + 1

    ' Tempering
    y = y Xor (y \ 2048)            ' >> 11
    y = y Xor ((y * 128) And &H9D2C5680)  ' << 7
    y = y Xor ((y * 32768) And &HEFC60000) ' << 15
    y = y Xor (y \ 262144)          ' >> 18

    MT_Random = (CDbl(y And &H7FFFFFFF) + 0.5) / 2147483648#
End Function

'===============================================================================
' Standard Normal via Box-Muller Transform
'===============================================================================
Private g_HasSpare As Boolean
Private g_Spare As Double

Public Function RandNormal01() As Double
    If g_HasSpare Then
        g_HasSpare = False
        RandNormal01 = g_Spare
        Exit Function
    End If

    Dim u As Double, v As Double, s As Double
    Do
        u = MT_Random() * 2# - 1#
        v = MT_Random() * 2# - 1#
        s = u * u + v * v
    Loop While s >= 1# Or s = 0#

    s = Sqr(-2# * Log(s) / s)
    g_Spare = v * s
    g_HasSpare = True
    RandNormal01 = u * s
End Function

'===============================================================================
' Normal CDF (Abramowitz & Stegun approximation)
'===============================================================================
Public Function NormalCDF(ByVal x As Double) As Double
    Dim t As Double, z As Double, p As Double
    z = Abs(x)
    t = 1# / (1# + 0.2316419 * z)
    p = t * (0.319381530# + t * (-0.356563782# + t * (1.781477937# + _
        t * (-1.821255978# + t * 1.330274429#))))
    p = 1# - SQRT2PI * Exp(-0.5 * z * z) * p
    ' Correction: need to divide by sqrt(2*pi) properly
    p = 1# - (1# / SQRT2PI) * Exp(-0.5 * z * z) * (t * (0.319381530# + t * (-0.356563782# + t * (1.781477937# + _
        t * (-1.821255978# + t * 1.330274429#)))))
    If x < 0 Then p = 1# - p
    NormalCDF = p
End Function

'===============================================================================
' Inverse Normal CDF (Rational Approximation - Peter Acklam)
'===============================================================================
Public Function InvNormalCDF(ByVal p As Double) As Double
    Const a1 As Double = -39.6968302866538
    Const a2 As Double = 220.946098424521
    Const a3 As Double = -275.928510446969
    Const a4 As Double = 138.357751867269
    Const a5 As Double = -30.6647980661472
    Const a6 As Double = 2.50662823884
    Const b1 As Double = -54.4760987982241
    Const b2 As Double = 161.585836858041
    Const b3 As Double = -155.698979859887
    Const b4 As Double = 66.8013118877198
    Const b5 As Double = -13.2806815528857
    Const c1 As Double = -0.00778489400243029
    Const c2 As Double = -0.322396458041136
    Const c3 As Double = -2.40075827716184
    Const c4 As Double = -2.54973253934373
    Const c5 As Double = 4.37466414146497
    Const c6 As Double = 2.93816398269878
    Const d1 As Double = 0.00778469570904146
    Const d2 As Double = 0.32246712907004
    Const d3 As Double = 2.445134137143
    Const d4 As Double = 3.75440866190742
    Const pLow As Double = 0.02425
    Const pHigh As Double = 0.97575

    Dim q As Double, r As Double

    If p <= 0 Then InvNormalCDF = -10: Exit Function
    If p >= 1 Then InvNormalCDF = 10: Exit Function

    If p < pLow Then
        q = Sqr(-2# * Log(p))
        InvNormalCDF = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                       ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#)
    ElseIf p <= pHigh Then
        q = p - 0.5
        r = q * q
        InvNormalCDF = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
                       (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1#)
    Else
        q = Sqr(-2# * Log(1# - p))
        InvNormalCDF = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                        ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#)
    End If
End Function

'===============================================================================
' Gamma function (Lanczos approximation)
'===============================================================================
Public Function GammaFunc(ByVal x As Double) As Double
    Dim g As Double, t As Double, w As Double
    Dim coef(0 To 6) As Double
    coef(0) = 1.000000000190015
    coef(1) = 76.18009172947146
    coef(2) = -86.50532032941677
    coef(3) = 24.01409824083091
    coef(4) = -1.231739572450155
    coef(5) = 0.001208650973866179
    coef(6) = -0.000005395239384953

    If x < 0.5 Then
        GammaFunc = PI / (Sin(PI * x) * GammaFunc(1# - x))
        Exit Function
    End If

    x = x - 1#
    g = 5.5
    t = x + g
    w = coef(0)
    Dim i As Long
    For i = 1 To 6
        w = w + coef(i) / (x + CDbl(i))
    Next i
    GammaFunc = Sqr(2# * PI) * (t ^ (x + 0.5)) * Exp(-t) * w
End Function

'===============================================================================
' Log Gamma function
'===============================================================================
Public Function LogGamma(ByVal x As Double) As Double
    LogGamma = Log(GammaFunc(x))
End Function

'===============================================================================
' Beta function
'===============================================================================
Public Function BetaFunc(ByVal a As Double, ByVal b As Double) As Double
    BetaFunc = GammaFunc(a) * GammaFunc(b) / GammaFunc(a + b)
End Function

'===============================================================================
' Incomplete Beta function (regularized) - continued fraction
'===============================================================================
Public Function RegBetaInc(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
    If x <= 0# Then RegBetaInc = 0#: Exit Function
    If x >= 1# Then RegBetaInc = 1#: Exit Function

    Dim bt As Double
    bt = Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b) + a * Log(x) + b * Log(1# - x))

    If x < (a + 1#) / (a + b + 2#) Then
        RegBetaInc = bt * BetaCF(x, a, b) / a
    Else
        RegBetaInc = 1# - bt * BetaCF(1# - x, b, a) / b
    End If
End Function

Private Function BetaCF(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
    Const MAXIT As Long = 200
    Const EPS As Double = 0.0000000003
    Const FPMIN As Double = 1E-30

    Dim m As Long, m2 As Long
    Dim aa As Double, c As Double, d As Double, del As Double, h As Double, qab As Double, qam As Double, qap As Double

    qab = a + b
    qap = a + 1#
    qam = a - 1#
    c = 1#
    d = 1# - qab * x / qap
    If Abs(d) < FPMIN Then d = FPMIN
    d = 1# / d
    h = d

    For m = 1 To MAXIT
        m2 = 2 * m
        aa = CDbl(m) * (b - CDbl(m)) * x / ((qam + CDbl(m2)) * (a + CDbl(m2)))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1# + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1# / d
        h = h * d * c

        aa = -(a + CDbl(m)) * (qab + CDbl(m)) * x / ((a + CDbl(m2)) * (qap + CDbl(m2)))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1# + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1# / d
        del = d * c
        h = h * del
        If Abs(del - 1#) < EPS Then Exit For
    Next m

    BetaCF = h
End Function

'===============================================================================
' Inverse Regularized Incomplete Beta (Newton-Raphson)
'===============================================================================
Public Function InvRegBetaInc(ByVal p As Double, ByVal a As Double, ByVal b As Double) As Double
    If p <= 0# Then InvRegBetaInc = 0#: Exit Function
    If p >= 1# Then InvRegBetaInc = 1#: Exit Function

    ' Initial guess using normal approximation
    Dim x As Double, lna As Double, lnb As Double
    Dim t As Double, u As Double, pp As Double
    Dim err As Double

    If a >= 1# And b >= 1# Then
        Dim yp As Double
        If p < 0.5 Then
            pp = p
        Else
            pp = 1# - p
        End If
        t = Sqr(-2# * Log(pp))
        x = (2.30753 + t * 0.27061) / (1# + t * (0.99229 + t * 0.04481)) - t
        If p < 0.5 Then x = -x
        Dim al As Double
        al = (x * x - 3#) / 6#
        Dim h As Double
        h = 2# / (1# / (2# * a - 1#) + 1# / (2# * b - 1#))
        Dim w As Double
        w = (x * Sqr(al + h) / h) - (1# / (2# * b - 1#) - 1# / (2# * a - 1#)) * (al + 5# / 6# - 2# / (3# * h))
        x = a / (a + b * Exp(2# * w))
    Else
        lna = Log(a / (a + b))
        lnb = Log(b / (a + b))
        t = Exp(a * lna) / a
        u = Exp(b * lnb) / b
        w = t + u
        If p < t / w Then
            x = (a * w * p) ^ (1# / a)
        Else
            x = 1# - (b * w * (1# - p)) ^ (1# / b)
        End If
    End If

    ' Newton-Raphson refinement
    Dim afac As Double
    afac = -LogGamma(a) - LogGamma(b) + LogGamma(a + b)
    Dim j As Long
    For j = 0 To 20
        If x = 0# Or x = 1# Then Exit For
        err = RegBetaInc(x, a, b) - p
        t = Exp((a - 1#) * Log(x) + (b - 1#) * Log(1# - x) + afac)
        u = err / t
        x = x - u
        If x <= 0# Then x = 0.5 * (x + u)
        If x >= 1# Then x = 0.5 * (x + u + 1#)
        If Abs(u) < 0.0000001 * x Then Exit For
    Next j

    InvRegBetaInc = x
End Function

'===============================================================================
' Incomplete Gamma function (regularized lower)
'===============================================================================
Public Function RegGammaInc(ByVal a As Double, ByVal x As Double) As Double
    If x < 0# Then RegGammaInc = 0#: Exit Function
    If x = 0# Then RegGammaInc = 0#: Exit Function

    If x < a + 1# Then
        ' Series expansion
        Dim ap As Double, del As Double, s As Double
        ap = a
        del = 1# / a
        s = del
        Dim n As Long
        For n = 1 To 200
            ap = ap + 1#
            del = del * x / ap
            s = s + del
            If Abs(del) < Abs(s) * 0.0000000003 Then Exit For
        Next n
        RegGammaInc = s * Exp(-x + a * Log(x) - LogGamma(a))
    Else
        ' Continued fraction
        Dim b0 As Double, c As Double, d As Double, h As Double
        b0 = x + 1# - a
        c = 1E+30
        d = 1# / b0
        h = d
        Dim i As Long
        For i = 1 To 200
            Dim an As Double, bn As Double
            an = -CDbl(i) * (CDbl(i) - a)
            bn = x + CDbl(2 * i) + 1# - a
            d = an * d + bn
            If Abs(d) < 1E-30 Then d = 1E-30
            c = bn + an / c
            If Abs(c) < 1E-30 Then c = 1E-30
            d = 1# / d
            del = d * c
            h = h * del
            If Abs(del - 1#) < 0.0000000003 Then Exit For
        Next i
        RegGammaInc = 1# - Exp(-x + a * Log(x) - LogGamma(a)) * h
    End If
End Function

'===============================================================================
' Inverse Incomplete Gamma (Newton-Raphson)
'===============================================================================
Public Function InvRegGammaInc(ByVal a As Double, ByVal p As Double) As Double
    If p <= 0# Then InvRegGammaInc = 0#: Exit Function
    If p >= 1# Then InvRegGammaInc = a * 10#: Exit Function

    ' Initial guess
    Dim x As Double, t As Double, lna As Double
    Dim gln As Double
    gln = LogGamma(a)

    If a > 1# Then
        lna = Log(a - 1#)
        t = InvNormalCDF(p)
        Dim s As Double
        s = 1# / (3# * Sqr(a))
        x = a * (1# - s + t * Sqr(s)) ^ 3
        If x < 0# Then x = 0.01
    Else
        t = 1# - a * (0.253 + a * 0.12)
        If p < t Then
            x = (p / t) ^ (1# / a)
        Else
            x = 1# - Log(1# - (p - t) / (1# - t))
        End If
    End If

    ' Newton-Raphson
    Dim j As Long, err As Double, u As Double
    For j = 0 To 20
        If x <= 0# Then InvRegGammaInc = 0#: Exit Function
        err = RegGammaInc(a, x) - p
        If a > 1# Then
            t = Exp((a - 1#) * Log(x) - x - gln)
        Else
            t = Exp(-x + (a - 1#) * Log(x) - gln)
        End If
        u = err / t
        x = x - u
        If x <= 0# Then x = 0.5 * (x + u)
        If Abs(u) < 0.0000001 * x Then Exit For
    Next j

    InvRegGammaInc = x
End Function

'===============================================================================
' Random Gamma variate (Marsaglia & Tsang)
'===============================================================================
Public Function RandGamma(ByVal shape As Double) As Double
    If shape < 1# Then
        RandGamma = RandGamma(shape + 1#) * MT_Random() ^ (1# / shape)
        Exit Function
    End If

    Dim d As Double, c As Double, x As Double, v As Double, u As Double
    d = shape - 1# / 3#
    c = 1# / Sqr(9# * d)

    Do
        Do
            x = RandNormal01()
            v = 1# + c * x
        Loop While v <= 0#
        v = v * v * v
        u = MT_Random()
        If u < 1# - 0.0331 * (x * x) * (x * x) Then
            RandGamma = d * v
            Exit Function
        End If
        If Log(u) < 0.5 * x * x + d * (1# - v + Log(v)) Then
            RandGamma = d * v
            Exit Function
        End If
    Loop
End Function
