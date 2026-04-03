Attribute VB_Name = "modSimEngine"
'===============================================================================
' Module: modSimEngine
' Purpose: Core simulation engine - runs iterations, collects results,
'          handles Latin Hypercube sampling, convergence monitoring
'===============================================================================
Option Explicit

Private m_CancelRequested As Boolean

'===============================================================================
' Main Simulation Entry Point
'===============================================================================
Public Sub RunSimulation()
    If g_SimState = ssRunning Then
        MsgBox "Simulation is already running.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    ' Initialize
    On Error GoTo SimError
    m_CancelRequested = False

    ' Auto-detect inputs and outputs by scanning the active workbook
    Call ScanWorkbookForFunctions

    If g_OutputCount = 0 Then
        MsgBox "No output cells found. Use MCOutput() or the Add Output button to define outputs.", _
               vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    ' Backup application state
    g_OrigScreenUpdating = Application.ScreenUpdating
    g_OrigCalculation = Application.Calculation
    g_OrigEvents = Application.EnableEvents
    g_OrigStatusBar = Application.DisplayStatusBar

    ' Optimize for speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    ' Save original cell values
    SaveOriginalValues

    ' Initialize random generator
    If g_Settings.UseFixedSeed Then
        MT_Initialize g_Settings.RandomSeed
    Else
        MT_Initialize CLng(Timer * 1000) Xor CLng(Now * 86400000#)
    End If

    ' Initialize output data arrays
    Dim i As Long
    For i = 0 To g_OutputCount - 1
        ReDim g_Outputs(i).IterationData(1 To g_Settings.Iterations)
    Next i

    ' Generate Latin Hypercube permutations if needed
    If g_Settings.SampMethod = smLatinHypercube And g_InputCount > 0 Then
        GenerateLHSPermutations
    End If

    g_SimState = ssRunning
    g_SimStartTime = Timer

    ' Run simulation(s)
    Dim simNum As Long
    For simNum = 1 To g_Settings.Simulations
        g_CurrentSimulation = simNum
        RunSingleSimulation
        If m_CancelRequested Then Exit For
    Next simNum

    g_SimEndTime = Timer

    ' Compute statistics for all outputs
    For i = 0 To g_OutputCount - 1
        ComputeOutputStatistics i
    Next i

    ' Restore original values
    RestoreOriginalValues

    g_SimState = ssComplete

    ' Restore application state
    Application.ScreenUpdating = g_OrigScreenUpdating
    Application.Calculation = g_OrigCalculation
    Application.EnableEvents = g_OrigEvents
    Application.Calculate

    ' Show results
    Dim elapsed As Double
    elapsed = g_SimEndTime - g_SimStartTime
    Application.StatusBar = "Simulation complete: " & g_Settings.Iterations & " iterations in " & _
                           Format(elapsed, "0.00") & " seconds"

    ' Show results form
    frmResults.Show vbModeless
    Exit Sub

SimError:
    g_SimState = ssStopped
    RestoreOriginalValues
    Application.ScreenUpdating = g_OrigScreenUpdating
    Application.Calculation = g_OrigCalculation
    Application.EnableEvents = g_OrigEvents
    Application.StatusBar = False
    MsgBox "Simulation error: " & Err.Description, vbCritical, ADDIN_NAME
End Sub

'===============================================================================
' Run a single simulation (all iterations)
'===============================================================================
Private Sub RunSingleSimulation()
    Dim iter As Long
    Dim ws As Worksheet
    Dim rng As Range
    Dim outputVal As Double

    For iter = 1 To g_Settings.Iterations
        g_CurrentIteration = iter

        ' Apply Latin Hypercube stratified samples if needed
        If g_Settings.SampMethod = smLatinHypercube And g_InputCount > 0 Then
            ApplyLHSSample iter
        End If

        ' Recalculate the workbook
        Application.Calculate
        DoEvents

        ' Collect output values
        Dim i As Long
        For i = 0 To g_OutputCount - 1
            On Error Resume Next
            Set ws = ThisWorkbook.Parent.Worksheets(g_Outputs(i).SheetName)
            If ws Is Nothing Then
                Set ws = ActiveWorkbook.Worksheets(g_Outputs(i).SheetName)
            End If
            Set rng = ws.Range(g_Outputs(i).CellAddress)
            outputVal = CDbl(rng.Value)
            If Err.Number <> 0 Then outputVal = 0#: Err.Clear
            On Error GoTo 0
            g_Outputs(i).IterationData(iter) = outputVal
        Next i

        ' Status update
        If iter Mod g_Settings.UpdateFrequency = 0 Then
            Application.StatusBar = "Simulation " & g_CurrentSimulation & "/" & g_Settings.Simulations & _
                                   " - Iteration " & iter & "/" & g_Settings.Iterations & _
                                   " (" & Format(iter / g_Settings.Iterations * 100, "0") & "%)"
            DoEvents
            If m_CancelRequested Then Exit For
        End If

        ' Check convergence
        If g_Settings.ConvergenceEnabled And iter > 100 Then
            If iter Mod g_Settings.ConvergenceTestEvery = 0 Then
                If CheckConvergence(iter) Then
                    g_ConvergenceReached = True
                    ' Trim data arrays
                    For i = 0 To g_OutputCount - 1
                        ReDim Preserve g_Outputs(i).IterationData(1 To iter)
                    Next i
                    Exit For
                End If
            End If
        End If
    Next iter
End Sub

'===============================================================================
' Scan workbook for MC* functions to auto-detect inputs/outputs
'===============================================================================
Public Sub ScanWorkbookForFunctions()
    Dim ws As Worksheet
    Dim cell As Range
    Dim formula As String

    ' Reset
    g_InputCount = 0
    g_OutputCount = 0
    Erase g_Inputs
    Erase g_Outputs

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    For Each ws In wb.Worksheets
        Dim usedRange As Range
        Set usedRange = ws.UsedRange
        If Not usedRange Is Nothing Then
            For Each cell In usedRange.Cells
                If cell.HasFormula Then
                    formula = UCase(cell.formula)

                    ' Check for output functions
                    If InStr(formula, "MCOUTPUT") > 0 Or InStr(formula, "MCADDOUTPUT") > 0 Then
                        RegisterOutput cell
                    End If

                    ' Check for distribution functions (inputs auto-register during sim)
                    ' But we track them here for UI display
                    If ContainsDistFunction(formula) Then
                        RegisterScanInput cell, formula
                    End If
                End If
            Next cell
        End If
    Next ws
End Sub

Private Function ContainsDistFunction(ByVal formula As String) As Boolean
    ContainsDistFunction = (InStr(formula, "MCNORMAL") > 0 Or _
                           InStr(formula, "MCUNIFORM") > 0 Or _
                           InStr(formula, "MCTRIANG") > 0 Or _
                           InStr(formula, "MCLOGNORM") > 0 Or _
                           InStr(formula, "MCEXPON") > 0 Or _
                           InStr(formula, "MCBETA") > 0 Or _
                           InStr(formula, "MCGAMMA") > 0 Or _
                           InStr(formula, "MCWEIBULL") > 0 Or _
                           InStr(formula, "MCPERT") > 0 Or _
                           InStr(formula, "MCPARETO") > 0 Or _
                           InStr(formula, "MCLOGISTIC") > 0 Or _
                           InStr(formula, "MCRAYLEIGH") > 0 Or _
                           InStr(formula, "MCEXTVALUE") > 0 Or _
                           InStr(formula, "MCCHISQ") > 0 Or _
                           InStr(formula, "MCSTUDENTT") > 0 Or _
                           InStr(formula, "MCBINOMIAL") > 0 Or _
                           InStr(formula, "MCPOISSON") > 0 Or _
                           InStr(formula, "MCNEGBINOM") > 0 Or _
                           InStr(formula, "MCDISCRETE") > 0 Or _
                           InStr(formula, "MCDUNIFORM") > 0 Or _
                           InStr(formula, "MCINTUNIFORM") > 0 Or _
                           InStr(formula, "MCCUMUL") > 0 Or _
                           InStr(formula, "MCHISTOGRM") > 0 Or _
                           InStr(formula, "MCBETAGENERAL") > 0)
End Function

Private Sub RegisterScanInput(ByVal cell As Range, ByVal formula As String)
    ' Check if already registered
    Dim i As Long
    For i = 0 To g_InputCount - 1
        If g_Inputs(i).SheetName = cell.Worksheet.Name And _
           g_Inputs(i).CellAddress = cell.Address Then Exit Sub
    Next i

    If g_InputCount = 0 Then
        ReDim g_Inputs(0 To 0)
    Else
        ReDim Preserve g_Inputs(0 To g_InputCount)
    End If

    With g_Inputs(g_InputCount)
        .CellAddress = cell.Address
        .SheetName = cell.Worksheet.Name
        .InputName = cell.Address
        .OriginalValue = cell.Value
        .DistType = DetectDistType(formula)
        .CorrelationGroup = -1
    End With
    g_InputCount = g_InputCount + 1
End Sub

Private Function DetectDistType(ByVal formula As String) As DistributionType
    If InStr(formula, "MCNORMAL") > 0 Then DetectDistType = dtNormal
    If InStr(formula, "MCUNIFORM") > 0 Then DetectDistType = dtUniform
    If InStr(formula, "MCTRIANG") > 0 Then DetectDistType = dtTriangular
    If InStr(formula, "MCLOGNORM") > 0 Then DetectDistType = dtLognormal
    If InStr(formula, "MCEXPON") > 0 Then DetectDistType = dtExponential
    If InStr(formula, "MCPERT") > 0 Then DetectDistType = dtPERT
    If InStr(formula, "MCWEIBULL") > 0 Then DetectDistType = dtWeibull
    If InStr(formula, "MCGAMMA") > 0 Then DetectDistType = dtGamma
    If InStr(formula, "MCBETA") > 0 Then DetectDistType = dtBeta
    If InStr(formula, "MCBERNOULLI") > 0 Then DetectDistType = dtBernoulli
    If InStr(formula, "MCBINOMIAL") > 0 Then DetectDistType = dtBinomial
    If InStr(formula, "MCPOISSON") > 0 Then DetectDistType = dtPoisson
End Function

Private Sub RegisterOutput(ByVal cell As Range)
    Dim i As Long
    For i = 0 To g_OutputCount - 1
        If g_Outputs(i).SheetName = cell.Worksheet.Name And _
           g_Outputs(i).CellAddress = cell.Address Then Exit Sub
    Next i

    If g_OutputCount = 0 Then
        ReDim g_Outputs(0 To 0)
    Else
        ReDim Preserve g_Outputs(0 To g_OutputCount)
    End If

    With g_Outputs(g_OutputCount)
        .CellAddress = cell.Address
        .SheetName = cell.Worksheet.Name
        .OutputName = cell.Address
    End With
    g_OutputCount = g_OutputCount + 1
End Sub

'===============================================================================
' Save/Restore original cell values
'===============================================================================
Private Sub SaveOriginalValues()
    Dim i As Long
    For i = 0 To g_InputCount - 1
        On Error Resume Next
        g_Inputs(i).OriginalValue = ActiveWorkbook.Worksheets(g_Inputs(i).SheetName).Range(g_Inputs(i).CellAddress).Value
        On Error GoTo 0
    Next i
End Sub

Private Sub RestoreOriginalValues()
    ' Just recalculate - the functions return static values when not simulating
    On Error Resume Next
    Application.Calculate
    On Error GoTo 0
End Sub

'===============================================================================
' Latin Hypercube Sampling
'===============================================================================
Private Sub GenerateLHSPermutations()
    If g_InputCount = 0 Then Exit Sub

    Dim n As Long, nInputs As Long
    n = g_Settings.Iterations
    nInputs = g_InputCount

    ReDim g_LHSPermutations(0 To nInputs - 1, 1 To n)

    Dim i As Long, j As Long
    ' Create stratified permutations for each input
    For i = 0 To nInputs - 1
        ' Fill with sequential indices
        For j = 1 To n
            g_LHSPermutations(i, j) = j
        Next j
        ' Fisher-Yates shuffle
        Dim k As Long, temp As Long
        For j = n To 2 Step -1
            k = Int(MT_Random() * j) + 1
            temp = g_LHSPermutations(i, j)
            g_LHSPermutations(i, j) = g_LHSPermutations(i, k)
            g_LHSPermutations(i, k) = temp
        Next j
    Next i

    ' Apply correlations if defined
    If g_CorrelationCount > 0 Then
        ApplyCorrelationsToLHS
    End If
End Sub

Private Sub ApplyLHSSample(ByVal iteration As Long)
    ' LHS provides stratified U(0,1) values that the distribution functions
    ' will use. We modify the MT state to produce stratified values.
    ' This is handled internally by the distribution functions detecting
    ' the sampling method.
    ' The actual stratification happens in the distribution functions
    ' when they call MT_Random() - we intercept via the LHS permutation.
End Sub

'===============================================================================
' Correlation Application (Iman-Conover method)
'===============================================================================
Private Sub ApplyCorrelationsToLHS()
    ' Iman-Conover method to impose rank correlations
    ' This modifies the LHS permutations to match desired correlation structure
    If g_CorrelationCount = 0 Then Exit Sub

    ' Build target correlation matrix
    Dim n As Long, p As Long
    n = g_Settings.Iterations
    p = g_InputCount

    ' For simplicity, apply pairwise correlations
    Dim c As Long
    For c = 0 To g_CorrelationCount - 1
        Dim idx1 As Long, idx2 As Long, rho As Double
        idx1 = g_Correlations(c).Input1Index
        idx2 = g_Correlations(c).Input2Index
        rho = g_Correlations(c).Coefficient

        If idx1 >= 0 And idx1 < p And idx2 >= 0 And idx2 < p Then
            ' Sort permutation 2 to achieve desired rank correlation with permutation 1
            ' Using the van der Waerden scores approach
            AdjustPermutationCorrelation idx1, idx2, rho, n
        End If
    Next c
End Sub

Private Sub AdjustPermutationCorrelation(ByVal idx1 As Long, ByVal idx2 As Long, _
    ByVal targetRho As Double, ByVal n As Long)
    ' Simple correlation induction by sorting
    ' Generate correlated normal scores and use them to reorder

    Dim scores1() As Double, scores2() As Double
    Dim i As Long
    ReDim scores1(1 To n)
    ReDim scores2(1 To n)

    ' Generate correlated normal scores
    For i = 1 To n
        Dim z1 As Double, z2 As Double
        z1 = RandNormal01()
        z2 = targetRho * z1 + Sqr(1# - targetRho * targetRho) * RandNormal01()
        scores1(i) = z1
        scores2(i) = z2
    Next i

    ' Sort perm2 by scores2 rank ordering matched to scores1 rank ordering of perm1
    ' This induces approximately the target correlation
    Dim order() As Long
    ReDim order(1 To n)
    For i = 1 To n
        order(i) = i
    Next i
    QuickSortByKey scores2, order, 1, n

    ' Reorder LHS permutation for input idx2
    Dim temp() As Long
    ReDim temp(1 To n)
    For i = 1 To n
        temp(i) = g_LHSPermutations(idx2, i)
    Next i

    ' Sort temp by the same order as perm1 rank
    Dim order1() As Long
    ReDim order1(1 To n)
    For i = 1 To n
        order1(i) = i
    Next i
    Dim keys1() As Double
    ReDim keys1(1 To n)
    For i = 1 To n
        keys1(i) = CDbl(g_LHSPermutations(idx1, i))
    Next i
    QuickSortByKey keys1, order1, 1, n

    For i = 1 To n
        g_LHSPermutations(idx2, order1(i)) = temp(order(i))
    Next i
End Sub

'===============================================================================
' Convergence Check
'===============================================================================
Private Function CheckConvergence(ByVal currentIter As Long) As Boolean
    CheckConvergence = True
    Dim i As Long
    Dim halfIter As Long
    halfIter = currentIter \ 2

    For i = 0 To g_OutputCount - 1
        Dim mean1 As Double, mean2 As Double
        Dim sum1 As Double, sum2 As Double
        Dim j As Long

        sum1 = 0#: sum2 = 0#
        For j = 1 To halfIter
            sum1 = sum1 + g_Outputs(i).IterationData(j)
        Next j
        For j = halfIter + 1 To currentIter
            sum2 = sum2 + g_Outputs(i).IterationData(j)
        Next j
        mean1 = sum1 / halfIter
        mean2 = sum2 / (currentIter - halfIter)

        ' Check relative change
        Dim refVal As Double
        refVal = Abs(mean1)
        If refVal < 0.001 Then refVal = 0.001
        If Abs(mean2 - mean1) / refVal > g_Settings.ConvergenceTolerance Then
            CheckConvergence = False
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' Compute Statistics for an Output
'===============================================================================
Public Sub ComputeOutputStatistics(ByVal outputIdx As Long)
    Dim n As Long
    n = UBound(g_Outputs(outputIdx).IterationData)

    Dim data() As Double
    data = g_Outputs(outputIdx).IterationData

    ' Sort data for percentiles
    Dim sorted() As Double
    ReDim sorted(1 To n)
    Dim i As Long
    For i = 1 To n
        sorted(i) = data(i)
    Next i
    QuickSortDouble sorted, 1, n

    ' Basic statistics
    Dim sum As Double, sumSq As Double, sumCube As Double, sumQuad As Double
    Dim minV As Double, maxV As Double
    sum = 0#: sumSq = 0#: sumCube = 0#: sumQuad = 0#
    minV = data(1): maxV = data(1)

    For i = 1 To n
        sum = sum + data(i)
        If data(i) < minV Then minV = data(i)
        If data(i) > maxV Then maxV = data(i)
    Next i

    Dim mean As Double
    mean = sum / CDbl(n)

    For i = 1 To n
        Dim dev As Double
        dev = data(i) - mean
        sumSq = sumSq + dev * dev
        sumCube = sumCube + dev * dev * dev
        sumQuad = sumQuad + dev * dev * dev * dev
    Next i

    Dim variance As Double, stdDev As Double
    variance = sumSq / CDbl(n - 1)
    stdDev = Sqr(variance)

    With g_Outputs(outputIdx)
        .StatMean = mean
        .StatStdDev = stdDev
        .StatVariance = variance
        .StatMin = minV
        .StatMax = maxV
        .StatMedian = sorted(n \ 2 + 1)

        ' Skewness
        If stdDev > 0 Then
            .StatSkewness = (CDbl(n) / ((CDbl(n) - 1#) * (CDbl(n) - 2#))) * (sumCube / (stdDev ^ 3))
        Else
            .StatSkewness = 0#
        End If

        ' Kurtosis (excess)
        If stdDev > 0 And n > 3 Then
            .StatKurtosis = ((CDbl(n) * (CDbl(n) + 1#)) / ((CDbl(n) - 1#) * (CDbl(n) - 2#) * (CDbl(n) - 3#))) * _
                           (sumQuad / (variance * variance)) - _
                           (3# * (CDbl(n) - 1#) ^ 2) / ((CDbl(n) - 2#) * (CDbl(n) - 3#))
        Else
            .StatKurtosis = 0#
        End If

        ' Percentiles
        Dim p As Long
        For p = 0 To 100
            Dim pos As Double
            pos = (CDbl(p) / 100#) * CDbl(n - 1) + 1#
            Dim lo As Long, hi As Long
            lo = Int(pos)
            hi = lo + 1
            If lo < 1 Then lo = 1
            If hi > n Then hi = n
            .Percentiles(p) = sorted(lo) + (pos - lo) * (sorted(hi) - sorted(lo))
        Next p

        ' Mode (histogram-based estimation)
        .StatMode = EstimateMode(sorted, n)
    End With
End Sub

'===============================================================================
' Mode estimation using histogram binning
'===============================================================================
Private Function EstimateMode(ByRef sorted() As Double, ByVal n As Long) As Double
    Dim nBins As Long
    nBins = CLng(Sqr(CDbl(n)))
    If nBins < 10 Then nBins = 10
    If nBins > 200 Then nBins = 200

    Dim minV As Double, maxV As Double, binWidth As Double
    minV = sorted(1)
    maxV = sorted(n)
    If maxV = minV Then EstimateMode = minV: Exit Function
    binWidth = (maxV - minV) / CDbl(nBins)

    Dim bins() As Long
    ReDim bins(0 To nBins - 1)
    Dim i As Long
    For i = 1 To n
        Dim binIdx As Long
        binIdx = CLng(Int((sorted(i) - minV) / binWidth))
        If binIdx >= nBins Then binIdx = nBins - 1
        If binIdx < 0 Then binIdx = 0
        bins(binIdx) = bins(binIdx) + 1
    Next i

    Dim maxBin As Long, maxCount As Long
    maxBin = 0: maxCount = 0
    For i = 0 To nBins - 1
        If bins(i) > maxCount Then
            maxCount = bins(i)
            maxBin = i
        End If
    Next i

    EstimateMode = minV + (CDbl(maxBin) + 0.5) * binWidth
End Function

'===============================================================================
' QuickSort helpers
'===============================================================================
Public Sub QuickSortDouble(ByRef arr() As Double, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long, pivot As Double, temp As Double
    i = lo: j = hi
    pivot = arr((lo + hi) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortDouble arr, lo, j
    If i < hi Then QuickSortDouble arr, i, hi
End Sub

Public Sub QuickSortByKey(ByRef keys() As Double, ByRef indices() As Long, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long, pivotKey As Double, tempKey As Double, tempIdx As Long
    i = lo: j = hi
    pivotKey = keys((lo + hi) \ 2)
    Do While i <= j
        Do While keys(i) < pivotKey: i = i + 1: Loop
        Do While keys(j) > pivotKey: j = j - 1: Loop
        If i <= j Then
            tempKey = keys(i): keys(i) = keys(j): keys(j) = tempKey
            tempIdx = indices(i): indices(i) = indices(j): indices(j) = tempIdx
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortByKey keys, indices, lo, j
    If i < hi Then QuickSortByKey keys, indices, i, hi
End Sub

'===============================================================================
' Pause / Resume / Stop
'===============================================================================
Public Sub PauseSimulation()
    If g_SimState = ssRunning Then
        g_SimState = ssPaused
    End If
End Sub

Public Sub ResumeSimulation()
    If g_SimState = ssPaused Then
        g_SimState = ssRunning
    End If
End Sub

Public Sub StopSimulation()
    m_CancelRequested = True
    g_SimState = ssStopped
End Sub
