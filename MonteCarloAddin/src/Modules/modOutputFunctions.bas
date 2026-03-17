Attribute VB_Name = "modOutputFunctions"
'===============================================================================
' Module: modOutputFunctions
' Purpose: @RISK-compatible output and statistic worksheet functions
'===============================================================================
Option Explicit

'===============================================================================
' OUTPUT DESIGNATION
'===============================================================================

'--- Mark cell as simulation output ---
Public Function MCOutput(Optional ByVal outputName As String = "") As Variant
Attribute MCOutput.VB_Description = "Marks cell as simulation output for tracking."
    On Error Resume Next
    Dim caller As Range
    Set caller = Application.Caller
    If caller Is Nothing Then MCOutput = "": Exit Function

    If g_SimState = ssRunning Then
        ' Register this cell as an output
        Dim addr As String
        addr = caller.Worksheet.Name & "!" & caller.Address
        Dim i As Long
        Dim found As Boolean
        found = False
        For i = 0 To g_OutputCount - 1
            If g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress = addr Then
                found = True
                Exit For
            End If
        Next i

        If Not found Then
            If g_OutputCount = 0 Then
                ReDim g_Outputs(0 To 0)
            Else
                ReDim Preserve g_Outputs(0 To g_OutputCount)
            End If
            With g_Outputs(g_OutputCount)
                .CellAddress = caller.Address
                .SheetName = caller.Worksheet.Name
                If outputName = "" Then
                    .OutputName = caller.Address
                Else
                    .OutputName = outputName
                End If
                ReDim .IterationData(1 To g_Settings.Iterations)
            End With
            g_OutputCount = g_OutputCount + 1
        End If

        ' Return the cell value (transparent pass-through)
        MCOutput = ""
    Else
        MCOutput = ""
    End If
    On Error GoTo 0
End Function

'--- Add output with cell reference (alternative syntax) ---
Public Function MCAddOutput(ByVal cellRef As Range, Optional ByVal outputName As String = "") As String
Attribute MCAddOutput.VB_Description = "Registers a cell as an output by reference."
    If outputName = "" Then outputName = cellRef.Address

    Dim found As Boolean, i As Long
    found = False
    For i = 0 To g_OutputCount - 1
        If g_Outputs(i).SheetName = cellRef.Worksheet.Name And _
           g_Outputs(i).CellAddress = cellRef.Address Then
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        If g_OutputCount = 0 Then
            ReDim g_Outputs(0 To 0)
        Else
            ReDim Preserve g_Outputs(0 To g_OutputCount)
        End If
        With g_Outputs(g_OutputCount)
            .CellAddress = cellRef.Address
            .SheetName = cellRef.Worksheet.Name
            .OutputName = outputName
        End With
        g_OutputCount = g_OutputCount + 1
    End If
    MCAddOutput = outputName
End Function

'===============================================================================
' STATISTIC FUNCTIONS - Return simulation results in cells
'===============================================================================

'--- Mean ---
Public Function MCMean(ByVal cellRef As Range) As Variant
Attribute MCMean.VB_Description = "Returns the mean of simulation results for the referenced cell."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCMean = g_Outputs(idx).StatMean
    Else
        MCMean = CVErr(xlErrNA)
    End If
End Function

'--- Standard Deviation ---
Public Function MCStdDev(ByVal cellRef As Range) As Variant
Attribute MCStdDev.VB_Description = "Returns the std deviation of simulation results."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCStdDev = g_Outputs(idx).StatStdDev
    Else
        MCStdDev = CVErr(xlErrNA)
    End If
End Function

'--- Minimum ---
Public Function MCMin(ByVal cellRef As Range) As Variant
Attribute MCMin.VB_Description = "Returns the minimum simulation value."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCMin = g_Outputs(idx).StatMin
    Else
        MCMin = CVErr(xlErrNA)
    End If
End Function

'--- Maximum ---
Public Function MCMax(ByVal cellRef As Range) As Variant
Attribute MCMax.VB_Description = "Returns the maximum simulation value."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCMax = g_Outputs(idx).StatMax
    Else
        MCMax = CVErr(xlErrNA)
    End If
End Function

'--- Variance ---
Public Function MCVariance(ByVal cellRef As Range) As Variant
Attribute MCVariance.VB_Description = "Returns the variance of simulation results."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCVariance = g_Outputs(idx).StatVariance
    Else
        MCVariance = CVErr(xlErrNA)
    End If
End Function

'--- Skewness ---
Public Function MCSkewness(ByVal cellRef As Range) As Variant
Attribute MCSkewness.VB_Description = "Returns skewness of simulation results."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCSkewness = g_Outputs(idx).StatSkewness
    Else
        MCSkewness = CVErr(xlErrNA)
    End If
End Function

'--- Kurtosis ---
Public Function MCKurtosis(ByVal cellRef As Range) As Variant
Attribute MCKurtosis.VB_Description = "Returns kurtosis of simulation results."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCKurtosis = g_Outputs(idx).StatKurtosis
    Else
        MCKurtosis = CVErr(xlErrNA)
    End If
End Function

'--- Percentile ---
Public Function MCPercentile(ByVal cellRef As Range, ByVal pct As Double) As Variant
Attribute MCPercentile.VB_Description = "Returns the specified percentile of simulation results."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        Dim pctIdx As Long
        pctIdx = CLng(pct * 100#)
        If pctIdx < 0 Then pctIdx = 0
        If pctIdx > 100 Then pctIdx = 100
        MCPercentile = g_Outputs(idx).Percentiles(pctIdx)
    Else
        MCPercentile = CVErr(xlErrNA)
    End If
End Function

'--- Percentile Descending ---
Public Function MCPercentileD(ByVal cellRef As Range, ByVal pct As Double) As Variant
Attribute MCPercentileD.VB_Description = "Returns the descending percentile of simulation results."
    MCPercentileD = MCPercentile(cellRef, 1# - pct)
End Function

'--- Target (Cumulative probability) ---
Public Function MCTarget(ByVal cellRef As Range, ByVal targetValue As Double) As Variant
Attribute MCTarget.VB_Description = "Returns probability that output <= target value."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        ' Count values <= target
        Dim count As Long, i As Long, n As Long
        n = UBound(g_Outputs(idx).IterationData)
        count = 0
        For i = 1 To n
            If g_Outputs(idx).IterationData(i) <= targetValue Then count = count + 1
        Next i
        MCTarget = CDbl(count) / CDbl(n)
    Else
        MCTarget = CVErr(xlErrNA)
    End If
End Function

'--- Target Descending ---
Public Function MCTargetD(ByVal cellRef As Range, ByVal targetValue As Double) As Variant
Attribute MCTargetD.VB_Description = "Returns probability that output > target value."
    Dim t As Variant
    t = MCTarget(cellRef, targetValue)
    If IsError(t) Then
        MCTargetD = t
    Else
        MCTargetD = 1# - CDbl(t)
    End If
End Function

'--- Median ---
Public Function MCMedian(ByVal cellRef As Range) As Variant
Attribute MCMedian.VB_Description = "Returns the median of simulation results."
    MCMedian = MCPercentile(cellRef, 0.5)
End Function

'--- Mode ---
Public Function MCMode(ByVal cellRef As Range) As Variant
Attribute MCMode.VB_Description = "Returns the mode of simulation results."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        MCMode = g_Outputs(idx).StatMode
    Else
        MCMode = CVErr(xlErrNA)
    End If
End Function

'--- Data (specific iteration value) ---
Public Function MCData(ByVal cellRef As Range, ByVal iteration As Long, Optional ByVal simNum As Long = 1) As Variant
Attribute MCData.VB_Description = "Returns the value from a specific iteration."
    Dim idx As Long
    idx = FindOutputIndex(cellRef)
    If idx >= 0 And g_SimState = ssComplete Then
        If iteration >= 1 And iteration <= UBound(g_Outputs(idx).IterationData) Then
            MCData = g_Outputs(idx).IterationData(iteration)
        Else
            MCData = CVErr(xlErrNA)
        End If
    Else
        MCData = CVErr(xlErrNA)
    End If
End Function

'===============================================================================
' HELPER: Find output index by cell reference
'===============================================================================
Private Function FindOutputIndex(ByVal cellRef As Range) As Long
    Dim i As Long
    For i = 0 To g_OutputCount - 1
        If g_Outputs(i).SheetName = cellRef.Worksheet.Name And _
           g_Outputs(i).CellAddress = cellRef.Address Then
            FindOutputIndex = i
            Exit Function
        End If
    Next i
    FindOutputIndex = -1
End Function
