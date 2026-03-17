Attribute VB_Name = "modRibbonCallbacks"
'===============================================================================
' Module: modRibbonCallbacks
' Purpose: Ribbon callback procedures for the custom ribbon tab
'===============================================================================
Option Explicit

Private g_Ribbon As IRibbonUI

'===============================================================================
' Ribbon Load
'===============================================================================
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set g_Ribbon = ribbon
    InitializeDefaults
End Sub

Public Sub RefreshRibbon()
    If Not g_Ribbon Is Nothing Then
        g_Ribbon.Invalidate
    End If
End Sub

'===============================================================================
' DEFINE GROUP CALLBACKS
'===============================================================================

'--- Add Input Distribution ---
Public Sub OnAddInput(control As IRibbonControl)
    frmAddInput.Show vbModal
End Sub

'--- Add Output ---
Public Sub OnAddOutput(control As IRibbonControl)
    If ActiveCell Is Nothing Then
        MsgBox "Select a cell first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    Dim outputName As String
    outputName = InputBox("Enter output name (or leave blank for cell address):", _
                         ADDIN_NAME, ActiveCell.Address)
    If outputName = "" Then outputName = ActiveCell.Address

    ' If cell has a formula, wrap it with MCOutput
    If ActiveCell.HasFormula Then
        Dim formula As String
        formula = ActiveCell.formula
        ' Don't double-wrap
        If InStr(UCase(formula), "MCOUTPUT") = 0 Then
            ActiveCell.formula = "=" & formula & "+MCOutput(""" & outputName & """)"
        End If
    Else
        ' Add MCOutput formula that references this cell
        ActiveCell.formula = "=" & ActiveCell.Value & "+MCOutput(""" & outputName & """)"
    End If

    MsgBox "Output '" & outputName & "' added at " & ActiveCell.Address, vbInformation, ADDIN_NAME
End Sub

'--- Define Correlation ---
Public Sub OnDefineCorrelation(control As IRibbonControl)
    frmCorrelation.Show vbModal
End Sub

'--- View Model ---
Public Sub OnViewModel(control As IRibbonControl)
    frmModelWindow.Show vbModeless
End Sub

'===============================================================================
' SIMULATION GROUP CALLBACKS
'===============================================================================

'--- Run Simulation ---
Public Sub OnRunSimulation(control As IRibbonControl)
    RunSimulation
End Sub

'--- Pause Simulation ---
Public Sub OnPauseSimulation(control As IRibbonControl)
    If g_SimState = ssRunning Then
        PauseSimulation
        MsgBox "Simulation paused at iteration " & g_CurrentIteration, vbInformation, ADDIN_NAME
    ElseIf g_SimState = ssPaused Then
        ResumeSimulation
    End If
End Sub

'--- Stop Simulation ---
Public Sub OnStopSimulation(control As IRibbonControl)
    StopSimulation
End Sub

'--- Simulation Settings ---
Public Sub OnSimSettings(control As IRibbonControl)
    frmSettings.Show vbModal
End Sub

'===============================================================================
' RESULTS GROUP CALLBACKS
'===============================================================================

'--- Browse Results ---
Public Sub OnBrowseResults(control As IRibbonControl)
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If
    frmResults.Show vbModeless
End Sub

'--- Generate Charts ---
Public Sub OnGenerateCharts(control As IRibbonControl)
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    If g_OutputCount = 0 Then
        MsgBox "No outputs defined.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    ' Generate charts for first output (or selected)
    Dim outputIdx As Long
    If g_OutputCount = 1 Then
        outputIdx = 0
    Else
        ' Show output selector
        frmOutputSelect.Show vbModal
        outputIdx = frmOutputSelect.SelectedIndex
        If outputIdx < 0 Then Exit Sub
    End If

    GenerateAllCharts outputIdx
End Sub

'--- Sensitivity Analysis ---
Public Sub OnSensitivity(control As IRibbonControl)
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If
    frmSensitivity.Show vbModal
End Sub

'--- Statistics Report ---
Public Sub OnStatisticsReport(control As IRibbonControl)
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateResultsSheet()
    ws.Cells.Clear
    CreateStatisticsTable ws, 1
    ws.Activate
    MsgBox "Statistics report generated on 'MC Results' sheet.", vbInformation, ADDIN_NAME
End Sub

'--- Export Data ---
Public Sub OnExportData(control As IRibbonControl)
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    ' Export iteration data to a new sheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("MC Raw Data")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "MC Raw Data"
    Else
        ws.Cells.Clear
    End If

    ' Header row
    Dim i As Long, j As Long
    ws.Cells(1, 1).Value = "Iteration"
    For i = 0 To g_OutputCount - 1
        ws.Cells(1, i + 2).Value = g_Outputs(i).OutputName
    Next i

    ' Data rows
    Dim n As Long
    n = UBound(g_Outputs(0).IterationData)
    For j = 1 To n
        ws.Cells(j + 1, 1).Value = j
        For i = 0 To g_OutputCount - 1
            ws.Cells(j + 1, i + 2).Value = g_Outputs(i).IterationData(j)
        Next i
    Next j

    ws.Columns.AutoFit
    ws.Activate
    MsgBox "Raw data exported: " & n & " iterations x " & g_OutputCount & " outputs.", vbInformation, ADDIN_NAME
End Sub

'--- Reset ---
Public Sub OnReset(control As IRibbonControl)
    If MsgBox("Reset all simulation data?", vbQuestion + vbYesNo, ADDIN_NAME) = vbYes Then
        ResetSimulation
        ' Clean up results sheets
        On Error Resume Next
        Application.DisplayAlerts = False
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets("MC Results")
        If Not ws Is Nothing Then ws.Delete
        Set ws = ActiveWorkbook.Worksheets("MC Raw Data")
        If Not ws Is Nothing Then ws.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Application.StatusBar = False
        MsgBox "Simulation reset.", vbInformation, ADDIN_NAME
    End If
End Sub

'===============================================================================
' RIBBON STATE CALLBACKS (for enabling/disabling buttons)
'===============================================================================
Public Sub GetEnabled_SimRunning(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = (g_SimState = ssRunning Or g_SimState = ssPaused)
End Sub

Public Sub GetEnabled_HasResults(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = (g_SimState = ssComplete)
End Sub

Public Sub GetImage(control As IRibbonControl, ByRef returnedVal As Variant)
    Select Case control.ID
        Case "btnRunSim": returnedVal = "PlayMacro"
        Case "btnPauseSim": returnedVal = "AnimationPause"
        Case "btnStopSim": returnedVal = "RecordStop"
        Case "btnSettings": returnedVal = "ControlProperties"
        Case "btnAddInput": returnedVal = "InsertFunctionDialog"
        Case "btnAddOutput": returnedVal = "ConditionalFormattingColorScalesGallery"
        Case "btnCorrelation": returnedVal = "TableRelationships"
        Case "btnViewModel": returnedVal = "DataModel"
        Case "btnBrowseResults": returnedVal = "ViewResultsTable"
        Case "btnCharts": returnedVal = "ChartInsert"
        Case "btnSensitivity": returnedVal = "AnalyzeTrends"
        Case "btnStatistics": returnedVal = "InfoFunctionInsertGallery"
        Case "btnExportData": returnedVal = "ExportToDatabase"
        Case "btnReset": returnedVal = "ClearAll"
        Case Else: returnedVal = "CustomMacro"
    End Select
End Sub
