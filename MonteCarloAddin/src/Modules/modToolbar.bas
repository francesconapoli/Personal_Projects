Attribute VB_Name = "modToolbar"
'===============================================================================
' Module: modToolbar
' Purpose: Build a custom CommandBar toolbar programmatically on load.
'          This replaces the Ribbon XML approach and works in ALL
'          Excel versions without needing Office RibbonX Editor.
'===============================================================================
Option Explicit

Private Const TOOLBAR_NAME As String = "MC Simulation"

'===============================================================================
' Create the toolbar (called from Workbook_Open / AddinInstall)
'===============================================================================
Public Sub CreateToolbar()
    On Error Resume Next
    ' Remove existing toolbar to avoid duplicates
    Application.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0

    Dim cb As CommandBar
    Set cb = Application.CommandBars.Add(Name:=TOOLBAR_NAME, Position:=msoBarTop, Temporary:=True)
    cb.Visible = True

    ' ============================================================
    ' DEFINE GROUP
    ' ============================================================
    Dim mnuDefine As CommandBarPopup
    Set mnuDefine = cb.Controls.Add(Type:=msoControlPopup)
    mnuDefine.Caption = "Define"
    mnuDefine.BeginGroup = False

    With mnuDefine.Controls.Add(Type:=msoControlButton)
        .Caption = "Add &Input Distribution..."
        .OnAction = "ShowAddInput"
        .FaceId = 385   ' function icon
        .Style = msoButtonIconAndCaption
    End With

    With mnuDefine.Controls.Add(Type:=msoControlButton)
        .Caption = "Add &Output..."
        .OnAction = "ShowAddOutput"
        .FaceId = 388
        .Style = msoButtonIconAndCaption
    End With

    With mnuDefine.Controls.Add(Type:=msoControlButton)
        .Caption = "Define &Correlations..."
        .OnAction = "ShowCorrelations"
        .FaceId = 283
        .Style = msoButtonIconAndCaption
    End With

    With mnuDefine.Controls.Add(Type:=msoControlButton)
        .Caption = "&Model Window..."
        .OnAction = "ShowModelWindow"
        .FaceId = 547
        .Style = msoButtonIconAndCaption
    End With

    With mnuDefine.Controls.Add(Type:=msoControlButton)
        .Caption = "&Fit Distribution to Data..."
        .OnAction = "FitDistributionToRange"
        .FaceId = 480
        .Style = msoButtonIconAndCaption
    End With

    ' ============================================================
    ' SIMULATION BUTTONS (directly on toolbar)
    ' ============================================================
    Dim btnRun As CommandBarButton
    Set btnRun = cb.Controls.Add(Type:=msoControlButton)
    With btnRun
        .Caption = "Run Simulation"
        .OnAction = "RunSimulation"
        .FaceId = 186   ' green play
        .Style = msoButtonIconAndCaption
        .BeginGroup = True
    End With

    Dim btnPause As CommandBarButton
    Set btnPause = cb.Controls.Add(Type:=msoControlButton)
    With btnPause
        .Caption = "Pause"
        .OnAction = "PauseSimulation"
        .FaceId = 2181
        .Style = msoButtonIconAndCaption
    End With

    Dim btnStop As CommandBarButton
    Set btnStop = cb.Controls.Add(Type:=msoControlButton)
    With btnStop
        .Caption = "Stop"
        .OnAction = "StopSimulation"
        .FaceId = 2186
        .Style = msoButtonIconAndCaption
    End With

    ' Settings
    Dim btnSettings As CommandBarButton
    Set btnSettings = cb.Controls.Add(Type:=msoControlButton)
    With btnSettings
        .Caption = "Settings..."
        .OnAction = "ShowSettings"
        .FaceId = 2950
        .Style = msoButtonIconAndCaption
        .BeginGroup = True
    End With

    ' ============================================================
    ' RESULTS GROUP
    ' ============================================================
    Dim mnuResults As CommandBarPopup
    Set mnuResults = cb.Controls.Add(Type:=msoControlPopup)
    mnuResults.Caption = "Results"
    mnuResults.BeginGroup = True

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "&Browse Results..."
        .OnAction = "ShowResults"
        .FaceId = 2529
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "&All Charts..."
        .OnAction = "ShowAllCharts"
        .FaceId = 1677
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "&Histogram"
        .OnAction = "ShowHistogram"
        .FaceId = 1658
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "&Cumulative Chart"
        .OnAction = "ShowCumulative"
        .FaceId = 1660
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "&Overlay Chart"
        .OnAction = "ShowOverlay"
        .FaceId = 1661
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "&Tornado Diagram"
        .OnAction = "ShowTornado"
        .FaceId = 1662
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "&Spider Plot"
        .OnAction = "ShowSpider"
        .FaceId = 1665
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "S&ensitivity Analysis..."
        .OnAction = "ShowSensitivityDlg"
        .FaceId = 481
        .Style = msoButtonIconAndCaption
        .BeginGroup = True
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "Statistics &Report"
        .OnAction = "ShowStatsReport"
        .FaceId = 2530
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "E&xport Raw Data"
        .OnAction = "ExportRawData"
        .FaceId = 2174
        .Style = msoButtonIconAndCaption
    End With

    With mnuResults.Controls.Add(Type:=msoControlButton)
        .Caption = "Rese&t All"
        .OnAction = "ResetAll"
        .FaceId = 2948
        .Style = msoButtonIconAndCaption
        .BeginGroup = True
    End With
End Sub

'===============================================================================
' Remove toolbar (called from Workbook_BeforeClose / AddinUninstall)
'===============================================================================
Public Sub RemoveToolbar()
    On Error Resume Next
    Application.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0
End Sub

'===============================================================================
' Wrapper Subs for OnAction (OnAction needs Public Subs in standard modules)
'===============================================================================
Public Sub ShowAddInput()
    frmAddInput.Show vbModal
End Sub

Public Sub ShowAddOutput()
    If ActiveCell Is Nothing Then
        MsgBox "Select a cell first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    Dim outputName As String
    outputName = InputBox("Enter output name (or leave blank for cell address):", _
                         ADDIN_NAME, ActiveCell.Address)
    If Len(outputName) = 0 Then outputName = ActiveCell.Address

    ' Append MCOutput to existing formula
    If ActiveCell.HasFormula Then
        Dim formula As String
        formula = ActiveCell.formula
        If InStr(UCase(formula), "MCOUTPUT") = 0 Then
            ActiveCell.formula = formula & "+MCOutput(""" & outputName & """)"
        Else
            MsgBox "This cell already has MCOutput.", vbInformation, ADDIN_NAME
        End If
    Else
        Dim v As String
        v = CStr(ActiveCell.Value)
        If Len(v) = 0 Then v = "0"
        ActiveCell.formula = "=" & v & "+MCOutput(""" & outputName & """)"
    End If

    MsgBox "Output '" & outputName & "' registered at " & ActiveCell.Address, _
           vbInformation, ADDIN_NAME
End Sub

Public Sub ShowCorrelations()
    frmCorrelation.Show vbModal
End Sub

Public Sub ShowModelWindow()
    frmModelWindow.Show vbModeless
End Sub

Public Sub ShowSettings()
    frmSettings.Show vbModal
End Sub

Public Sub ShowResults()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If
    frmResults.Show vbModeless
End Sub

Public Sub ShowAllCharts()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If
    Dim idx As Long
    idx = PickOutput()
    If idx >= 0 Then GenerateAllCharts idx
End Sub

Public Sub ShowHistogram()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If
    Dim idx As Long: idx = PickOutput()
    If idx >= 0 Then CreateHistogram idx
End Sub

Public Sub ShowCumulative()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If
    Dim idx As Long: idx = PickOutput()
    If idx >= 0 Then CreateCumulativeChart idx
End Sub

Public Sub ShowOverlay()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If
    Dim idx As Long: idx = PickOutput()
    If idx >= 0 Then CreateOverlayChart idx
End Sub

Public Sub ShowTornado()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If
    Dim idx As Long: idx = PickOutput()
    If idx >= 0 Then CreateTornadoDiagram idx
End Sub

Public Sub ShowSpider()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If
    Dim idx As Long: idx = PickOutput()
    If idx >= 0 Then CreateSpiderPlot idx
End Sub

Public Sub ShowSensitivityDlg()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If
    frmSensitivity.Show vbModal
End Sub

Public Sub ShowStatsReport()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = GetOrCreateResultsSheet()
    ws.Cells.Clear
    CreateStatisticsTable ws, 1
    ws.Activate
    MsgBox "Statistics report generated on 'MC Results' sheet.", vbInformation, ADDIN_NAME
End Sub

Public Sub ExportRawData()
    If g_SimState <> ssComplete Then
        MsgBox "Run a simulation first.", vbExclamation, ADDIN_NAME: Exit Sub
    End If

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

    ws.Cells(1, 1).Value = "Iteration"
    Dim i As Long, j As Long
    For i = 0 To g_OutputCount - 1
        ws.Cells(1, i + 2).Value = g_Outputs(i).OutputName
    Next i

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
    MsgBox "Exported " & n & " iterations x " & g_OutputCount & " outputs.", vbInformation, ADDIN_NAME
End Sub

Public Sub ResetAll()
    If MsgBox("Reset all simulation data?", vbQuestion + vbYesNo, ADDIN_NAME) = vbYes Then
        ResetSimulation
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
' Helper: Pick output index (returns -1 if cancelled)
'===============================================================================
Private Function PickOutput() As Long
    If g_OutputCount = 0 Then
        MsgBox "No outputs defined.", vbExclamation, ADDIN_NAME
        PickOutput = -1
        Exit Function
    End If
    If g_OutputCount = 1 Then
        PickOutput = 0
        Exit Function
    End If

    frmOutputSelect.Show vbModal
    PickOutput = frmOutputSelect.SelectedIndex
End Function
