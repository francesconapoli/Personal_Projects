VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSensitivity
   Caption         =   "Sensitivity Analysis"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   OleObjectBlob   =   "frmSensitivity.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSensitivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Form: frmSensitivity
' Purpose: Sensitivity analysis configuration and results
'===============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    Me.Width = 520
    Me.Height = 450

    ' Output selector
    Dim lblOutput As MSForms.Label
    Set lblOutput = Me.Controls.Add("Forms.Label.1", "lblOutput")
    With lblOutput
        .Caption = "Analyze Output:"
        .Left = 12: .Top = 12: .Width = 90: .Height = 15
        .Font.Bold = True
    End With

    Dim cboOutput As MSForms.ComboBox
    Set cboOutput = Me.Controls.Add("Forms.ComboBox.1", "cboOutput")
    With cboOutput
        .Left = 110: .Top = 10: .Width = 250: .Height = 20
        .Style = fmStyleDropDownList
    End With

    Dim i As Long
    For i = 0 To g_OutputCount - 1
        cboOutput.AddItem g_Outputs(i).OutputName
    Next i
    If g_OutputCount > 0 Then cboOutput.ListIndex = 0

    ' Analysis method
    Dim fraMethod As MSForms.Frame
    Set fraMethod = Me.Controls.Add("Forms.Frame.1", "fraMethod")
    With fraMethod
        .Caption = "Sensitivity Method"
        .Left = 12: .Top = 40: .Width = 480: .Height = 90
    End With

    Dim methods As Variant
    methods = Array("Regression Coefficients", "Rank Correlation (Spearman)", _
                   "Contribution to Variance", "Change in Output Statistic")
    For i = 0 To 3
        Dim opt As MSForms.OptionButton
        Set opt = fraMethod.Controls.Add("Forms.OptionButton.1", "optMethod" & i)
        With opt
            .Caption = CStr(methods(i))
            .Left = 12 + (i Mod 2) * 230
            .Top = 18 + (i \ 2) * 28
            .Width = 220: .Height = 18
            If i = 1 Then .Value = True
        End With
    Next i

    ' Run button
    Dim btnRun As MSForms.CommandButton
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Run Analysis"
        .Left = 370: .Top = 140: .Width = 120: .Height = 28
    End With

    ' Results
    Dim lblResults As MSForms.Label
    Set lblResults = Me.Controls.Add("Forms.Label.1", "lblResults")
    With lblResults
        .Caption = "Results:"
        .Left = 12: .Top = 145: .Width = 60: .Height = 15
        .Font.Bold = True
    End With

    Dim lblR2 As MSForms.Label
    Set lblR2 = Me.Controls.Add("Forms.Label.1", "lblRSquared")
    With lblR2
        .Caption = ""
        .Left = 80: .Top = 145: .Width = 200: .Height = 15
    End With

    Dim lstResults As MSForms.ListBox
    Set lstResults = Me.Controls.Add("Forms.ListBox.1", "lstResults")
    With lstResults
        .Left = 12: .Top = 170: .Width = 480: .Height = 160
        .ColumnCount = 4
        .ColumnWidths = "40;150;100;100"
    End With

    ' Chart buttons
    Dim btnTornado As MSForms.CommandButton
    Set btnTornado = Me.Controls.Add("Forms.CommandButton.1", "btnTornado")
    With btnTornado
        .Caption = "Tornado Chart"
        .Left = 12: .Top = 340: .Width = 110: .Height = 28
    End With

    Dim btnSpider As MSForms.CommandButton
    Set btnSpider = Me.Controls.Add("Forms.CommandButton.1", "btnSpider")
    With btnSpider
        .Caption = "Spider Plot"
        .Left = 130: .Top = 340: .Width = 110: .Height = 28
    End With

    Dim btnExport As MSForms.CommandButton
    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExport")
    With btnExport
        .Caption = "Export to Sheet"
        .Left = 248: .Top = 340: .Width = 110: .Height = 28
    End With

    ' Close
    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = 410: .Top = 380: .Width = 80: .Height = 28
        .Cancel = True
    End With
End Sub

Private Sub btnRun_Click()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboOutput")
    If cbo.ListIndex < 0 Then
        MsgBox "Select an output.", vbExclamation
        Exit Sub
    End If

    ' Determine method
    Dim method As SensitivityMethod
    Dim fra As MSForms.Frame
    Set fra = Me.Controls("fraMethod")
    If fra.Controls("optMethod0").Value Then method = snRegression
    If fra.Controls("optMethod1").Value Then method = snRankCorrelation
    If fra.Controls("optMethod2").Value Then method = snContributionToVariance
    If fra.Controls("optMethod3").Value Then method = snChangeInStatistic

    ' Run analysis
    RunSensitivityAnalysis cbo.ListIndex, method

    ' Display results
    Me.Controls("lblRSquared").Caption = "R-Squared = " & Format(g_SensRSquared, "0.0000")

    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstResults")
    lst.Clear

    ' Header
    lst.AddItem
    lst.List(0, 0) = "Rank"
    lst.List(0, 1) = "Input"
    lst.List(0, 2) = "Coefficient"
    lst.List(0, 3) = "Contribution %"

    Dim i As Long
    For i = 0 To g_SensResultCount - 1
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = CStr(i + 1)
        lst.List(lst.ListCount - 1, 1) = g_SensResults(i).InputName
        lst.List(lst.ListCount - 1, 2) = Format(g_SensResults(i).Coefficient, "0.0000")
        lst.List(lst.ListCount - 1, 3) = Format(g_SensResults(i).ContributionPct, "0.00") & "%"
    Next i
End Sub

Private Sub btnTornado_Click()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboOutput")
    If cbo.ListIndex >= 0 Then CreateTornadoDiagram cbo.ListIndex
End Sub

Private Sub btnSpider_Click()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboOutput")
    If cbo.ListIndex >= 0 Then CreateSpiderPlot cbo.ListIndex
End Sub

Private Sub btnExport_Click()
    If g_SensResultCount = 0 Then
        MsgBox "Run analysis first.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateResultsSheet()
    Dim startRow As Long
    startRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 3

    ws.Cells(startRow, 1).Value = "Sensitivity Analysis Results"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 12
    ws.Cells(startRow + 1, 1).Value = "R-Squared:"
    ws.Cells(startRow + 1, 2).Value = g_SensRSquared

    ws.Cells(startRow + 3, 1).Value = "Rank"
    ws.Cells(startRow + 3, 2).Value = "Input"
    ws.Cells(startRow + 3, 3).Value = "Coefficient"
    ws.Cells(startRow + 3, 4).Value = "Contribution %"
    ws.Range(ws.Cells(startRow + 3, 1), ws.Cells(startRow + 3, 4)).Font.Bold = True

    Dim i As Long
    For i = 0 To g_SensResultCount - 1
        ws.Cells(startRow + 4 + i, 1).Value = i + 1
        ws.Cells(startRow + 4 + i, 2).Value = g_SensResults(i).InputName
        ws.Cells(startRow + 4 + i, 3).Value = g_SensResults(i).Coefficient
        ws.Cells(startRow + 4 + i, 4).Value = g_SensResults(i).ContributionPct
    Next i

    ws.Columns.AutoFit
    MsgBox "Results exported to 'MC Results' sheet.", vbInformation
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub
