VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorrelation
   Caption         =   "Define Correlations"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   OleObjectBlob   =   "frmCorrelation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCorrelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Form: frmCorrelation
' Purpose: Define rank correlations between input distributions
'===============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    Me.Width = 500
    Me.Height = 400

    ' Scan for inputs first
    ScanWorkbookForFunctions

    ' Instructions
    Dim lblInstr As MSForms.Label
    Set lblInstr = Me.Controls.Add("Forms.Label.1", "lblInstructions")
    With lblInstr
        .Caption = "Define Spearman rank correlations between input distributions. " & _
                  "Values range from -1 (perfect negative) to +1 (perfect positive)."
        .Left = 12: .Top = 8: .Width = 460: .Height = 30
        .WordWrap = True
    End With

    ' Input 1
    Dim lblInput1 As MSForms.Label
    Set lblInput1 = Me.Controls.Add("Forms.Label.1", "lblInput1")
    With lblInput1
        .Caption = "Input 1:"
        .Left = 12: .Top = 50: .Width = 60: .Height = 15
        .Font.Bold = True
    End With

    Dim cboInput1 As MSForms.ComboBox
    Set cboInput1 = Me.Controls.Add("Forms.ComboBox.1", "cboInput1")
    With cboInput1
        .Left = 80: .Top = 48: .Width = 200: .Height = 20
        .Style = fmStyleDropDownList
    End With

    ' Input 2
    Dim lblInput2 As MSForms.Label
    Set lblInput2 = Me.Controls.Add("Forms.Label.1", "lblInput2")
    With lblInput2
        .Caption = "Input 2:"
        .Left = 12: .Top = 80: .Width = 60: .Height = 15
        .Font.Bold = True
    End With

    Dim cboInput2 As MSForms.ComboBox
    Set cboInput2 = Me.Controls.Add("Forms.ComboBox.1", "cboInput2")
    With cboInput2
        .Left = 80: .Top = 78: .Width = 200: .Height = 20
        .Style = fmStyleDropDownList
    End With

    ' Populate input combos
    Dim i As Long
    For i = 0 To g_InputCount - 1
        Dim entry As String
        entry = g_Inputs(i).InputName & " (" & g_Inputs(i).SheetName & "!" & g_Inputs(i).CellAddress & ")"
        cboInput1.AddItem entry
        cboInput2.AddItem entry
    Next i

    ' Correlation coefficient
    Dim lblCoeff As MSForms.Label
    Set lblCoeff = Me.Controls.Add("Forms.Label.1", "lblCoefficient")
    With lblCoeff
        .Caption = "Correlation:"
        .Left = 12: .Top = 112: .Width = 70: .Height = 15
        .Font.Bold = True
    End With

    Dim txtCoeff As MSForms.TextBox
    Set txtCoeff = Me.Controls.Add("Forms.TextBox.1", "txtCoefficient")
    With txtCoeff
        .Left = 80: .Top = 110: .Width = 60: .Height = 20
        .Text = "0.5"
    End With

    Dim lblRange As MSForms.Label
    Set lblRange = Me.Controls.Add("Forms.Label.1", "lblRange")
    With lblRange
        .Caption = "(-1.0 to +1.0)"
        .Left = 150: .Top = 112: .Width = 100: .Height = 15
        .ForeColor = RGB(100, 100, 100)
    End With

    ' Add button
    Dim btnAdd As MSForms.CommandButton
    Set btnAdd = Me.Controls.Add("Forms.CommandButton.1", "btnAdd")
    With btnAdd
        .Caption = "Add Correlation"
        .Left = 300: .Top = 78: .Width = 120: .Height = 28
    End With

    ' Existing correlations list
    Dim lblExisting As MSForms.Label
    Set lblExisting = Me.Controls.Add("Forms.Label.1", "lblExisting")
    With lblExisting
        .Caption = "Defined Correlations:"
        .Left = 12: .Top = 148: .Width = 150: .Height = 15
        .Font.Bold = True
    End With

    Dim lstCorr As MSForms.ListBox
    Set lstCorr = Me.Controls.Add("Forms.ListBox.1", "lstCorrelations")
    With lstCorr
        .Left = 12: .Top = 166: .Width = 460: .Height = 140
        .ColumnCount = 3
        .ColumnWidths = "180;180;60"
    End With

    ' Load existing correlations
    For i = 0 To g_CorrelationCount - 1
        lstCorr.AddItem
        lstCorr.List(lstCorr.ListCount - 1, 0) = g_Inputs(g_Correlations(i).Input1Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 1) = g_Inputs(g_Correlations(i).Input2Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 2) = Format(g_Correlations(i).Coefficient, "0.00")
    Next i

    ' Remove button
    Dim btnRemove As MSForms.CommandButton
    Set btnRemove = Me.Controls.Add("Forms.CommandButton.1", "btnRemove")
    With btnRemove
        .Caption = "Remove Selected"
        .Left = 340: .Top = 316: .Width = 120: .Height = 25
    End With

    ' Close button
    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = 380: .Top = 350: .Width = 80: .Height = 28
        .Cancel = True
    End With
End Sub

Private Sub btnAdd_Click()
    Dim cbo1 As MSForms.ComboBox, cbo2 As MSForms.ComboBox
    Set cbo1 = Me.Controls("cboInput1")
    Set cbo2 = Me.Controls("cboInput2")

    If cbo1.ListIndex < 0 Or cbo2.ListIndex < 0 Then
        MsgBox "Select both inputs.", vbExclamation
        Exit Sub
    End If

    If cbo1.ListIndex = cbo2.ListIndex Then
        MsgBox "Select two different inputs.", vbExclamation
        Exit Sub
    End If

    Dim coeff As Double
    On Error GoTo InvalidCoeff
    coeff = CDbl(Me.Controls("txtCoefficient").Text)
    On Error GoTo 0

    If coeff < -1# Or coeff > 1# Then
        MsgBox "Correlation must be between -1 and +1.", vbExclamation
        Exit Sub
    End If

    ' Add correlation
    If g_CorrelationCount = 0 Then
        ReDim g_Correlations(0 To 0)
    Else
        ReDim Preserve g_Correlations(0 To g_CorrelationCount)
    End If
    g_Correlations(g_CorrelationCount).Input1Index = cbo1.ListIndex
    g_Correlations(g_CorrelationCount).Input2Index = cbo2.ListIndex
    g_Correlations(g_CorrelationCount).Coefficient = coeff
    g_CorrelationCount = g_CorrelationCount + 1

    ' Update list
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstCorrelations")
    lst.AddItem
    lst.List(lst.ListCount - 1, 0) = g_Inputs(cbo1.ListIndex).InputName
    lst.List(lst.ListCount - 1, 1) = g_Inputs(cbo2.ListIndex).InputName
    lst.List(lst.ListCount - 1, 2) = Format(coeff, "0.00")
    Exit Sub

InvalidCoeff:
    MsgBox "Enter a valid number.", vbExclamation
End Sub

Private Sub btnRemove_Click()
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstCorrelations")
    If lst.ListIndex < 0 Then Exit Sub

    Dim idx As Long
    idx = lst.ListIndex

    ' Remove from array
    Dim i As Long
    For i = idx To g_CorrelationCount - 2
        g_Correlations(i) = g_Correlations(i + 1)
    Next i
    g_CorrelationCount = g_CorrelationCount - 1

    lst.RemoveItem idx
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub
