'===============================================================================
' Form: frmCorrelation
' Purpose: Define rank correlations between input distributions
'===============================================================================
'@FORM:500,400,Define Correlations
'@CTRL:Label,lblInstructions,,12,8,460,30,Define Spearman rank correlations between input distributions. Values range from -1 (perfect negative) to +1 (perfect positive).
'@SET:lblInstructions,WordWrap,True
'@CTRL:Label,lblInput1,,12,50,60,15,Input 1:
'@SET:lblInput1,Font.Bold,True
'@CTRL:ComboBox,cboInput1,,80,48,200,20
'@SET:cboInput1,Style,2
'@CTRL:Label,lblInput2,,12,80,60,15,Input 2:
'@SET:lblInput2,Font.Bold,True
'@CTRL:ComboBox,cboInput2,,80,78,200,20
'@SET:cboInput2,Style,2
'@CTRL:Label,lblCoefficient,,12,112,70,15,Correlation:
'@SET:lblCoefficient,Font.Bold,True
'@CTRL:TextBox,txtCoefficient,,80,110,60,20
'@CTRL:Label,lblRange,,150,112,100,15,(-1.0 to +1.0)
'@SET:lblRange,ForeColor,6579300
'@CTRL:CommandButton,btnAdd,,300,78,120,28,Add Correlation
'@CTRL:Label,lblExisting,,12,148,150,15,Defined Correlations:
'@SET:lblExisting,Font.Bold,True
'@CTRL:ListBox,lstCorrelations,,12,166,460,140
'@SET:lstCorrelations,ColumnCount,3
'@SET:lstCorrelations,ColumnWidths,180;180;60
'@CTRL:CommandButton,btnRemove,,340,316,120,25,Remove Selected
'@CTRL:CommandButton,btnClose,,380,350,80,28,Close
'@SET:btnClose,Cancel,True
'@END
Option Explicit

Private Sub UserForm_Initialize()
    ' Scan for inputs first
    ScanWorkbookForFunctions

    ' Set default coefficient
    Me.Controls("txtCoefficient").Text = "0.5"

    ' Populate input combos
    Dim cboInput1 As MSForms.ComboBox, cboInput2 As MSForms.ComboBox
    Set cboInput1 = Me.Controls("cboInput1")
    Set cboInput2 = Me.Controls("cboInput2")
    Dim i As Long
    For i = 0 To g_InputCount - 1
        Dim entry As String
        entry = g_Inputs(i).InputName & " (" & g_Inputs(i).SheetName & "!" & g_Inputs(i).CellAddress & ")"
        cboInput1.AddItem entry
        cboInput2.AddItem entry
    Next i

    ' Load existing correlations
    Dim lstCorr As MSForms.ListBox
    Set lstCorr = Me.Controls("lstCorrelations")
    For i = 0 To g_CorrelationCount - 1
        lstCorr.AddItem
        lstCorr.List(lstCorr.ListCount - 1, 0) = g_Inputs(g_Correlations(i).Input1Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 1) = g_Inputs(g_Correlations(i).Input2Index).InputName
        lstCorr.List(lstCorr.ListCount - 1, 2) = Format(g_Correlations(i).Coefficient, "0.00")
    Next i
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
