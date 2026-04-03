'===============================================================================
' Form: frmOutputSelect
' Purpose: Simple output selector dialog
'===============================================================================
'@FORM:320,260,Select Output
'@CTRL:Label,lblPrompt,,12,8,280,15,Select an output to analyze:
'@SET:lblPrompt,Font.Bold,True
'@CTRL:ListBox,lstOutputs,,12,28,280,150
'@CTRL:CommandButton,btnOK,,120,190,80,25,OK
'@SET:btnOK,Default,True
'@CTRL:CommandButton,btnCancel,,210,190,80,25,Cancel
'@SET:btnCancel,Cancel,True
'@END
Option Explicit

Public SelectedIndex As Long

Private Sub UserForm_Initialize()
    SelectedIndex = -1

    Dim lstOutputs As MSForms.ListBox
    Set lstOutputs = Me.Controls("lstOutputs")
    Dim i As Long
    For i = 0 To g_OutputCount - 1
        lstOutputs.AddItem g_Outputs(i).OutputName & " (" & g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress & ")"
    Next i
    If g_OutputCount > 0 Then lstOutputs.ListIndex = 0
End Sub

Private Sub btnOK_Click()
    SelectedIndex = Me.Controls("lstOutputs").ListIndex
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    SelectedIndex = -1
    Me.Hide
End Sub
