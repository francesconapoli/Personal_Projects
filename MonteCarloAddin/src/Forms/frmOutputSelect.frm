'===============================================================================
' Form: frmOutputSelect
' Purpose: Simple output selector dialog
'===============================================================================
Option Explicit

Public SelectedIndex As Long

Private Sub UserForm_Initialize()
    Me.Caption = "Select Output"
    Me.Width = 320
    Me.Height = 260
    SelectedIndex = -1

    Dim lblPrompt As MSForms.Label
    Set lblPrompt = Me.Controls.Add("Forms.Label.1", "lblPrompt")
    With lblPrompt
        .Caption = "Select an output to analyze:"
        .Left = 12: .Top = 8: .Width = 280: .Height = 15
        .Font.Bold = True
    End With

    Dim lstOutputs As MSForms.ListBox
    Set lstOutputs = Me.Controls.Add("Forms.ListBox.1", "lstOutputs")
    With lstOutputs
        .Left = 12: .Top = 28: .Width = 280: .Height = 150
    End With

    Dim i As Long
    For i = 0 To g_OutputCount - 1
        lstOutputs.AddItem g_Outputs(i).OutputName & " (" & g_Outputs(i).SheetName & "!" & g_Outputs(i).CellAddress & ")"
    Next i
    If g_OutputCount > 0 Then lstOutputs.ListIndex = 0

    Dim btnOK As MSForms.CommandButton
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With btnOK
        .Caption = "OK"
        .Left = 120: .Top = 190: .Width = 80: .Height = 25
        .Default = True
    End With

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Caption = "Cancel"
        .Left = 210: .Top = 190: .Width = 80: .Height = 25
        .Cancel = True
    End With
End Sub

Private Sub btnOK_Click()
    SelectedIndex = Me.Controls("lstOutputs").ListIndex
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    SelectedIndex = -1
    Me.Hide
End Sub
