VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' ThisWorkbook: MCSimAddin
' Purpose: Workbook-level event handlers and initialization
'===============================================================================
Option Explicit

Private Sub Workbook_Open()
    ' Initialize default settings and build the toolbar
    InitializeDefaults
    CreateToolbar
    Application.StatusBar = ADDIN_NAME & " v" & ADDIN_VERSION & " loaded. See 'MC Simulation' toolbar."
End Sub

Private Sub Workbook_AddinInstall()
    InitializeDefaults
    CreateToolbar
    MsgBox ADDIN_NAME & " v" & ADDIN_VERSION & " installed successfully!" & vbCrLf & vbCrLf & _
           "Look for the 'MC Simulation' toolbar at the top of Excel." & vbCrLf & vbCrLf & _
           "Available worksheet functions:" & vbCrLf & _
           "  MCNormal, MCUniform, MCTriang, MCLognorm, MCPERT," & vbCrLf & _
           "  MCBeta, MCGamma, MCWeibull, MCExpon, MCPoisson," & vbCrLf & _
           "  MCBinomial, MCPareto, MCLogistic, MCChiSq, etc." & vbCrLf & vbCrLf & _
           "Use MCOutput() to mark cells as outputs.", _
           vbInformation, ADDIN_NAME
End Sub

Private Sub Workbook_AddinUninstall()
    RemoveToolbar
    Application.StatusBar = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If g_SimState = ssRunning Then
        If MsgBox("Simulation is running. Stop and close?", vbQuestion + vbYesNo) = vbNo Then
            Cancel = True
            Exit Sub
        End If
        StopSimulation
    End If
    RemoveToolbar
    Application.StatusBar = False
End Sub
