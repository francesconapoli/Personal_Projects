Attribute VB_Name = "SampleModel"
'===============================================================================
' SampleModel.bas
' Purpose: Creates a sample Monte Carlo model to demonstrate the addin
' Import this into any workbook after installing MCSimAddin
' Then run: CreateSampleModel
'===============================================================================
Option Explicit

Public Sub CreateSampleModel()
    '===================================================================
    ' Sample: Project Cost Estimation with Uncertainty
    '===================================================================
    Dim ws As Worksheet

    ' Create or clear the sheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("MC Sample Model")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "MC Sample Model"
    Else
        ws.Cells.Clear
    End If

    ' ---- Title ----
    ws.Range("A1").Value = "PROJECT COST ESTIMATION - Monte Carlo Model"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1:G1").Merge

    ws.Range("A3").Value = "This sample model demonstrates MCSimAddin functions."
    ws.Range("A3").Font.Italic = True
    ws.Range("A3:G3").Merge

    ' ---- Input Section ----
    ws.Range("A5").Value = "INPUTS (Uncertain Variables)"
    ws.Range("A5").Font.Bold = True
    ws.Range("A5").Font.Size = 11
    ws.Range("A5").Interior.Color = RGB(70, 130, 180)
    ws.Range("A5").Font.Color = RGB(255, 255, 255)
    ws.Range("A5:G5").Merge

    ' Headers
    ws.Range("A6").Value = "Cost Item"
    ws.Range("B6").Value = "Distribution"
    ws.Range("C6").Value = "Min / Mean"
    ws.Range("D6").Value = "Most Likely / StdDev"
    ws.Range("E6").Value = "Max"
    ws.Range("F6").Value = "Formula"
    ws.Range("G6").Value = "Static Value"
    ws.Range("A6:G6").Font.Bold = True

    ' Input 1: Labor Cost (Triangular)
    ws.Range("A7").Value = "Labor Cost"
    ws.Range("B7").Value = "Triangular"
    ws.Range("C7").Value = 50000
    ws.Range("D7").Value = 75000
    ws.Range("E7").Value = 120000
    ws.Range("F7").formula = "=MCTriang(C7,D7,E7)"
    ws.Range("G7").formula = "=(C7+4*D7+E7)/6"

    ' Input 2: Material Cost (Normal)
    ws.Range("A8").Value = "Material Cost"
    ws.Range("B8").Value = "Normal"
    ws.Range("C8").Value = 30000
    ws.Range("D8").Value = 5000
    ws.Range("F8").formula = "=MCNormal(C8,D8)"
    ws.Range("G8").formula = "=C8"

    ' Input 3: Equipment Cost (PERT)
    ws.Range("A9").Value = "Equipment Cost"
    ws.Range("B9").Value = "PERT"
    ws.Range("C9").Value = 15000
    ws.Range("D9").Value = 20000
    ws.Range("E9").Value = 40000
    ws.Range("F9").formula = "=MCPERT(C9,D9,E9)"
    ws.Range("G9").formula = "=(C9+4*D9+E9)/6"

    ' Input 4: Contingency (Lognormal)
    ws.Range("A10").Value = "Contingency"
    ws.Range("B10").Value = "Lognormal"
    ws.Range("C10").Value = 10000
    ws.Range("D10").Value = 5000
    ws.Range("F10").formula = "=MCLognorm(C10,D10)"
    ws.Range("G10").formula = "=C10"

    ' Input 5: Risk Events (Binomial)
    ws.Range("A11").Value = "Risk Events (#)"
    ws.Range("B11").Value = "Binomial"
    ws.Range("C11").Value = 5
    ws.Range("D11").Value = 0.3
    ws.Range("F11").formula = "=MCBinomial(C11,D11)"
    ws.Range("G11").formula = "=ROUND(C11*D11,0)"

    ' Input 6: Cost per Risk Event (Uniform)
    ws.Range("A12").Value = "Cost per Risk Event"
    ws.Range("B12").Value = "Uniform"
    ws.Range("C12").Value = 2000
    ws.Range("D12").Value = 8000
    ws.Range("F12").formula = "=MCUniform(C12,D12)"
    ws.Range("G12").formula = "=(C12+D12)/2"

    ' Input 7: Duration Factor (Beta)
    ws.Range("A13").Value = "Duration Factor"
    ws.Range("B13").Value = "Beta General"
    ws.Range("C13").Value = 0.8
    ws.Range("D13").Value = 1.5
    ws.Range("F13").formula = "=MCBetaGeneral(2,5,C13,D13)"
    ws.Range("G13").formula = "=(C13+D13)/2"

    ' ---- Output Section ----
    ws.Range("A15").Value = "OUTPUTS"
    ws.Range("A15").Font.Bold = True
    ws.Range("A15").Font.Size = 11
    ws.Range("A15").Interior.Color = RGB(50, 150, 50)
    ws.Range("A15").Font.Color = RGB(255, 255, 255)
    ws.Range("A15:G15").Merge

    ' Total Base Cost
    ws.Range("A16").Value = "Total Base Cost"
    ws.Range("F16").formula = "=F7+F8+F9+F10+MCOutput(""Total Base Cost"")"
    ws.Range("G16").formula = "=G7+G8+G9+G10"

    ' Total Risk Cost
    ws.Range("A17").Value = "Total Risk Cost"
    ws.Range("F17").formula = "=F11*F12+MCOutput(""Total Risk Cost"")"
    ws.Range("G17").formula = "=G11*G12"

    ' Total with Duration Adj
    ws.Range("A18").Value = "Duration-Adjusted Total"
    ws.Range("F18").formula = "=(F16+F17)*F13+MCOutput(""Duration-Adjusted Total"")"
    ws.Range("G18").formula = "=(G16+G17)*G13"

    ' Grand Total
    ws.Range("A19").Value = "GRAND TOTAL"
    ws.Range("A19").Font.Bold = True
    ws.Range("F19").formula = "=F18+MCOutput(""Grand Total"")"
    ws.Range("F19").Font.Bold = True
    ws.Range("G19").formula = "=G18"
    ws.Range("G19").Font.Bold = True

    ' ---- Statistics Section ----
    ws.Range("A21").Value = "SIMULATION RESULTS (populated after running simulation)"
    ws.Range("A21").Font.Bold = True
    ws.Range("A21").Font.Size = 11
    ws.Range("A21").Interior.Color = RGB(180, 100, 50)
    ws.Range("A21").Font.Color = RGB(255, 255, 255)
    ws.Range("A21:G21").Merge

    ws.Range("A22").Value = "Statistic"
    ws.Range("B22").Value = "Grand Total"
    ws.Range("A22:B22").Font.Bold = True

    ws.Range("A23").Value = "Mean"
    ws.Range("B23").formula = "=MCMean(F19)"

    ws.Range("A24").Value = "Std Deviation"
    ws.Range("B24").formula = "=MCStdDev(F19)"

    ws.Range("A25").Value = "Minimum"
    ws.Range("B25").formula = "=MCMin(F19)"

    ws.Range("A26").Value = "Maximum"
    ws.Range("B26").formula = "=MCMax(F19)"

    ws.Range("A27").Value = "5th Percentile"
    ws.Range("B27").formula = "=MCPercentile(F19,0.05)"

    ws.Range("A28").Value = "Median (50th)"
    ws.Range("B28").formula = "=MCPercentile(F19,0.5)"

    ws.Range("A29").Value = "95th Percentile"
    ws.Range("B29").formula = "=MCPercentile(F19,0.95)"

    ws.Range("A30").Value = "P(Cost > $150,000)"
    ws.Range("B30").formula = "=MCTargetD(F19,150000)"

    ws.Range("A31").Value = "P(Cost > $200,000)"
    ws.Range("B31").formula = "=MCTargetD(F19,200000)"

    ' Formatting
    ws.Range("F7:G19").NumberFormat = "$#,##0"
    ws.Range("B23:B29").NumberFormat = "$#,##0"
    ws.Range("B30:B31").NumberFormat = "0.00%"
    ws.Columns("A:G").AutoFit

    ' Color coding
    ws.Range("F7:F13").Interior.Color = RGB(220, 230, 255)  ' Light blue for inputs
    ws.Range("F16:F19").Interior.Color = RGB(220, 255, 220)  ' Light green for outputs

    ws.Activate
    ws.Range("A1").Select

    MsgBox "Sample model created on 'MC Sample Model' sheet!" & vbCrLf & vbCrLf & _
           "Column F contains MC distribution functions." & vbCrLf & _
           "Column G contains static expected values." & vbCrLf & vbCrLf & _
           "To run simulation:" & vbCrLf & _
           "1. Go to the MC Simulation ribbon tab" & vbCrLf & _
           "2. Click 'Run' button" & vbCrLf & _
           "3. View results in 'Browse Results'" & vbCrLf & vbCrLf & _
           "Rows 22-31 will auto-populate with statistics after simulation.", _
           vbInformation, "MCSimAddin - Sample Model"
End Sub
