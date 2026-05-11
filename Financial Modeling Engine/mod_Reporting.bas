Attribute VB_Name = "mod_Reporting"
' Module: Generate formatted reports and export to PDF.

Option Explicit


Public Sub GenerateFullReport(cf As CashFlow, _
                              discountRate As Double, _
                              periods As Integer, _
                              trials As Long, _
                              reportName As String)
    Dim ws As Worksheet
    
    ' Try to get the Report sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Report")
    On Error GoTo 0
    
    ' If not found, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Report"
    End If
    
    ' Clear old content
    ws.Cells.Clear
    
    ' Header
    ws.Range("A1").Value = "Financial Modeling Report"
    ws.Range("A2").Value = "Generated: " & Now
    
    ' Base inputs
    ws.Range("A4").Value = "Revenue"
    ws.Range("B4").Value = cf.Revenue
    ws.Range("A5").Value = "Expenses"
    ws.Range("B5").Value = cf.Expenses
    ws.Range("A6").Value = "Growth Rate"
    ws.Range("B6").Value = cf.GrowthRate
    ws.Range("A7").Value = "Discount Rate"
    ws.Range("B7").Value = discountRate
    ws.Range("A8").Value = "Periods"
    ws.Range("B8").Value = periods
    ws.Range("A9").Value = "Trials"
    ws.Range("B9").Value = trials
    
    ' Results
    ws.Range("A11").Value = "DCF Result"
    ws.Range("B11").Value = DiscountedCashFlow(cf, discountRate, periods)
    
    ws.Range("A12").Value = "Monte Carlo Avg DCF"
    ws.Range("B12").Value = MonteCarloDCF(cf, discountRate, periods, trials)
    
    ws.Range("A13").Value = "Scenario - Base Case"
    ws.Range("B13").Value = ScenarioDCF(cf, discountRate, periods, BaseCase)
    
    ws.Range("A14").Value = "Scenario - Best Case"
    ws.Range("B14").Value = ScenarioDCF(cf, discountRate, periods, BestCase)
    
    ws.Range("A15").Value = "Scenario - Worst Case"
    ws.Range("B15").Value = ScenarioDCF(cf, discountRate, periods, WorstCase)
    
    ' Format the report nicely
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Size = 14
    End With

    ws.Range("A2").Font.Italic = True

    ' Bold labels
    ws.Range("A4:A15").Font.Bold = True

    ' Add borders
    With ws.Range("A4:B15").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' Currency formatting for results
    ws.Range("B11:B15").NumberFormat = "#,##0.00"
    
    ws.Columns("A:B").AutoFit
  
    ' Export to PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ThisWorkbook.Path & "\" & reportName & ".pdf", _
        Quality:=xlQualityStandard
End Sub



