VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16200
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub btnExportReport_Click()

    On Error GoTo ErrHandler
    
    Dim cf As New CashFlow
    cf.Revenue = CDbl(txtRevenue.Value)
    cf.Expenses = CDbl(txtExpense.Value)
    cf.GrowthRate = CDbl(txtGrowthRate.Value)
    
    GenerateFullReport cf, CDbl(txtDiscountRate.Value), _
                       CInt(txtPeriods.Value), _
                       CLng(txtTrials.Value), _
                       "Financial_Model_Report"
    
    MsgBox "Report exported successfully!", vbInformation
    Exit Sub
    
ErrHandler:
    HandleError "btnExportReport_Click"

End Sub

Private Sub btnRunDCF_Click()

    On Error GoTo ErrHandler
    
    Dim cf As New CashFlow
    cf.Revenue = CDbl(txtRevenue.Value)
    cf.Expenses = CDbl(txtExpense.Value)
    cf.GrowthRate = CDbl(txtGrowthRate.Value)
    
    Dim discountRate As Double
    Dim periods As Integer
    discountRate = CDbl(txtDiscountRate.Value)
    periods = CInt(txtPeriods.Value)
    
    Dim result As Double
    result = DiscountedCashFlow(cf, discountRate, periods)
    
    'MsgBox "DCF Result: " & Format(result, "#,##0.00"), vbInformation
    
    ' Clear and add results to list box of user form
    'lstResults.Clear
    lstResults.AddItem "DCF Result" & vbTab & Format(result, "#,##0.00")
    
    Exit Sub
ErrHandler:
    HandleError "btnRunDCF_Click"


End Sub

Private Sub btnRunMonteCarlo_Click()

    On Error GoTo ErrHandler
    
    Dim cf As New CashFlow
    cf.Revenue = CDbl(txtRevenue.Value)
    cf.Expenses = CDbl(txtExpense.Value)
    cf.GrowthRate = CDbl(txtGrowthRate.Value)
    
    Dim discountRate As Double
    Dim periods As Integer
    Dim trials As Long
    discountRate = CDbl(txtDiscountRate.Value)
    periods = CInt(txtPeriods.Value)
    trials = CLng(txtTrials.Value)
    
    Dim result As Double
    result = MonteCarloDCF(cf, discountRate, periods, trials)
    
    'MsgBox "Monte Carlo Average DCF: " & Format(result, "#,##0.00"), vbInformation
    
    ' add results to list box of user form
    lstResults.AddItem "MC Result" & vbTab & Format(result, "#,##0.00")
    
    Exit Sub
ErrHandler:
    HandleError "btnRunMonteCarlo_Click"

End Sub



Private Sub btnRunScenario_Click()

    Dim cf As New CashFlow
    cf.Revenue = CDbl(txtRevenue.Value)
    cf.Expenses = CDbl(txtExpense.Value)
    cf.GrowthRate = CDbl(txtGrowthRate.Value)
    
    'lstResults.Clear
    
        ' Section header
    lstResults.AddItem "Scenario Analysis" & vbTab & ""
    
    
    lstResults.AddItem "    Base Case" & vbTab & _
        Format(ScenarioDCF(cf, CDbl(txtDiscountRate.Value), CInt(txtPeriods.Value), BaseCase), "#,##0.00")
    
    lstResults.AddItem "    Best Case" & vbTab & _
        Format(ScenarioDCF(cf, CDbl(txtDiscountRate.Value), CInt(txtPeriods.Value), BestCase), "#,##0.00")
    
    lstResults.AddItem "    Worst Case" & vbTab & _
        Format(ScenarioDCF(cf, CDbl(txtDiscountRate.Value), CInt(txtPeriods.Value), WorstCase), "#,##0.00")

End Sub

Private Sub UserForm_Initialize()

    txtRevenue.Value = 100000
    txtExpense.Value = 40000
    txtGrowthRate.Value = 0.05
    txtDiscountRate.Value = 0.1
    txtPeriods.Value = 5
    txtTrials.Value = 1000

End Sub
