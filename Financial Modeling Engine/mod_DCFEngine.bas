Attribute VB_Name = "mod_DCFEngine"
' Module: Handles discounted cash flow calculations.
Option Explicit

Public Function DiscountedCashFlow(cf As CashFlow, _
                                   discountRate As Double, _
                                   periods As Integer) As Double
    Dim i As Integer
    Dim cashFlowValue As Double
    Dim totalDCF As Double
    
    On Error GoTo ErrHandler
    
    cashFlowValue = cf.NetCashFlow
    
    For i = 1 To periods
        cashFlowValue = cashFlowValue * (1 + cf.GrowthRate)
        totalDCF = totalDCF + (cashFlowValue / ((1 + discountRate) ^ i))
    Next i
    
    DiscountedCashFlow = totalDCF
    
    Exit Function
    
ErrHandler:
    LogError "RunDCFAnalysis", Err.Description
    
End Function


