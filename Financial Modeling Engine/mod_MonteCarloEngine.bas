Attribute VB_Name = "mod_MonteCarloEngine"
' Module: Runs stochastic simulations with random growth rates.
Option Explicit

Public Function MonteCarloDCF(cf As CashFlow, _
                              discountRate As Double, _
                              periods As Integer, _
                              trials As Long) As Double
                              
    Dim i As Long
    Dim totalDCF As Double
    Dim tempCF As CashFlow
    
    totalDCF = 0
    
    On Error GoTo ErrHandler
    
    For i = 1 To trials
        Set tempCF = New CashFlow
        tempCF.Revenue = cf.Revenue * (1 + Rnd() * 0.1 - 0.05) ' ±5% random
        tempCF.Expenses = cf.Expenses * (1 + Rnd() * 0.1 - 0.05)
        tempCF.GrowthRate = cf.GrowthRate * (1 + Rnd() * 0.1 - 0.05)
        
        totalDCF = totalDCF + DiscountedCashFlow(tempCF, discountRate, periods)
    Next i
    
    MonteCarloDCF = totalDCF / trials
    
    
    Exit Function
    
ErrHandler:
    LogError "RunMCAnalysis", Err.Description
    
End Function


