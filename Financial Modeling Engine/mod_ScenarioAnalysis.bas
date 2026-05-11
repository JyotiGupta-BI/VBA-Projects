Attribute VB_Name = "mod_ScenarioAnalysis"
' Module: Create a module that runs best-case, base-case, and worst-case assumptions
Option Explicit

Public Enum ScenarioType
    BaseCase = 0
    BestCase = 1
    WorstCase = 2
End Enum

Public Function ScenarioDCF(cf As CashFlow, _
                            discountRate As Double, _
                            periods As Integer, _
                            scenario As ScenarioType) As Double
    
    On Error GoTo ErrHandler
    
    Dim tempCF As New CashFlow
    
    
    ' Copy base values
    tempCF.Revenue = cf.Revenue
    tempCF.Expenses = cf.Expenses
    tempCF.GrowthRate = cf.GrowthRate
    
    ' Adjust assumptions by scenario
    Select Case scenario
        Case BaseCase
            ' no change
        Case BestCase
            tempCF.Revenue = tempCF.Revenue * 1.1
            tempCF.Expenses = tempCF.Expenses * 0.9
            tempCF.GrowthRate = tempCF.GrowthRate + 0.02
        Case WorstCase
            tempCF.Revenue = tempCF.Revenue * 0.9
            tempCF.Expenses = tempCF.Expenses * 1.1
            tempCF.GrowthRate = tempCF.GrowthRate - 0.02
    End Select
    
    ' Run DCF with adjusted assumptions
    ScenarioDCF = DiscountedCashFlow(tempCF, discountRate, periods)
    
    Exit Function
    
ErrHandler:
    LogError "RunScenarioAnalysis", Err.Description

End Function



