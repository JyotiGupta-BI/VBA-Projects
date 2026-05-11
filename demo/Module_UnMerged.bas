Attribute VB_Name = "Module_UnMerged"

Sub SplitInvoicesMergedData()
    Dim wsDirty As Worksheet, wsClean As Worksheet
    Dim orderID As String, categories As String, amounts As String
    Dim catArr() As String, amtArr() As String
    Dim i As Long, j As Long, outRow As Long
    
    ' Set source and target sheets
    Set wsDirty = ThisWorkbook.Sheets("Merged")
    On Error Resume Next
    Set wsClean = ThisWorkbook.Sheets("UnMerged")
    If wsClean Is Nothing Then
        Set wsClean = ThisWorkbook.Sheets.Add
        wsClean.Name = "Clean"
    End If
    On Error GoTo 0
    
    wsClean.Cells.Clear
    wsClean.Range("A1:C1").Value = Array("Order ID", "Category", "Amount")
    outRow = 2
    
    ' Loop through each order in Dirty sheet
    For i = 2 To wsDirty.Cells(wsDirty.Rows.count, 1).End(xlUp).Row
        orderID = Trim(wsDirty.Cells(i, 1).Value)
        categories = Trim(wsDirty.Cells(i, 2).Value)
        amounts = Trim(wsDirty.Cells(i, 3).Value)
        
        ' Split categories and amounts by "|"
        catArr = Split(categories, "|")
        amtArr = Split(amounts, "|")
        
        ' Write each item and amount on a new row
        For j = LBound(catArr) To UBound(catArr)
            wsClean.Cells(outRow, 1).Value = orderID
            wsClean.Cells(outRow, 2).Value = Trim(catArr(j))
            
            ' Match amount if available
            If j <= UBound(amtArr) Then
                If IsNumeric(Trim(amtArr(j))) Then
                    wsClean.Cells(outRow, 3).Value = CDbl(Trim(amtArr(j)))
                Else
                    wsClean.Cells(outRow, 3).Value = Trim(amtArr(j))
                End If
            End If
            
            outRow = outRow + 1
        Next j
    Next i
    
    MsgBox "Split complete: " & outRow - 2 & " rows created in 'Clean'."
End Sub

