Attribute VB_Name = "Module_StructureMultiHeader"

Sub UnpivotSalesData2()
    Dim wsDirty As Worksheet, wsClean As Worksheet
    Dim shipMode As String, segment As String
    Dim orderDate As Variant, saleVal As Variant
    Dim i As Long, j As Long, lastRow As Long, lastCol As Long
    Dim outRow As Long
    
    Set wsDirty = ThisWorkbook.Sheets("Dirty")
    On Error Resume Next
    Set wsClean = ThisWorkbook.Sheets("Clean")
    If wsClean Is Nothing Then
        Set wsClean = ThisWorkbook.Sheets.Add
        wsClean.Name = "Clean"
    End If
    On Error GoTo 0
    
    wsClean.Cells.Clear
    wsClean.Range("A1:D1").Value = Array("Ship Mode", "Segment", "Order Date", "Sales")
    outRow = 2
    
    ' Detect last row and column
    lastRow = wsDirty.Cells(wsDirty.Rows.count, 1).End(xlUp).Row
    lastCol = wsDirty.Cells(2, wsDirty.Columns.count).End(xlToLeft).Column
    
    ' Loop through each Ship Mode / Segment column
    For j = 2 To lastCol
        shipMode = Trim(wsDirty.Cells(1, j).Value)
        segment = Trim(wsDirty.Cells(2, j).Value)
        
        For i = 3 To lastRow
            orderDate = wsDirty.Cells(i, 1).Value
            saleVal = wsDirty.Cells(i, j).Value
            
            ' Only write if both date and numeric sales exist
            If IsDate(orderDate) And Len(saleVal) > 0 And IsNumeric(saleVal) Then
                wsClean.Cells(outRow, 1).Value = shipMode
                wsClean.Cells(outRow, 2).Value = segment
                wsClean.Cells(outRow, 3).Value = CDate(orderDate)
                wsClean.Cells(outRow, 4).Value = CDbl(saleVal)
                outRow = outRow + 1
            End If
        Next i
    Next j
    
    MsgBox "Unpivot complete: " & outRow - 2 & " rows created in 'Clean 2'."
End Sub

