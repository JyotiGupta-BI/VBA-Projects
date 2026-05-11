Attribute VB_Name = "mod_JumbledDataClean"
Option Explicit
Sub SplitJumbledDataToNewSheet()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim parts() As String
    Dim i As Long
    Dim lastRow As Long
    Dim addr As String
    
    ' Source sheet = first sheet
    Set wsSource = ThisWorkbook.Sheets(1)
    
    ' Create new sheet for cleaned data
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("CleanedData")
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add
        wsTarget.Name = "CleanedData"
    End If
    On Error GoTo 0
    
    ' Add headers
    wsTarget.Cells(1, 1).Value = "Name"
    wsTarget.Cells(1, 2).Value = "Address"
    wsTarget.Cells(1, 3).Value = "Age"
    wsTarget.Cells(1, 4).Value = "Gender"
    
    ' Find last row in source column A
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    Set rng = wsSource.Range("A2:A" & lastRow)
    
    Dim rowOut As Long
    rowOut = 2
    
    ' Loop through each jumbled cell
    For Each cell In rng
        If Len(cell.Value) > 0 Then
            parts = Split(cell.Value, " ")
            
            ' Extract Name (after "Name")
            wsTarget.Cells(rowOut, 1).Value = parts(1) & " " & parts(2)
            
            ' Extract Address (after "Address")
            addr = ""
            For i = 4 To UBound(parts) - 4
                addr = addr & parts(i) & " "
            Next i
            wsTarget.Cells(rowOut, 2).Value = Trim(addr)
            
            ' Extract Age (after "Age")
            wsTarget.Cells(rowOut, 3).Value = parts(UBound(parts) - 2)
            
            ' Extract Gender (after "Gender")
            wsTarget.Cells(rowOut, 4).Value = parts(UBound(parts))
            
            rowOut = rowOut + 1
        End If
    Next cell
    
    MsgBox "Data split complete! Check the 'CleanedData' sheet."
End Sub

