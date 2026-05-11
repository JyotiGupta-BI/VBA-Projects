Attribute VB_Name = "Module_DataCleaning"
Option Explicit

' Master procedure
Sub CleanLaunchCodeData()
    Call HandleMissingData
    Call DetectIrregularData
    Call RemoveUnnecessaryData
    Call FixInconsistentData
    MsgBox "All cleaning exercises completed!"
End Sub

' 1. Missing Data
Sub HandleMissingData()
    Dim ws As Worksheet, rng As Range, cell As Range
    Set ws = ActiveSheet
    Set rng = ws.UsedRange
    
    For Each cell In rng
        If IsEmpty(cell.Value) Then
            cell.Interior.Color = vbYellow ' highlight blanks
        End If
    Next cell
End Sub

' 2. Irregular Data (outliers in transaction_total)
Sub DetectIrregularData()
    Dim ws As Worksheet, rng As Range, cell As Range
    Dim Q1 As Double, Q3 As Double, IQR As Double
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row ' transaction_total in column L
    Set rng = ws.Range("L2:L" & lastRow)
    
    ' Use Percentile for quartiles
    Q1 = Application.WorksheetFunction.Percentile(rng, 0.25)
    Q3 = Application.WorksheetFunction.Percentile(rng, 0.75)
    IQR = Q3 - Q1
    
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value > Q3 + 1.5 * IQR Or cell.Value < Q1 - 1.5 * IQR Then
                cell.Interior.Color = vbRed ' highlight outlier
            End If
        End If
    Next cell
End Sub

' 3. Unnecessary Data (duplicate columns)
Sub RemoveUnnecessaryData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Example: if there are two "email" columns, delete the second
    Dim col As Long
    For col = ws.UsedRange.Columns.Count To 1 Step -1
        If LCase(ws.Cells(1, col).Value) = "email" And col <> 3 Then
            ws.Columns(col).Delete
        End If
    Next col
End Sub

' 4. Inconsistent Data (invalid emails, transaction_total formatting)
Sub FixInconsistentData()
    Dim ws As Worksheet, rng As Range, cell As Range
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row ' email column C
    Set rng = ws.Range("C2:C" & lastRow)
    
    ' Validate emails
    For Each cell In rng
        If Len(cell.Value) > 0 Then
            If InStr(cell.Value, "@") = 0 Or Left(cell.Value, 1) = "@" Then
                cell.Interior.Color = vbRed ' highlight invalid email
            End If
        End If
    Next cell
    
    ' Clean transaction_total (column L)
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    Set rng = ws.Range("L2:L" & lastRow)
    
    For Each cell In rng
        If Len(cell.Value) > 0 Then
            cell.Value = Replace(cell.Value, "$", "")
            If IsNumeric(cell.Value) Then
                cell.Value = CDbl(cell.Value)
            End If
        End If
    Next cell
End Sub

