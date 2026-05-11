Attribute VB_Name = "Module_SalesData1DataCleaning"
Option Explicit

' === Master Procedure ===
Sub RunPhase1Cleaning()
    Call InitLog
    
    ' LaunchCode dataset cleaning
    Call CleanLaunchCodeData
    
    ' Badly Structured Sales Data 1 cleaning
    Call CleanSalesData1
    
    MsgBox "Phase 1 cleaning completed! Check 'CleaningLog' for details."
End Sub

' === Logging Utilities ===
Sub InitLog()
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("CleaningLog")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "CleaningLog"
        wsLog.Cells(1, 1).Value = "Step"
        wsLog.Cells(1, 2).Value = "Action"
        wsLog.Cells(1, 3).Value = "Details"
    End If
    On Error GoTo 0
End Sub

Sub WriteLog(stepName As String, action As String, details As String)
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Set wsLog = ThisWorkbook.Sheets("CleaningLog")
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(nextRow, 1).Value = stepName
    wsLog.Cells(nextRow, 2).Value = action
    wsLog.Cells(nextRow, 3).Value = details
End Sub

' === LaunchCode Dataset Cleaning ===
Sub CleanLaunchCodeData()
    Call HandleMissingData
    Call DetectIrregularData
    Call RemoveUnnecessaryData
    Call FixInconsistentData
    Call WriteLog("LaunchCode", "Dataset cleaned", "All four exercises applied")
End Sub

Sub HandleMissingData()
    Dim ws As Worksheet, rng As Range, cell As Range
    Set ws = ActiveSheet
    Set rng = ws.UsedRange
    
    For Each cell In rng
        If IsEmpty(cell.Value) Then
            cell.Interior.Color = vbYellow
            Call WriteLog("Missing Data", "Highlighted blank cell", "Row " & cell.Row & ", Col " & cell.Column)
        End If
    Next cell
End Sub

Function GetPercentile(rng As Range, p As Double) As Double
    Dim arr() As Double
    Dim cell As Range
    Dim i As Long, j As Long, temp As Double
    Dim n As Long, pos As Double
    
    ' Load numeric values into array
    n = 0
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            n = n + 1
            ReDim Preserve arr(1 To n)
            arr(n) = cell.Value
        End If
    Next cell
    
    If n = 0 Then
        GetPercentile = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' Sort array (simple bubble sort)
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' Position for percentile
    pos = p * (n + 1)
    
    ' Handle edge cases
    If pos < 1 Then
        GetPercentile = arr(1)
    ElseIf pos >= n Then
        GetPercentile = arr(n)
    ElseIf pos = Int(pos) Then
        GetPercentile = arr(pos)
    Else
        ' Linear interpolation between nearest ranks
        GetPercentile = arr(Int(pos)) + (pos - Int(pos)) * (arr(Int(pos) + 1) - arr(Int(pos)))
    End If
End Function




Sub DetectIrregularData()
    Dim ws As Worksheet, rng As Range, cell As Range
    Dim Q1 As Double, Q3 As Double, IQR As Double
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    Set rng = ws.Range("L2:L" & lastRow)
    
    ' Use Evaluate to call Excel formulas
    Q1 = GetPercentile(rng, 0.25)
    Q3 = GetPercentile(rng, 0.75)
    IQR = Q3 - Q1
    
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value > Q3 + 1.5 * IQR Or cell.Value < Q1 - 1.5 * IQR Then
                cell.Interior.Color = vbRed
                Call WriteLog("Irregular Data", "Outlier flagged", "Row " & cell.Row & " ? " & cell.Value)
            End If
        End If
    Next cell
End Sub

Sub RemoveUnnecessaryData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim col As Long
    For col = ws.UsedRange.Columns.Count To 1 Step -1
        If LCase(ws.Cells(1, col).Value) = "email" And col <> 3 Then
            ws.Columns(col).Delete
            Call WriteLog("Unnecessary Data", "Duplicate column removed", "Column " & col)
        End If
    Next col
End Sub

Sub FixInconsistentData()
    Dim ws As Worksheet, rng As Range, cell As Range
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    Set rng = ws.Range("C2:C" & lastRow)
    
    For Each cell In rng
        If Len(cell.Value) > 0 Then
            If InStr(cell.Value, "@") = 0 Or Left(cell.Value, 1) = "@" Then
                cell.Interior.Color = vbRed
                Call WriteLog("Inconsistent Data", "Invalid email flagged", "Row " & cell.Row & " ? " & cell.Value)
            End If
        End If
    Next cell
    
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    Set rng = ws.Range("L2:L" & lastRow)
    
    For Each cell In rng
        If Len(cell.Value) > 0 Then
            cell.Value = Replace(cell.Value, "$", "")
            If IsNumeric(cell.Value) Then
                cell.Value = CDbl(cell.Value)
                Call WriteLog("Inconsistent Data", "Transaction cleaned", "Row " & cell.Row & " ? " & cell.Value)
            End If
        End If
    Next cell
End Sub

' === Badly Structured Sales Data 1 Cleaning ===
Sub CleanSalesData1()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    Set ws = ActiveSheet
    
    ' Remove merged header
    ws.Rows(1).UnMerge
    ws.Rows(1).Delete
    Call WriteLog("Sales Data 1", "Removed merged header", "Row 1 deleted")
    
    ' Delete blank rows
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If Application.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
            Call WriteLog("Sales Data 1", "Blank row deleted", "Row " & i)
        End If
    Next i
    
    ' Remove totals rows
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If InStr(1, LCase(ws.Cells(i, 1).Value), "total") > 0 Then
            ws.Rows(i).Delete
            Call WriteLog("Sales Data 1", "Totals row removed", "Row " & i)
        End If
    Next i
    
    ' Standardize dates (column B assumed)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, 2).Value) Then
            ws.Cells(i, 2).Value = Format(ws.Cells(i, 2).Value, "yyyy-mm-dd")
            Call WriteLog("Sales Data 1", "Date standardized", "Row " & i & " ? " & ws.Cells(i, 2).Value)
        End If
    Next i
    
    Call WriteLog("Sales Data 1", "Dataset cleaned", "All steps applied")
End Sub

