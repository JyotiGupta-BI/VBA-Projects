Attribute VB_Name = "mod_ErrorHandler"
' Module:Centralized error handler with logging.
Option Explicit

Public Sub HandleError(ByVal procName As String)
    MsgBox "Error in procedure: " & procName & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Error"
    LogError procName, Err.Description
    Err.Clear
End Sub

Public Sub LogError(procName As String, errDesc As String)
    Dim ws As Worksheet
    
    ' Check if ErrorLog sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ErrorLog")
    On Error GoTo 0
    
    ' If not found, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "ErrorLog"
        ws.Range("A1:C1").Value = Array("Timestamp", "Procedure", "Error Description")
    End If
    
    ' Append new error entry
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(lastRow, 1).Value = Now
    ws.Cells(lastRow, 2).Value = procName
    ws.Cells(lastRow, 3).Value = errDesc
    
End Sub


