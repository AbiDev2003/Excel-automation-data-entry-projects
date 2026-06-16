Attribute VB_Name = "Module_Utilities"
Option Explicit

' ============================================================
' Module     : Module_Utilities
' Purpose    : Helper functions used across the application
'              - Sheet utilities, logging, formatting
' ============================================================

'-------------------------------------------------------------
' Finds the last used row in a worksheet column
'-------------------------------------------------------------
Public Function GetLastRow(ByRef ws As Worksheet, Optional ByVal col As Long = 1) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    GetLastRow = lastRow
End Function

'-------------------------------------------------------------
' Generates the next Employee ID (max existing + 1)
' Starts at 1001 if no records exist
'-------------------------------------------------------------
Public Function GetNextEmployeeID() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    Dim lastRow As Long
    lastRow = GetLastRow(ws, 1)
    If lastRow = 1 Then
        GetNextEmployeeID = 1001
    Else
        Dim maxID As Long
        maxID = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow))
        GetNextEmployeeID = maxID + 1
    End If
End Function

'-------------------------------------------------------------
' Logs an action to the Logs sheet (auto-creates if missing)
' Columns: Timestamp | Action | Employee ID
'-------------------------------------------------------------
Public Sub LogAction(ByVal action As String, ByVal empID As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Logs")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Logs"
        ws.Range("A1:C1").Value = Array("Timestamp", "Action", "Employee ID")
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns("A:C").AutoFit
    End If
    Dim lastRow As Long
    lastRow = GetLastRow(ws, 1) + 1
    ws.Cells(lastRow, 1).Value = Now
    ws.Cells(lastRow, 2).Value = action
    ws.Cells(lastRow, 3).Value = empID
    ws.Columns("A:C").AutoFit
End Sub

'-------------------------------------------------------------
' Sets up sheets on workbook open
' - Formats Employee Database sheet
' - Creates Logs sheet if missing
'-------------------------------------------------------------
Public Sub InitializeDatabase()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Employee Database")
    FormatEmployeeSheet ws
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Logs")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Logs"
        ws.Range("A1:C1").Value = Array("Timestamp", "Action", "Employee ID")
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns("A:C").AutoFit
    End If
End Sub

'-------------------------------------------------------------
' Applies professional formatting to the Employee sheet
' - Converts to a styled Excel Table with filters
' - Formats Salary column with Indian Rupee symbol
'-------------------------------------------------------------
Public Sub FormatEmployeeSheet(ByRef ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastRow(ws, 1)
    If lastRow < 1 Then lastRow = 1
    
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H" & lastRow), , xlYes)
        tbl.Name = "tblEmployees"
        tbl.TableStyle = "TableStyleMedium9"
    Else
        tbl.Resize ws.Range("A1:H" & lastRow)
    End If
    
    With tbl.DataBodyRange
        .HorizontalAlignment = xlLeft
        .Columns(1).HorizontalAlignment = xlCenter
    End With
    
    With tbl.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    ' Format Salary column (Column 8 / H) with Rupee symbol
    If lastRow > 1 Then
        tbl.DataBodyRange.Columns(8).NumberFormat = ChrW(8377) & "#,##0"
    End If
    
    ws.Columns("A:H").AutoFit
End Sub

'-------------------------------------------------------------
' Clears all TextBox and ComboBox controls on a UserForm
'-------------------------------------------------------------
Public Sub ClearForm(ByRef frm As Object)
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Text = ""
        ElseIf TypeName(ctrl) = "ComboBox" Then
            ctrl.Value = ""
        End If
    Next ctrl
End Sub

'-------------------------------------------------------------
' Finds the row number of an employee by ID
' Returns 0 if not found
'-------------------------------------------------------------
Public Function GetEmployeeRow(ByVal empID As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    Dim lastRow As Long
    lastRow = GetLastRow(ws, 1)
    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = empID Then
            GetEmployeeRow = i
            Exit Function
        End If
    Next i
    GetEmployeeRow = 0
End Function
