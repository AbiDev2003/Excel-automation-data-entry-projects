Attribute VB_Name = "Module_EmployeeCRUD"
Option Explicit

' ============================================================
' Module     : Module_EmployeeCRUD
' Purpose    : Core database operations
'              - Add, Search, Update, Delete employees
'              - All interactions go through the Employee sheet
' ============================================================

'-------------------------------------------------------------
' Adds a new employee to the database
' Expects frm to have named controls matching field names
'-------------------------------------------------------------
Public Sub AddNewEmployee(ByRef frm As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    
    Dim lastRow As Long
    lastRow = Module_Utilities.GetLastRow(ws, 1) + 1
    
    ws.Cells(lastRow, 1).Value = CLng(frm.txtEmployeeID.Text)
    ws.Cells(lastRow, 2).Value = Trim(frm.txtFullName.Text)
    ws.Cells(lastRow, 3).Value = Trim(frm.cboDepartment.Text)
    ws.Cells(lastRow, 4).Value = Trim(frm.txtDesignation.Text)
    ws.Cells(lastRow, 5).Value = Trim(frm.txtEmail.Text)
    ws.Cells(lastRow, 6).Value = Trim(frm.txtPhone.Text)
    ws.Cells(lastRow, 7).Value = Trim(frm.txtJoiningDate.Text)
    ws.Cells(lastRow, 8).Value = Val(frm.txtSalary.Text)
    
    ws.Columns("A:H").AutoFit
    Module_Utilities.LogAction "ADD", frm.txtEmployeeID.Text
End Sub

'-------------------------------------------------------------
' Searches for an employee by ID
' Returns a 7-element array of field values
' Returns empty array if not found
'-------------------------------------------------------------
Public Function SearchEmployeeByID(ByVal empID As String) As Variant
    Dim result(1 To 7) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    
    Dim rowNum As Long
    rowNum = Module_Utilities.GetEmployeeRow(empID)
    
    If rowNum = 0 Then
        SearchEmployeeByID = Array()
        Exit Function
    End If
    
    result(1) = ws.Cells(rowNum, 2).Value  ' Full Name
    result(2) = ws.Cells(rowNum, 3).Value  ' Department
    result(3) = ws.Cells(rowNum, 4).Value  ' Designation
    result(4) = ws.Cells(rowNum, 5).Value  ' Email
    result(5) = ws.Cells(rowNum, 6).Value  ' Phone
    result(6) = ws.Cells(rowNum, 7).Value  ' Joining Date
    result(7) = ws.Cells(rowNum, 8).Value  ' Salary
    
    SearchEmployeeByID = result
End Function

'-------------------------------------------------------------
' Updates an existing employee record
' Searches by Employee ID, then overwrites all fields
'-------------------------------------------------------------
Public Sub UpdateEmployeeRecord(ByRef frm As Object)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    
    Dim rowNum As Long
    rowNum = Module_Utilities.GetEmployeeRow(frm.txtEmployeeID.Text)
    
    If rowNum = 0 Then
        MsgBox "Employee ID not found in database.", vbExclamation, "Update Failed"
        Exit Sub
    End If
    
    ws.Cells(rowNum, 2).Value = Trim(frm.txtFullName.Text)
    ws.Cells(rowNum, 3).Value = Trim(frm.cboDepartment.Text)
    ws.Cells(rowNum, 4).Value = Trim(frm.txtDesignation.Text)
    ws.Cells(rowNum, 5).Value = Trim(frm.txtEmail.Text)
    ws.Cells(rowNum, 6).Value = Trim(frm.txtPhone.Text)
    ws.Cells(rowNum, 7).Value = Trim(frm.txtJoiningDate.Text)
    ws.Cells(rowNum, 8).Value = Val(frm.txtSalary.Text)
    
    ws.Columns("A:H").AutoFit
    Module_Utilities.LogAction "UPDATE", frm.txtEmployeeID.Text
End Sub

'-------------------------------------------------------------
' Deletes an employee record by ID
' Shifts remaining rows up to fill the gap
'-------------------------------------------------------------
Public Sub DeleteEmployeeRecord(ByVal empID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    
    Dim rowNum As Long
    rowNum = Module_Utilities.GetEmployeeRow(empID)
    
    If rowNum = 0 Then
        MsgBox "Employee ID not found in database.", vbExclamation, "Delete Failed"
        Exit Sub
    End If
    
    ws.Rows(rowNum).Delete Shift:=xlUp
    ws.Columns("A:H").AutoFit
    Module_Utilities.LogAction "DELETE", empID
End Sub

'-------------------------------------------------------------
' Returns a 2D array of all employee records
' Used to populate the listbox on the UserForm
' Columns: Employee ID, Full Name, Department, Designation
'-------------------------------------------------------------
Public Function GetAllEmployees() As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    
    Dim lastRow As Long
    lastRow = Module_Utilities.GetLastRow(ws, 1)
    
    If lastRow < 2 Then
        GetAllEmployees = Array()
        Exit Function
    End If
    
    Dim data As Variant
    data = ws.Range("A2:D" & lastRow).Value
    
    GetAllEmployees = data
End Function
