Attribute VB_Name = "frmEmployeeManagement_Code"
Option Explicit

Private Sub UserForm_Initialize()
    InitializeControls
End Sub

Private Sub btnAdd_Click()
    Dim errMsg As String
    errMsg = Module_Validation.ValidateAllFields(Me)
    If errMsg <> "" Then
        MsgBox "Please fix the following errors:" & vbCrLf & vbCrLf & errMsg, vbExclamation, "Validation Error"
        Exit Sub
    End If
    If Module_Validation.IsDuplicateID(txtEmployeeID.Text) Then
        MsgBox "Employee ID " & txtEmployeeID.Text & " already exists." & vbCrLf & "Please use a different ID.", vbExclamation, "Duplicate ID"
        Exit Sub
    End If
    Module_EmployeeCRUD.AddNewEmployee Me
    MsgBox "Employee " & txtFullName.Text & " added successfully!", vbInformation, "Success"
    Module_Utilities.ClearForm Me
    PopulateEmployeeList
    txtEmployeeID.Text = Module_Utilities.GetNextEmployeeID
    txtEmployeeID.SetFocus
End Sub

Private Sub btnSearch_Click()
    SearchEmployee
End Sub

Private Sub btnUpdate_Click()
    If Not Module_Validation.IsRequiredField(txtEmployeeID.Text) Then
        MsgBox "Please enter or search for an Employee ID first.", vbExclamation, "Update Failed"
        Exit Sub
    End If
    Dim errMsg As String
    errMsg = Module_Validation.ValidateAllFields(Me)
    If errMsg <> "" Then
        MsgBox "Please fix the following errors:" & vbCrLf & vbCrLf & errMsg, vbExclamation, "Validation Error"
        Exit Sub
    End If
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Are you sure you want to update Employee ID: " & txtEmployeeID.Text & "?", vbQuestion + vbYesNo, "Confirm Update")
    If confirm = vbNo Then Exit Sub
    Module_EmployeeCRUD.UpdateEmployeeRecord Me
    MsgBox "Employee " & txtFullName.Text & " updated successfully!", vbInformation, "Success"
    Module_Utilities.ClearForm Me
    PopulateEmployeeList
    txtEmployeeID.Text = Module_Utilities.GetNextEmployeeID
    txtEmployeeID.SetFocus
End Sub

Private Sub btnDelete_Click()
    If Not Module_Validation.IsRequiredField(txtEmployeeID.Text) Then
        MsgBox "Please enter or search for an Employee ID first.", vbExclamation, "Delete Failed"
        Exit Sub
    End If
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Are you sure you want to DELETE Employee ID: " & txtEmployeeID.Text & " (" & txtFullName.Text & ")?", vbQuestion + vbYesNo, "Confirm Delete")
    If confirm = vbNo Then Exit Sub
    Module_EmployeeCRUD.DeleteEmployeeRecord txtEmployeeID.Text
    MsgBox "Employee record deleted successfully!", vbInformation, "Success"
    Module_Utilities.ClearForm Me
    PopulateEmployeeList
    txtEmployeeID.Text = Module_Utilities.GetNextEmployeeID
    txtEmployeeID.SetFocus
End Sub

Private Sub btnClear_Click()
    Module_Utilities.ClearForm Me
    txtEmployeeID.Text = Module_Utilities.GetNextEmployeeID
    txtEmployeeID.SetFocus
End Sub

Private Sub btnRefresh_Click()
    PopulateEmployeeList
    MsgBox "Employee list refreshed.", vbInformation, "Refresh"
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub lstEmployees_Click()
    If lstEmployees.ListIndex < 0 Then Exit Sub
    PopulateFormWithID lstEmployees.Column(0)
End Sub

Private Sub txtSearch_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SearchEmployee
End Sub

Private Sub InitializeControls()
    cboDepartment.List = Array("Sales", "HR", "IT", "Finance", "Marketing", "Operations")
    txtEmployeeID.Text = Module_Utilities.GetNextEmployeeID
    PopulateEmployeeList
End Sub

Private Sub PopulateEmployeeList()
    lstEmployees.Clear
    Dim data As Variant
    data = Module_EmployeeCRUD.GetAllEmployees
    If UBound(data) < 0 Then Exit Sub
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstEmployees.AddItem data(i, 1)
        lstEmployees.List(lstEmployees.ListCount - 1, 1) = data(i, 2)
        lstEmployees.List(lstEmployees.ListCount - 1, 2) = data(i, 3)
        lstEmployees.List(lstEmployees.ListCount - 1, 3) = data(i, 4)
    Next i
End Sub

Private Sub PopulateFormWithID(ByVal empID As String)
    Dim result As Variant
    result = Module_EmployeeCRUD.SearchEmployeeByID(empID)
    If UBound(result) < 0 Then
        MsgBox "Employee ID not found.", vbExclamation, "Not Found"
        Exit Sub
    End If
    txtEmployeeID.Text = empID
    txtFullName.Text = result(1)
    cboDepartment.Text = result(2)
    txtDesignation.Text = result(3)
    txtEmail.Text = result(4)
    txtPhone.Text = result(5)
    txtJoiningDate.Text = result(6)
    txtSalary.Text = result(7)
End Sub

Private Sub SearchEmployee()
    Dim empID As String
    empID = Trim(txtSearch.Text)
    If empID = "" Then
        MsgBox "Please enter an Employee ID to search.", vbExclamation, "Search Failed"
        Exit Sub
    End If
    PopulateFormWithID empID
    txtSearch.Text = ""
End Sub
