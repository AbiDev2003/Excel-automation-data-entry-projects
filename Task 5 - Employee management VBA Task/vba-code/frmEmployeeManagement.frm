VERSION 5.00
Begin {C62A69C0-16DC-11CE-9E98-00AA00574A4F} frmEmployeeManagement
   Caption         =   "Employee Management System"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   StartUpPosition =   2
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblTitle
      Caption         =   "EMPLOYEE MANAGEMENT SYSTEM"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblSubtitle
      Caption         =   "Manage your employee records"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblEmployeeID
      Caption         =   "Employee ID:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblFullName
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblDepartment
      Caption         =   "Department:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblDesignation
      Caption         =   "Designation:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblEmail
      Caption         =   "Email:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblPhone
      Caption         =   "Phone:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblJoiningDate
      Caption         =   "Joining Date:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblSalary
      Caption         =   "Salary:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin {7BFDAA50-2C48-11CE-9E98-00AA00574A4F} lblSearch
      Caption         =   "Search Employee:"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtEmployeeID
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtFullName
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   1320
      Width           =   2295
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} cboDepartment
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   1680
      Width           =   2295
      Style           =   2
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtDesignation
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   2040
      Width           =   2295
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtEmail
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   2400
      Width           =   2295
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtPhone
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   2760
      Width           =   2295
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtJoiningDate
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   3120
      Width           =   1335
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtSalary
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Top             =   3480
      Width           =   1335
   End
   Begin {3B7C8860-D78F-101B-B9B5-04021C009402} txtSearch
      Height          =   285
      Left            =   4560
      TabIndex        =   19
      Top             =   1200
      Width           =   1575
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnSearch
      Caption         =   "Search"
      Height          =   285
      Left            =   6240
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
   Begin {34A378A0-11FB-11CE-8F81-00AA005BA74F} lstEmployees
      Height          =   2295
      Left            =   4560
      TabIndex        =   21
      Top             =   1680
      Width           =   2655
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnAdd
      Caption         =   "Add Employee"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   4080
      Width           =   1335
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnUpdate
      Caption         =   "Update Employee"
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   4080
      Width           =   1335
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnDelete
      Caption         =   "Delete Employee"
      Height          =   375
      Left            =   3240
      TabIndex        =   24
      Top             =   4080
      Width           =   1335
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnClear
      Caption         =   "Clear Form"
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnRefresh
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   2040
      TabIndex        =   26
      Top             =   4560
      Width           =   1095
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnExit
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   27
      Top             =   4560
      Width           =   1095
   End
End
Attribute VB_Name = "frmEmployeeManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
