Attribute VB_Name = "Module_Validation"
Option Explicit

' ============================================================
' Module     : Module_Validation
' Purpose    : All input validation functions
'              - Each function returns True if valid
'              - ValidateAllFields runs all checks at once
' ============================================================

'-------------------------------------------------------------
' Checks that a field has a non-empty value
'-------------------------------------------------------------
Public Function IsRequiredField(ByVal value As String) As Boolean
    IsRequiredField = (Trim(value) <> "")
End Function

'-------------------------------------------------------------
' Validates email format:
' - Must contain '@' with at least 1 char before it
' - Must contain '.' after '@' with at least 1 char after
'-------------------------------------------------------------
Public Function IsValidEmail(ByVal email As String) As Boolean
    Dim atPos As Long
    Dim dotPos As Long
    atPos = InStr(email, "@")
    If atPos = 0 Then
        IsValidEmail = False
        Exit Function
    End If
    If atPos < 2 Then
        IsValidEmail = False
        Exit Function
    End If
    dotPos = InStr(atPos + 1, email, ".")
    If dotPos = 0 Or dotPos = Len(email) Then
        IsValidEmail = False
        Exit Function
    End If
    IsValidEmail = True
End Function

'-------------------------------------------------------------
' Validates salary:
' - Optional field (empty = valid)
' - Must be numeric and positive if provided
'-------------------------------------------------------------
Public Function IsValidSalary(ByVal salary As String) As Boolean
    If Trim(salary) = "" Then
        IsValidSalary = True
        Exit Function
    End If
    If Not IsNumeric(salary) Then
        IsValidSalary = False
        Exit Function
    End If
    If CDbl(salary) <= 0 Then
        IsValidSalary = False
        Exit Function
    End If
    IsValidSalary = True
End Function

'-------------------------------------------------------------
' Validates Employee ID:
' - Required, numeric, positive, max 10 digits
'-------------------------------------------------------------
Public Function IsValidEmployeeID(ByVal empID As String) As Boolean
    If Trim(empID) = "" Then
        IsValidEmployeeID = False
        Exit Function
    End If
    If Not IsNumeric(empID) Then
        IsValidEmployeeID = False
        Exit Function
    End If
    If CLng(empID) <= 0 Then
        IsValidEmployeeID = False
        Exit Function
    End If
    If Len(Trim(empID)) > 10 Then
        IsValidEmployeeID = False
        Exit Function
    End If
    IsValidEmployeeID = True
End Function

'-------------------------------------------------------------
' Validates phone number:
' - Optional field (empty = valid)
' - Must contain 10-15 digits (ignores dashes/spaces)
'-------------------------------------------------------------
Public Function IsValidPhone(ByVal phone As String) As Boolean
    If Trim(phone) = "" Then
        IsValidPhone = True
        Exit Function
    End If
    Dim digitsOnly As String
    digitsOnly = ""
    Dim i As Long
    For i = 1 To Len(phone)
        If Mid(phone, i, 1) Like "[0-9]" Then
            digitsOnly = digitsOnly & Mid(phone, i, 1)
        End If
    Next i
    If Len(digitsOnly) < 10 Or Len(digitsOnly) > 15 Then
        IsValidPhone = False
        Exit Function
    End If
    IsValidPhone = True
End Function

'-------------------------------------------------------------
' Validates date field:
' - Optional field (empty = valid)
' - Must be a recognized date format if provided
'-------------------------------------------------------------
Public Function IsValidDateField(ByVal dateStr As String) As Boolean
    If Trim(dateStr) = "" Then
        IsValidDateField = True
        Exit Function
    End If
    If Not IsDate(dateStr) Then
        IsValidDateField = False
        Exit Function
    End If
    IsValidDateField = True
End Function

'-------------------------------------------------------------
' Checks if an Employee ID already exists in the database
' Returns True if duplicate found
'-------------------------------------------------------------
Public Function IsDuplicateID(ByVal empID As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee Database")
    Dim lastRow As Long
    lastRow = Module_Utilities.GetLastRow(ws, 1)
    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = empID Then
            IsDuplicateID = True
            Exit Function
        End If
    Next i
    IsDuplicateID = False
End Function

'-------------------------------------------------------------
' Master validation: runs all checks on the form fields
' Returns an error string (empty = all valid)
'-------------------------------------------------------------
Public Function ValidateAllFields(ByRef frm As Object) As String
    Dim errMsg As String
    errMsg = ""
    
    If Not IsRequiredField(frm.txtEmployeeID.Text) Then
        errMsg = errMsg & "- Employee ID is required." & vbCrLf
    ElseIf Not IsValidEmployeeID(frm.txtEmployeeID.Text) Then
        errMsg = errMsg & "- Employee ID must be numeric, positive, and at most 10 digits." & vbCrLf
    End If
    
    If Not IsRequiredField(frm.txtFullName.Text) Then
        errMsg = errMsg & "- Full Name is required." & vbCrLf
    End If
    
    If Not IsRequiredField(frm.cboDepartment.Text) Then
        errMsg = errMsg & "- Department is required." & vbCrLf
    End If
    
    If Not IsRequiredField(frm.txtEmail.Text) Then
        errMsg = errMsg & "- Email is required." & vbCrLf
    ElseIf Not IsValidEmail(frm.txtEmail.Text) Then
        errMsg = errMsg & "- Email must contain '@' with characters before and a valid domain." & vbCrLf
    End If
    
    If Not IsValidPhone(frm.txtPhone.Text) Then
        errMsg = errMsg & "- Phone must contain 10-15 digits." & vbCrLf
    End If
    
    If Not IsValidDateField(frm.txtJoiningDate.Text) Then
        errMsg = errMsg & "- Please enter a valid date for Joining Date." & vbCrLf
    End If
    
    If Not IsValidSalary(frm.txtSalary.Text) Then
        errMsg = errMsg & "- Salary must be a positive number." & vbCrLf
    End If
    
    ValidateAllFields = errMsg
End Function
