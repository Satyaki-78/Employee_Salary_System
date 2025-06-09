Attribute VB_Name = "Module4"
Public Function IsPasswordStrong(pwd As String) As Boolean
    Dim re As RegExp
    
    ' Check length
    If Len(pwd) < 12 Then
        IsPasswordStrong = False
        Exit Function
    End If

    Set re = New RegExp
    re.IgnoreCase = False
    re.Global = False

    ' Check for at least one digit
    re.Pattern = "\d"
    If Not re.Test(pwd) Then
        IsPasswordStrong = False
        Exit Function
    End If

    ' Check for at least one lowercase letter
    re.Pattern = "[a-z]"
    If Not re.Test(pwd) Then
        IsPasswordStrong = False
        Exit Function
    End If

    ' Check for at least one uppercase letter
    re.Pattern = "[A-Z]"
    If Not re.Test(pwd) Then
        IsPasswordStrong = False
        Exit Function
    End If

    ' Check for at least one special character (non-alphanumeric)
    re.Pattern = "[^a-zA-Z0-9]"
    If Not re.Test(pwd) Then
        IsPasswordStrong = False
        Exit Function
    End If

    ' All checks passed
    IsPasswordStrong = True
End Function

