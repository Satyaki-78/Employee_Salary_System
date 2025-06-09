Attribute VB_Name = "Module5"
Public Function UserExists(tableName As String, username As String, password As String) As Boolean

Dim userFound As Boolean

'Call OpenDbConn

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM " & tableName, gCon, adOpenDynamic, adLockReadOnly

'Search Code for given Username and Password
If rs.RecordCount = 0 Then
userFound = False
Else

rs.MoveFirst
Do Until rs.EOF

If rs("Username") = username And rs("Password") = password Then
userFound = True
Exit Do
Else
rs.MoveNext
userFound = False
End If

Loop
End If

If userFound = False Then
UserExists = False
Else
UserExists = True
End If

'Call CloseDbConn

End Function
