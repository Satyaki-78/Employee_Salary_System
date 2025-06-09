Attribute VB_Name = "Module3"
Public Function UserIsUnique(userName As String, password As String) As Boolean

'Call OpenDbConn
'If gCon.State = adStateClosed Then gCon.Open
'gCon.Open

'conn = GetDbConn
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Dim sql As String
sql = "SELECT Username FROM Users UNION SELECT Username FROM AdminUsers;"

rs.Open sql, gCon, adOpenStatic, adLockReadOnly

'Output to msgbox showing
'Dim output As String
'output = ""
'Do While Not rs.EOF
'    output = output & rs("Username") & " - " & rs("Password") & vbCrLf
'    rs.MoveNext
'Loop
'MsgBox output
'Exit Function

' Ensure the global connection is initialized
'If gCon Is Nothing Then SetDbConn
'If gCon.State = adStateClosed Then gCon.Open


Dim userFound As Boolean
userFound = False

'Search Code for given Username and Password
If rs.RecordCount = 0 Then
userFound = False
Else

rs.MoveFirst
Do Until rs.EOF

If rs("Username") = userName Then
userFound = True
Exit Do

Else
userFound = False
rs.MoveNext
End If

Loop
End If

If userFound = False Then
UserIsUnique = True
Else
UserIsUnique = False
End If

'gCon.Close

End Function







