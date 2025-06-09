Attribute VB_Name = "Module1"
Private dbPass As String

Public Sub SetDbPswd(ByVal Pswd As String)
dbPass = Pswd
End Sub

Public Function GetDbPswd() As String
GetDbPswd = dbPass
End Function

