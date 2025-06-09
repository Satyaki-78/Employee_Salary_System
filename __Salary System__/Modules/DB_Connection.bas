Attribute VB_Name = "Module2"
Public gCon As ADODB.Connection


Public Function SetDbConn()

Set gCon = New ADODB.Connection
gCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=Infosys.mdb;JET OLEDB:Database Password =12345; Persist Security Info=False"
gCon.Open
End Function


Public Function OpenDbConn()

End Function


Public Function CloseDbConn()

End Function


