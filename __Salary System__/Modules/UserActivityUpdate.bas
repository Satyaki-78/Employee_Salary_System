Attribute VB_Name = "Module6"
Public Function UpdateUserActivity(activity As String, tblName As String, recordId As String)

'activityTime As Date
Dim CurrTime As Date
CurrTime = Time
Dim activityTime As String
activityTime = Format(CurrTime, "hh:mm:ss AM/PM")
'MsgBox activityTime

Dim rsUA As ADODB.Recordset
Set rsUA = New ADODB.Recordset
rsUA.CursorLocation = adUseClient
rsUA.Open "UserActivity", gCon, adOpenDynamic, adLockOptimistic, adCmdTable

rsUA.AddNew
rsUA("Username") = GetCurrentUser()
rsUA("Activity") = activity
rsUA("TableName") = tblName
rsUA("RecordID") = recordId
rsUA("ActivityTime") = activityTime
rsUA.Update
rsUA.Close

End Function

