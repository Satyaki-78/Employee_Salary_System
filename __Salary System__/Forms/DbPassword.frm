VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbPass As String

Private Sub Form_Load()
Dim con As ADODB.Connection
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Users.mdb;JET OLEDB:Database Password=1234;" ' & Pass
con.Open
If con.State = adStateOpen Then
Attempt_DB_Connection = True
MsgBox "Users DB Connected"
Else
Attempt_DB_Connection = False
MsgBox "Users DB Failed to Connected"
End If
End Sub
