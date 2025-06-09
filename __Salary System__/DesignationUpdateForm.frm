VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Admin Designation Update"
   ClientHeight    =   6045
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10380
   LinkTopic       =   "Form4"
   ScaleHeight     =   6045
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Designation"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Designation"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5160
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add Designation From Here"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Designations"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If DesigExists(Text1.Text) = False Then
Call AddDesignation
MsgBox "Designation added successfully!!"
Call UpdateUserActivity("DESIGNATION ADD", "Designation", Text1.Text)
Else
MsgBox "Cannot Add Duplicate Designation!!"
End If

End Sub


Private Function AddDesignation()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "Designation", gCon, adOpenDynamic, adLockOptimistic, adCmdTable

rs.AddNew
rs.fields(0) = Text1.Text
rs.Update
rs.Close
Call Display_Designation
End Function


Private Function DesigExists(desigName As String)
Dim rs11 As ADODB.Recordset
Set rs11 = New ADODB.Recordset
rs11.Open "SELECT * FROM Designation WHERE Desig = '" & desigName & "'", gCon, adOpenStatic, adLockReadOnly

If rs11.RecordCount = 0 Then
DesigExists = False
Else
DesigExists = True
End If
End Function


Private Function Display_Designation()

List1.Clear

Dim rs12 As ADODB.Recordset
Set rs12 = New ADODB.Recordset
rs12.Open "SELECT * FROM Designation", gCon, adOpenStatic, adLockReadOnly

rs12.MoveFirst
Do Until rs12.EOF
With List1
.AddItem rs12.fields(0)
End With
rs12.MoveNext
Loop

End Function


Private Function DeleteDesignation()

Dim flag As Boolean
Dim rs12 As ADODB.Recordset
Set rs12 = New ADODB.Recordset
rs12.CursorLocation = adUseClient
rs12.Open "Designation", gCon, adOpenDynamic, adLockOptimistic, adCmdTable

If rs12.RecordCount = 0 Then
MsgBox "Cannot Delete !! No Existing Designations !!"
Exit Function
End If

rs12.MoveFirst
Do Until rs12.EOF
If rs12("Desig") = Text1.Text Then
flag = True
Exit Do
Else
flag = False
rs12.MoveNext
End If
Loop

If flag = True Then
rs12.Delete
rs12.Close
Call Display_Designation
End If

End Function


Private Sub Command2_Click()

If DesigExists(Text1.Text) = True Then
Dim response As Integer
reponse = MsgBox("Are you sure you want to delete this record ?", vbYesNo, "Confirmation")
If reponse = vbYes Then
Call DeleteDesignation
MsgBox "Designation deleted successfully !!"
End If
Else
MsgBox "Cannot Delete !! Designation "" & Text1.Text & "" Doesnt Exist !!"
End If

End Sub

Private Sub Form_Load()

Call OpenDbConn

Call Display_Designation

End Sub


Private Sub List1_DblClick()
Text1.Text = List1.List(List1.ListIndex)
End Sub
