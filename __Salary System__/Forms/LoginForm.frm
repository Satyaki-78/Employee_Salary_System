VERSION 5.00
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   Caption         =   "Admin Login Form"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   LinkTopic       =   "Form12"
   ScaleHeight     =   2970
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Log &In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter Username"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim m As Integer
Dim Usrname, Pswd, UserType As String


Private Function Attempt_DB_Connection(Pass As String) As Boolean
On Error Resume Next 'GoTo ErrorHandler

Dim con As ADODB.Connection
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Users.mdb;Persist Security Info=False;JET OLEDB:Database Password = " & Pass
con.Open

If con.State = adStateOpen Then
Attempt_DB_Connection = True
Else
Attempt_DB_Connection = False
End If

Exit Function

ErrorHandler:
Attempt_DB_Connection = False

End Function


Private Function UserExists(username As String, password As String) As Boolean
Dim userFound As Boolean

Conn = GetDbConn
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "Select * from Users", Conn, adOpenStatic, adLockReadOnly

'Search Code for given Username and Password
If rs.RecordCount = 0 Then
userFound = False
Else

rs.MoveFirst
Do Until rs.EOF

If rs("Username") = username And rs("password") = password Then
userFound = True
UserType = rs("UserType")
Exit Do

Else
rs.MoveNext
userFound = False
End If

Loop
End If

If userFound = False Then
UserExists = False
MsgBox "No record with the Username and Password found !!" & vbNewLine & "Register yourself if you are a new User"
Else
UserExists = True
Usrname = username
Pswd = password
End If

End Function


Private Function GetUserType() As String

Dim usrtype As String
If UserExists(Text1.Text, Text2.Text) = True Then
GetUserType = UserType
Exit Function
End If

End Function


Private Function LoginSuccessAction()

If UserType = "admin" Then
MDIForm1.MDEntry.Enabled = True
End If

End Function


Private Function AuthCriteria() As Boolean

If UserExists(Text1.Text, Text2.Text) = True Then
AuthCriteria = True
Call LoginSuccessAction
MDIForm1.LogIn.Enabled = False
Exit Function
Else
AuthCriteria = False
End If

End Function


Private Function Run_DB_Login_Process()

Dim str As String
str = Text1.Text

Do

    If ctr = 3 Then
        If AuthCriteria = False Then
            m = MsgBox("ATTEMPTS EXCEEDED", 16)
            End
            Exit Function
        Else
            m = MsgBox("Authorized User")
            SetDbPswd str
            Unload Me
            Exit Function
        End If
    End If
    
    If AuthCriteria = True And ctr <= 3 Then
        m = MsgBox("Authorized User")
        SetDbPswd str
        Unload Me
        Exit Function
    End If

ctr = ctr + 1
Text2.Text = "Attempt - " & ctr
    
    If ctr = 3 And AuthCriteria = False Then
        Text2.BackColor = vbRed
        Text2.Text = "LAST ATTEMPT"
    End If

    If AuthCriteria = False Then
        m = MsgBox("INVALID PASSWORD", 16)
        Text1.SetFocus
        Exit Do
    End If
    
Loop Until ctr = 3

End Function


Private Sub Command1_Click()

Call Run_DB_Login_Process

End Sub


Private Sub Form_Load()
ctr = 1
'Text2.Text = "Attempt - " & ctr
End Sub
