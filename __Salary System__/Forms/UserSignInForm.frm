VERSION 5.00
Begin VB.Form Form89 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   Caption         =   "User Sign In Form"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6090
   LinkTopic       =   "Form12"
   ScaleHeight     =   3120
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Sign In User"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
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
      Left            =   3000
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
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      Caption         =   " Password Strength: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter User Password"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
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
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Form89"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim m As Integer


Private Sub AddRecord()

'Conn = GetDbConn
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "Users", gCon, adOpenDynamic, adLockOptimistic, adCmdTable

rs.AddNew
rs("Username") = Text1.Text
rs("Password") = Text2.Text
rs.Update

rs.Close
Set rs = Nothing

End Sub


Private Function SigninCriteria() As Boolean
'Dim flag As Boolean
'flag = True

If UserIsUnique(Text1.Text, Text2.Text) = False Then
MsgBox "User already exists. Cannot Sign In Duplicate User !!"
SigninCriteria = False
Exit Function
End If

If IsPasswordStrong(Text2.Text) = False Then
MsgBox "Passoword is weak. Enter a strong password to continue"
SigninCriteria = False
Exit Function
End If

SigninCriteria = True

End Function


Private Sub Command1_Click()

If SigninCriteria = True Then
Call AddRecord
MsgBox "New User Successfully Signed In !!"
Call UpdateUserActivity("SIGNIN", "Users", Text1.Text)
Unload Me
End If

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If IsPasswordStrong(Text2.Text) = False Then
Label4.Caption = " Bad"
Label4.ForeColor = vbRed
Else
Label4.Caption = "Good"
Label4.ForeColor = &H8000&
End If
End Sub

