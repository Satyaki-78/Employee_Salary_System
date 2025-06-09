VERSION 5.00
Begin VB.Form Form77 
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
      Left            =   2400
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
      Left            =   2400
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
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter AdminName"
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
      Width           =   2055
   End
End
Attribute VB_Name = "Form77"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim m As Integer


Private Function LoginSuccessAction()
MDIForm1.MDEntry.Enabled = True
MDIForm1.UserLogin.Enabled = False
MDIForm1.AdminLogin.Enabled = False
Call SetCurrentUser(Text1.Text)
Call UpdateUserActivity("LOGIN", "AdminUsers", Text1.Text)
End Function


Private Function AuthCriteria() As Boolean

If UserExists(Text1.Text, Text2.Text) = True Then
AuthCriteria = True
Call LoginSuccessAction
Else
AuthCriteria = False
End If

End Function


Private Function Run_DB_Login_Process()

If AuthCriteria = True Then
MsgBox "Login Successfull !!"
Unload Me
Else
MsgBox "Login Failed. Unauthorized Admin User !!"
End If

End Function


Private Sub Command1_Click()

Call Run_DB_Login_Process

End Sub
