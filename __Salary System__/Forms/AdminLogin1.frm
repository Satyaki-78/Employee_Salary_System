VERSION 5.00
Begin VB.Form Form99 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   Caption         =   "Form12"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   LinkTopic       =   "Form12"
   ScaleHeight     =   2265
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&CLICK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Height          =   450
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter Database Password below"
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
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim m As Integer


Private Function Attempt_DB_Connection(Pass As String) As Boolean
On Error Resume Next 'GoTo ErrorHandler

Dim Con As ADODB.Connection
Set Con = New ADODB.Connection
Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Users.mdb;Persist Security Info=False;JET OLEDB:Database Password = " & Pass
Con.Open

If Con.State = adStateOpen Then
Attempt_DB_Connection = True
Else
Attempt_DB_Connection = False
End If

Exit Function

ErrorHandler:
Attempt_DB_Connection = False

End Function


Private Function Run_DB_Login_Process()

Dim str As String
str = Text1.Text

Do

    If ctr = 3 Then
        If Attempt_DB_Connection(str) = False Then
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
    
    If Attempt_DB_Connection(str) = True And ctr <= 3 Then
        m = MsgBox("Authorized User")
        SetDbPswd str
        Unload Me
        Exit Function
    End If

ctr = ctr + 1
Text2.Text = "Attempt - " & ctr
    
    If ctr = 3 And Attempt_DB_Connection(str) = False Then
        Text2.BackColor = vbRed
        Text2.Text = "LAST ATTEMPT"
    End If

    If Attempt_DB_Connection(str) = False Then
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
Text2.Text = "Attempt - " & ctr
End Sub
