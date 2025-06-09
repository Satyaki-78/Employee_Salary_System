VERSION 5.00
Begin VB.Form Form12 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   Caption         =   "Form12"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   LinkTopic       =   "Form12"
   ScaleHeight     =   2400
   ScaleWidth      =   5805
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
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter Password below"
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
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim pass As String
Dim m As Integer

Private Sub Command1_Click()
Dim str As String

pass = "123"

str = Text1.Text

Do

    If ctr = 3 Then
        If str <> pass Then
            m = MsgBox("ATTEMPTS EXCEEDED" & vbNewLine & vbNewLine & "ACCOUNT LOCKED", 16)
            End
            Exit Sub
        Else
            m = MsgBox("Authorized User")
            Unload Me
            Form12.Show
            Exit Sub
        End If
    End If
    
    If str = pass And ctr <= 3 Then
        m = MsgBox("Authorized User")
        Unload Me
        Form12.Show
        Exit Sub
    End If

ctr = ctr + 1
Text2.Text = "Attempt - " & ctr
    
    If ctr = 3 And str <> pass Then
        Text2.BackColor = vbRed
        Text2.Text = "LAST ATTEMPT"
    End If

    If str <> pass Then
        m = MsgBox("INVALID PASSWORD", 16)
        Text1.SetFocus
        Exit Do
    End If
    
Loop Until ctr > 3

End Sub

Private Sub Form_Load()
ctr = 1
Text2.Text = "Attempt - " & ctr
End Sub
