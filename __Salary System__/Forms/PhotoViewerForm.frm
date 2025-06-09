VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Photo Upload Form"
   ClientHeight    =   4530
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9285
   Icon            =   "PhotoViewerForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   4530
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer Photo"
      Height          =   492
      Left            =   3720
      TabIndex        =   7
      Top             =   3960
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2160
      TabIndex        =   6
      Text            =   "*.jpg"
      Top             =   600
      Width           =   2772
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   2892
   End
   Begin VB.DirListBox Dir1 
      Height          =   2016
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1812
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label5 
      Caption         =   "Image Path"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   1212
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Height          =   372
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   8772
   End
   Begin VB.Label Label3 
      Caption         =   "Search file Type"
      Height          =   372
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   2292
   End
   Begin VB.Label Label2 
      Caption         =   "Select Folders"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "Select Drives"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1692
   End
   Begin VB.Image Image1 
      Height          =   3012
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3012
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form1.Label7.Caption = Label4.Caption
Form1.Image1.Picture = LoadPicture(Label4.Caption)
Unload Me
Form1.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Pattern = Text1.Text
End Sub

Private Sub Dir1_Click()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dir1.Path = Dir1.List(Dir1.ListIndex)
File1.Path = Dir1.Path
File1.SetFocus
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
If Right(File1.Path, 1) = "\" Then
Label4.Caption = File1.Path & File1.FileName
Image1.Picture = LoadPicture(Label4.Caption)
Else
Label4.Caption = File1.Path & "\" & File1.FileName
Image1.Picture = LoadPicture(Label4.Caption)
End If
Command1.SetFocus
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
File1_DblClick
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
Form1.Show
End If
End Sub

Private Sub Form_Load()
Drive1.Drive = App.Path
Dir1.Path = App.Path
File1.Pattern = Text1.Text
End Sub

