VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   6048
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   12912
   LinkTopic       =   "Form5"
   ScaleHeight     =   6048
   ScaleWidth      =   12912
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2292
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11892
      _ExtentX        =   20976
      _ExtentY        =   4043
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "&EPF Details"
      TabPicture(0)   =   "Form5.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&PAN Details"
      TabPicture(1)   =   "Form5.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form5.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   612
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   2772
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
