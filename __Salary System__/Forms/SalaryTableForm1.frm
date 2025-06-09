VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form22 
   Caption         =   "Salary Table Form"
   ClientHeight    =   9090
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   17415
   Icon            =   "SalaryTableForm1.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9090
   ScaleWidth      =   17415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Record Operation"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   14160
      TabIndex        =   81
      Top             =   5640
      Width           =   3015
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   84
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Update Salary Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   83
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Search Salary Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   82
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label38 
         Caption         =   "Select Payment Period"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   480
         Width           =   2415
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   8493
      _Version        =   393216
      Tab             =   2
      TabHeight       =   706
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial TUR"
         Size            =   11.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Salary Details"
      TabPicture(0)   =   "SalaryTableForm1.frx":1084A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(1)=   "Label16"
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(6)=   "Label27"
      Tab(0).Control(7)=   "Label32"
      Tab(0).Control(8)=   "Label33"
      Tab(0).Control(9)=   "Label34"
      Tab(0).Control(10)=   "Label76"
      Tab(0).Control(11)=   "Text14"
      Tab(0).Control(12)=   "Text15"
      Tab(0).Control(13)=   "Text16"
      Tab(0).Control(14)=   "Text17"
      Tab(0).Control(15)=   "Text18"
      Tab(0).Control(16)=   "Text19"
      Tab(0).Control(17)=   "Text26"
      Tab(0).Control(18)=   "Text31"
      Tab(0).Control(19)=   "Text32"
      Tab(0).Control(20)=   "Text33"
      Tab(0).Control(21)=   "Command4"
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Deductions"
      TabPicture(1)   =   "SalaryTableForm1.frx":10866
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(1)=   "Label22"
      Tab(1).Control(2)=   "Label23"
      Tab(1).Control(3)=   "Label25"
      Tab(1).Control(4)=   "Label26"
      Tab(1).Control(5)=   "Label28"
      Tab(1).Control(6)=   "Label30"
      Tab(1).Control(7)=   "Label24"
      Tab(1).Control(8)=   "Label36"
      Tab(1).Control(9)=   "Text20"
      Tab(1).Control(10)=   "Text21"
      Tab(1).Control(11)=   "Text22"
      Tab(1).Control(12)=   "Text24"
      Tab(1).Control(13)=   "Text25"
      Tab(1).Control(14)=   "Text27"
      Tab(1).Control(15)=   "Text28"
      Tab(1).Control(16)=   "Command5"
      Tab(1).Control(17)=   "Text23"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Net Salary"
      TabPicture(2)   =   "SalaryTableForm1.frx":10882
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label29"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label31"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label35"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label37"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text29"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text30"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Combo1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         TabIndex        =   80
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -66480
         TabIndex        =   76
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -64560
         TabIndex        =   73
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -64800
         TabIndex        =   72
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text33 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69600
         MaxLength       =   2
         TabIndex        =   68
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69600
         MaxLength       =   2
         TabIndex        =   67
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69600
         MaxLength       =   2
         TabIndex        =   66
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5160
         TabIndex        =   65
         Top             =   2400
         Width           =   8415
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5160
         TabIndex        =   64
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -64920
         TabIndex        =   63
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -66480
         TabIndex        =   62
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -65160
         TabIndex        =   61
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71400
         TabIndex        =   60
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71400
         TabIndex        =   59
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71400
         TabIndex        =   58
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71400
         TabIndex        =   57
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71400
         TabIndex        =   56
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72000
         TabIndex        =   55
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72000
         TabIndex        =   54
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72000
         TabIndex        =   53
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72000
         TabIndex        =   52
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72000
         TabIndex        =   51
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72000
         TabIndex        =   50
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7800
         TabIndex        =   35
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label37 
         Caption         =   "Payment Period"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   79
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -66720
         TabIndex        =   78
         Top             =   1560
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label24 
         Caption         =   "Working Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68280
         TabIndex        =   77
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   75
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -72240
         TabIndex        =   74
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label34 
         Caption         =   "% of Basic Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68640
         TabIndex        =   71
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label33 
         Caption         =   "% of Basic Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68640
         TabIndex        =   70
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label32 
         Caption         =   "% of Basic Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68640
         TabIndex        =   69
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label31 
         Caption         =   "Net Pay In Words"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "Leave Without Pay (Enter Leave Days)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68640
         TabIndex        =   34
         Top             =   840
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label29 
         Caption         =   "Net Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   33
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Total Deductions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66840
         TabIndex        =   32
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label27 
         Caption         =   "Gross Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66480
         TabIndex        =   31
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Loan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72240
         TabIndex        =   30
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Salary Advance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   29
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "Professional Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73440
         TabIndex        =   28
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "ESI Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72840
         TabIndex        =   27
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Provident Fund"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Performance Bonus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74280
         TabIndex        =   25
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Medical Allowance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   24
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label18 
         Caption         =   "Travel Allowance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   23
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "House Rent Allowance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   22
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "Dearness Allowance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   21
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Basic Pay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Employee Data"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   17175
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11640
         TabIndex        =   49
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11640
         TabIndex        =   48
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11640
         TabIndex        =   47
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   46
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   45
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   44
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   43
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   42
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   41
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   40
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   39
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   38
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   6840
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search Employee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8640
         MouseIcon       =   "SalaryTableForm1.frx":1089E
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "IFSC Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Bank A/C No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "ESI No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "UAN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "PF No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Work Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Deparment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   2415
         Left            =   14760
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "EmpPhoto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   15240
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "EmpName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "EmpCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "DOJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "DOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   2880
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Record Operation"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   14520
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Con As ADODB.Connection
Dim rs, RS_Sal As ADODB.Recordset
Dim sql As String
Dim empFound As Boolean
Dim fempcode As String
Dim EmpPhoto As String
Dim MandFields() As Object

Private Function Add_MandatoryField(Field As Object)
'To Do

End Function

Private Function Get_CurrentPeriod() As String

Dim mIndex(12) As Integer
mIndex(0) = 1
mIndex(1) = 2
mIndex(2) = 3
mIndex(3) = 4
mIndex(4) = 5
mIndex(5) = 6
mIndex(6) = 7
mIndex(7) = 8
mIndex(8) = 9
mIndex(9) = 10
mIndex(10) = 11
mIndex(11) = 12

Dim mName(12) As String
mName(0) = "January"
mName(1) = "February"
mName(2) = "March"
mName(3) = "April"
mName(4) = "May"
mName(5) = "June"
mName(6) = "July"
mName(7) = "August"
mName(8) = "September"
mName(9) = "October"
mName(10) = "November"
mName(11) = "December"


Dim current_year, current_month As String

current_year = Year(Date)
current_month = ""

'Converting the month number to month name
For i = 0 To 11
If Month(Date) = mIndex(i) Then
current_month = mName(i)
Exit For
End If
Next

Get_CurrentPeriod = current_month & " " & current_year

End Function


Private Function AddRecord()

RS_Sal.AddNew
'Employee Information
RS_Sal(0) = Text1.Text
RS_Sal(1) = Text2.Text
RS_Sal(2) = EmpPhoto
RS_Sal(3) = Text5.Text
RS_Sal(4) = Text6.Text
RS_Sal(5) = Text3.Text
RS_Sal(6) = Text7.Text
RS_Sal(7) = Text8.Text
RS_Sal(8) = Text9.Text
RS_Sal(9) = Text10.Text
RS_Sal(10) = Text11.Text
RS_Sal(11) = Text12.Text
RS_Sal(12) = Text13.Text
'Salary Information
RS_Sal(13) = Text14.Text
RS_Sal(14) = Text15.Text
RS_Sal(15) = Text16.Text
RS_Sal(16) = Text17.Text
RS_Sal(17) = Text18.Text
RS_Sal(18) = Text19.Text
RS_Sal(19) = Text26.Text
RS_Sal(20) = Text20.Text
RS_Sal(21) = Text21.Text
RS_Sal(22) = Text22.Text
RS_Sal(23) = Text24.Text
RS_Sal(24) = Text25.Text
RS_Sal(25) = Text27.Text
RS_Sal(26) = Text28.Text
RS_Sal(27) = Text29.Text
RS_Sal(28) = Text30.Text
RS_Sal(29) = Combo1.Text
RS_Sal.Update

MsgBox "Record Updated Successfully !!"

End Function

Private Function Display_Found_Employee_Record()

'EMPLOYEE DATA

Text2.Text = rs.fields(1)
Text3.Text = rs.fields(26)
If Not IsNull(rs.fields(27)) Then
Text4.Text = rs.fields(27)
End If
Text5.Text = rs.fields(24)
Text6.Text = rs.fields(25)
Text7.Text = rs.fields(39)
If Not IsNull(rs.fields(29)) Then
Text8.Text = rs.fields(29)
Else
Text20.Text = 0
Text20.Enabled = False
Text20.BackColor = &H8000000B
End If
If Not IsNull(rs.fields(30)) Then
Text9.Text = rs.fields(30)
End If
If Not IsNull(rs.fields(31)) Then
Text10.Text = rs.fields(31)
Else
Text21.Text = 0
Text21.Enabled = False
Text21.BackColor = &H8000000B
End If
Text11.Text = rs.fields(35)
Text12.Text = rs.fields(34)
Text13.Text = rs.fields(36)
If Not IsNull(rs.fields(20)) Then
EmpPhoto = rs.fields(20)
Image1.Picture = LoadPicture(rs.fields(20))
End If

End Function

Private Function Is_Emp_Paid_For_Period() As Boolean

Dim returnFlag As Boolean

Dim empID As String
empID = Text1.Text

Dim PayPeriod As String
PayPeriod = Combo1.Text

If RS_Sal.RecordCount = 0 Then
returnFlag = False
Else
RS_Sal.MoveFirst
Do Until RS_Sal.EOF
If RS_Sal(0) = empID Then
'Checking if employee is paid for current payment period
If RS_Sal(29) = PayPeriod Then
'Exiting Search once payment period is found
returnFlag = True
Exit Do
End If
End If
RS_Sal.MoveNext
Loop
End If

Is_Emp_Paid_For_Period = returnFlag

End Function


Private Function ClearAllEntries()

For Each Ctrl In Me.Controls
If TypeOf Ctrl Is TextBox Then
Ctrl.Text = ""
End If
Next
Combo1.Text = ""
Image1.Picture = LoadPicture("")

End Function

Private Sub Command1_Click()

If Text4.Text <> "" Then
MsgBox "Cannot Pay Salary To A Resigned Employee !!"
Exit Sub
End If

If Text30.Text = "" Then
MsgBox "Net Pay In Words Cannot Be Left Empty !!"
Exit Sub
End If

If Is_Emp_Paid_For_Period = True Then
MsgBox "Cannot Save Payment Data !!" & vbNewLine & "The Employee Is Already Paid For Period " & Combo1.Text
Exit Sub
End If

'Check for Combo1(Pay Period) entries other than already given one
For i = 0 To 11
If Combo1.Text <> Combo1.List(i) Then
MsgBox "Only Select From Given Pay Period !!" & vbNewLine & "Manual Entry Is Not Allowed !!"
Combo1.SetFocus
Exit For
End If
Next

Call AddRecord
Call ClearAllEntries

End Sub

Private Sub Command2_Click()

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Image1.Picture = LoadPicture("")
Text20.Text = ""
Text20.Enabled = True
Text20.BackColor = &H80000005
Text21.Text = ""
Text21.Enabled = True
Text21.BackColor = &H80000005
SSTab1.Enabled = False

fempcode = Text1.Text

If rs.RecordCount = 0 Then
empFound = False
Else

rs.MoveFirst
Do Until rs.EOF

If rs.fields(0) = fempcode Then
empFound = True
SSTab1.Enabled = True
Call Display_Found_Employee_Record
Exit Sub

Else
rs.MoveNext
empFound = False
End If

Loop
End If

If empFound = False Then
MsgBox "No record with the EmpCode found !!"
End If

End Sub

Private Sub Command3_Click()

If Text26.Text = "" Or Text28.Text = "" Then
MsgBox "Cannot Generate Net Pay Without Having Gross Pay & Total Deductions Calculated !!"
Exit Sub
End If

Text29.Text = Val(Text26.Text) - Val(Text28.Text)

End Sub

Private Sub Command4_Click()

If Text14.Text = "" Then
MsgBox "Basic Pay cannot be left empty !!"
Exit Sub
End If

Call Set_Earnings_Empty_Fields_To_Zero

Text15.Text = (Val(Text31.Text) / 100) * Val(Text14.Text)
Text16.Text = (Val(Text32.Text) / 100) * Val(Text14.Text)
Text17.Text = (Val(Text33.Text) / 100) * Val(Text14.Text)

'Generating Gross Pay and displaying on UI
Text26.Text = Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text)

End Sub

Private Sub Command5_Click()
'Getting P Tax and displaying on Text Box
Text22.Text = Generate_ProfessionalTax()

Call Set_Deductions_Empty_Fields_To_Zero

'Setting Working Days as Mandatory when leave is given
If Text27.Text <> "" Then
'If working days is empty when leave days is not
If Text23.Text = "" Then
MsgBox "Working Days Cannot Be Left Empty If Leave Is Given !!"
Text23.SetFocus
Exit Sub
Else
Dim leaveDeduction As Integer
leaveDeduction = Calc_LeaveDeduction
End If
End If

'Generating Total Deduction
Text28.Text = Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text24.Text) + Val(Text25.Text) + leaveDeduction

End Sub

Private Function Generate_ProfessionalTax() As Integer

'Generating Professional Tax for West Bengal

Dim gross_pay As Double
Dim ptax As Integer

gross_pay = Val(Text26.Text)

'Generate Professional Tax based on West Bengal slab
If gross_pay >= 0 And gross_pay <= 10000 Then
ptax = 0
ElseIf gross_pay > 10000 And gross_pay <= 15000 Then
ptax = 110
ElseIf gross_pay > 15000 And gross_pay <= 25000 Then
ptax = 130
ElseIf gross_pay > 25000 And gross_pay <= 40000 Then
ptax = 150
Else
ptax = 200
End If

'Return P Tax amount
Generate_ProfessionalTax = ptax

End Function

Private Function Calc_LeaveDeduction() As Double

'Setting Leave Deduction to 0 when TextBox is left empty
If Text27.Text = "" Then
Text27.Text = 0
Calc_LeaveDeduction = 0
Exit Function
End If

Dim leave_days, working_days As Integer
Dim gross_pay As Double

leave_days = Val(Text27.Text)
working_days = Val(Text23.Text)
gross_pay = Val(Text26.Text)

pay_per_day = gross_pay / working_days

total_leave_deduction = leave_days * pay_per_day

Calc_LeaveDeduction = total_leave_deduction

End Function

Private Sub Command6_Click()
Unload Me
End
'Call Is_Emp_Paid_For_Period("101")
End Sub


Private Function Display_SalaryTable_Data()

'EARNINGS
'Basic Salary
Text14.Text = RS_Sal.fields(13)
'DA
Text15.Text = RS_Sal.fields(14)
'HRA
Text16.Text = RS_Sal.fields(15)
'TA
Text17.Text = RS_Sal.fields(16)
'Medical Allowance
Text18.Text = RS_Sal.fields(17)
'Perfomance Bonus
Text19.Text = RS_Sal.fields(18)
'Gross Pay
Text26.Text = RS_Sal.fields(19)

'DEDUCTIONS
'PF
Text20.Text = RS_Sal.fields(20)
'ESI
Text21.Text = RS_Sal.fields(21)
'P.Tax
Text22.Text = RS_Sal.fields(22)
'Salary Advance
Text24.Text = RS_Sal.fields(23)
'Loan
Text25.Text = RS_Sal.fields(24)
'Leave Without Pay
Text27.Text = RS_Sal.fields(25)
'Total Deductions
Text28.Text = RS_Sal.fields(26)

'NET SALARY
'Net Pay
Text29.Text = RS_Sal.fields(27)
'Net Pay In Words
Text30.Text = RS_Sal.fields(28)
'Payment Period
Combo1.Text = RS_Sal.fields(29)

End Function


Private Function Search_Salary_Table()

Dim RecordFound As Boolean

If RS_Sal.RecordCount = 0 Then
MsgBox "Salary Table Is Empty !! "
RecordFound = False
GoTo FunctionStop
End If

Dim PayPeriod As String
PayPeriod = Combo2.Text

If PayPeriod <> "" Then
RS_Sal.MoveFirst
Do Until RS_Sal.EOF
If RS_Sal(0) = Text1.Text Then
'Checking if employee is paid for given payment period
If RS_Sal(29) = PayPeriod Then
'Setting Cursor to Found Row And Exiting Search once payment period is found
RecordFound = True
Exit Do
GoTo FunctionStop
Else
End If
End If
RecordFound = False
RS_Sal.MoveNext
Loop
End If

FunctionStop:
Search_Salary_Table = RecordFound
Exit Function

End Function


Private Sub Command7_Click()

'If Search_Salary_Table = False Then
'MsgBox "Employee Record does not exist in Salary Table" & vbNewLine & "Update is allowed for existing record"
'Exit Sub
'End If

If Search_Salary_Table = True Then
'Enabling Update Salary Table Button
Command8.Enabled = True
'Displaying Employee Data from Employee Table
Command2_Click
Call Display_SalaryTable_Data
Else
MsgBox "Salary Table Record With Given Employee ID And Given Payment Period Not Found !!"
End If

End Sub


Private Function Set_Earnings_Empty_Fields_To_Zero()
'Setting Medical Allowance to 0 when Empty
If Text18.Text = "" Then
Text18.Text = 0
End If
'Setting Performance Bonus to 0 when Empty
If Text19.Text = "" Then
Text19.Text = 0
End If

End Function


Private Function Set_Deductions_Empty_Fields_To_Zero()
'Setting PF to 0 when Empty
If Text20.Text = "" Then
Text20.Text = 0
End If
'Setting ESI Amount to 0 when Empty
If Text21.Text = "" Then
Text21.Text = 0
End If
'Setting Salary Advance to 0 when Empty
If Text24.Text = "" Then
Text24.Text = 0
End If
'Setting Loan to 0 when Empty
If Text25.Text = "" Then
Text25.Text = 0
End If
End Function


Private Sub Command8_Click()

Call Set_Earnings_Empty_Fields_To_Zero
Call Set_Deductions_Empty_Fields_To_Zero

Call AddRecord
Call ClearAllEntries

End Sub


Private Sub Form_Load()
'Set Con = New ADODB.Connection
'Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Infosys.mdb;Persist Security Info=False"
'Con.Open
'Conn = GetDbConn
Call OpenDbConn

Set rs = New ADODB.Recordset
rs.Open "Select * from Employee", gCon, adOpenStatic, adLockReadOnly

Set RS_Sal = New ADODB.Recordset
RS_Sal.CursorLocation = adUseClient
RS_Sal.Open "Salary", gCon, adOpenDynamic, adLockOptimistic, adCmdTable

SSTab1.Enabled = False
SSTab1.Tab = 0

Dim currentYear As Integer
currentYear = Year(Date)
With Combo1
.AddItem "January " & currentYear
.AddItem "February " & currentYear
.AddItem "March " & currentYear
.AddItem "April " & currentYear
.AddItem "May " & currentYear
.AddItem "June " & currentYear
.AddItem "July " & currentYear
.AddItem "August " & currentYear
.AddItem "September " & currentYear
.AddItem "October " & currentYear
.AddItem "November " & currentYear
.AddItem "December " & currentYear
End With

With Combo2
For i = 0 To 11
.AddItem Combo1.List(i)
Next
End With

'Disabling Update Salary Table Button
Command8.Enabled = False

End Sub

Private Sub Text27_LostFocus()
If Val(Text27.Text) <> 0 Then
Call Add_MandatoryField(Text27)
Text23.Visible = True
Label24.Visible = True
Label36.Visible = True
Text23.SetFocus
Else
Text23.Visible = False
Label24.Visible = False
Label36.Visible = False
End If
End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub Text32_Validate(Cancel As Boolean)
If Not (Val(Text32.Text) >= 0 And Val(Text32.Text) <= 100) Then
MsgBox "Percentage cannot be more than 100 or less than 0 !!"
Cancel = True
End If
End Sub

Private Sub Text33_Validate(Cancel As Boolean)
If Not (Val(Text33.Text) >= 0 And Val(Text33.Text) <= 100) Then
MsgBox "Percentage cannot be more than 100 or less than 0 !!"
Cancel = True
End If
End Sub

Private Sub Text31_Validate(Cancel As Boolean)
If Not (Val(Text31.Text) >= 0 And Val(Text31.Text) <= 100) Then
MsgBox "Percentage cannot be more than 100 or less than 0 !!"
Cancel = True
End If
End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub











