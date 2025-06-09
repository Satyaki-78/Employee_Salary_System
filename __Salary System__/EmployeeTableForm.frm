VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Employee Table Form"
   ClientHeight    =   9420
   ClientLeft      =   -5160
   ClientTop       =   -1560
   ClientWidth     =   18060
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EmployeeTableForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   18060
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   132
      Top             =   6360
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command7 
      Caption         =   "P&rev Tab"
      BeginProperty Font 
         Name            =   "Candara Light"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12720
      TabIndex        =   121
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next &Tab"
      BeginProperty Font 
         Name            =   "Candara Light"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   14280
      TabIndex        =   120
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "NOTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   53
      Top             =   5400
      Width           =   5532
      Begin VB.Label Label45 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Indicates atleast one of the marked fields is mandatory"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   5172
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label43 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label42 
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
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Indicates fields are always mandatory"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   5172
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Record Operation"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   15960
      TabIndex        =   51
      Top             =   960
      Width           =   1935
      Begin VB.CommandButton Command5 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   240
         TabIndex        =   50
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   240
         TabIndex        =   49
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S&earch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
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
         Left            =   240
         TabIndex        =   47
         Top             =   480
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   58
      Top             =   120
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   706
      TabMaxWidth     =   18
      WordWrap        =   0   'False
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial TUR"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Personal Info"
      TabPicture(0)   =   "EmployeeTableForm.frx":1084A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label58"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label57"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label56"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label55"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label38"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label32"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label29"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label28"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label27"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label25"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label19"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text28"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text27"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text21"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text8"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text7"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Combo9"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "C&ontact Info"
      TabPicture(1)   =   "EmployeeTableForm.frx":10866
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label20"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label31"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text17"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text16"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Identification Info"
      TabPicture(2)   =   "EmployeeTableForm.frx":10882
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text35"
      Tab(2).Control(1)=   "Text31"
      Tab(2).Control(2)=   "Text30"
      Tab(2).Control(3)=   "Text24"
      Tab(2).Control(4)=   "Text25"
      Tab(2).Control(5)=   "Text26"
      Tab(2).Control(6)=   "Label62"
      Tab(2).Control(7)=   "Label54"
      Tab(2).Control(8)=   "Label53"
      Tab(2).Control(9)=   "Label39"
      Tab(2).Control(10)=   "Label44"
      Tab(2).Control(11)=   "Label46"
      Tab(2).Control(12)=   "Label51"
      Tab(2).Control(13)=   "Label52"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "&Qualification Info"
      TabPicture(3)   =   "EmployeeTableForm.frx":1089E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text22"
      Tab(3).Control(1)=   "Text23"
      Tab(3).Control(2)=   "Label47"
      Tab(3).Control(3)=   "Label48"
      Tab(3).Control(4)=   "Label49"
      Tab(3).Control(5)=   "Label50"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "P&hoto Upload"
      TabPicture(4)   =   "EmployeeTableForm.frx":108BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command6"
      Tab(4).Control(1)=   "Label6"
      Tab(4).Control(2)=   "Image1"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "O&fficial Info"
      TabPicture(5)   =   "EmployeeTableForm.frx":108D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Bank Details"
      TabPicture(6)   =   "EmployeeTableForm.frx":108F2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Text34"
      Tab(6).Control(1)=   "Text33"
      Tab(6).Control(2)=   "Text32"
      Tab(6).Control(3)=   "Label66"
      Tab(6).Control(4)=   "Label65"
      Tab(6).Control(5)=   "Label64"
      Tab(6).Control(6)=   "Label61"
      Tab(6).Control(7)=   "Label60"
      Tab(6).Control(8)=   "Label59"
      Tab(6).ControlCount=   9
      TabCaption(7)   =   "&Duty Details"
      TabPicture(7)   =   "EmployeeTableForm.frx":1090E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label70"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label71"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Label72"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Label73"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Label75"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "Label76"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Label77"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "Label78"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "Combo7"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "Text36"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "Combo8"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "Text38"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).ControlCount=   12
      Begin VB.ComboBox Combo9 
         Height          =   360
         Left            =   -72960
         TabIndex        =   8
         Tag             =   "m"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text38 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -71760
         TabIndex        =   45
         Tag             =   "m"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox Combo8 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -66600
         TabIndex        =   46
         Tag             =   "m"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Text36 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -66600
         TabIndex        =   44
         Tag             =   "m"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.ComboBox Combo7 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -71760
         TabIndex        =   43
         Tag             =   "m"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   -74880
         TabIndex        =   98
         Top             =   480
         Width           =   15375
         Begin VB.ComboBox Combo6 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2400
            TabIndex        =   25
            Tag             =   "m"
            Top             =   960
            Width           =   2535
         End
         Begin VB.ComboBox Combo5 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2400
            TabIndex        =   26
            Tag             =   "m"
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   3360
            MaxLength       =   4
            TabIndex        =   34
            Tag             =   "o"
            Top             =   3960
            Width           =   612
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   33
            Tag             =   "o"
            Top             =   3960
            Width           =   372
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   32
            Tag             =   "o"
            Top             =   3960
            Width           =   372
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   3360
            MaxLength       =   4
            TabIndex        =   31
            Tag             =   "m"
            Top             =   3360
            Width           =   612
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   30
            Tag             =   "m"
            Top             =   3360
            Width           =   372
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   29
            Tag             =   "m"
            Top             =   3360
            Width           =   372
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2400
            TabIndex        =   28
            Tag             =   "m"
            Top             =   2760
            Width           =   2892
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2400
            TabIndex        =   27
            Tag             =   "m"
            Top             =   2160
            Width           =   2895
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   6480
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   99
            Tag             =   "o"
            Top             =   3840
            Width           =   5055
         End
         Begin VB.ComboBox Combo3 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   8880
            TabIndex        =   38
            Top             =   2640
            Width           =   2172
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Fresher"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   23
            Tag             =   "m"
            Top             =   360
            Width           =   972
         End
         Begin VB.CheckBox Check2 
            Caption         =   "E&xperienced"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3480
            TabIndex        =   24
            Tag             =   "m"
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox Combo4 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   8880
            TabIndex        =   39
            Top             =   3120
            Width           =   2172
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8280
            MaxLength       =   255
            TabIndex        =   36
            Tag             =   "o"
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8280
            MaxLength       =   255
            TabIndex        =   35
            Tag             =   "o"
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8280
            MaxLength       =   255
            TabIndex        =   37
            Tag             =   "o"
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label Label69 
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
            Left            =   2040
            TabIndex        =   123
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label68 
            Caption         =   "Employment Type"
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
            Left            =   120
            TabIndex        =   122
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label67 
            Caption         =   "Employee Status"
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
            Left            =   120
            TabIndex        =   119
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label63 
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
            Left            =   1920
            TabIndex        =   118
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label15 
            Caption         =   " dd       mm       yyyy"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   117
            Top             =   4200
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   " dd       mm       yyyy"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   116
            Top             =   3600
            Width           =   1575
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
            Left            =   1200
            TabIndex        =   115
            Top             =   3960
            Width           =   495
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
            Left            =   1200
            TabIndex        =   114
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Department"
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
            Left            =   480
            TabIndex        =   113
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label8 
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
            Left            =   480
            TabIndex        =   112
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Reason of Resignation"
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
            Left            =   5040
            TabIndex        =   111
            Top             =   3840
            Width           =   1335
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label16 
            Caption         =   "Emp Category"
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
            Left            =   360
            TabIndex        =   110
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "EPF_Status"
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
            Left            =   7560
            TabIndex        =   109
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "ESI_Status"
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
            Left            =   7560
            TabIndex        =   108
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label22 
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
            Left            =   7200
            TabIndex        =   107
            Top             =   1080
            Width           =   855
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label23 
            Caption         =   "EPF No"
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
            Left            =   7200
            TabIndex        =   106
            Top             =   480
            Width           =   855
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label24 
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
            Left            =   7200
            TabIndex        =   105
            Top             =   1920
            Width           =   735
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label34 
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
            Left            =   1680
            TabIndex        =   104
            Top             =   3360
            Width           =   135
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
            Left            =   1800
            TabIndex        =   103
            Top             =   2760
            Width           =   135
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
            Left            =   1800
            TabIndex        =   102
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label37 
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
            Left            =   1920
            TabIndex        =   101
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label40 
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
            Left            =   8040
            TabIndex        =   100
            Top             =   1080
            Width           =   135
         End
      End
      Begin VB.TextBox Text35 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -67320
         MaxLength       =   11
         TabIndex        =   19
         Tag             =   "o"
         Top             =   1080
         Width           =   2412
      End
      Begin VB.TextBox Text34 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71160
         MaxLength       =   11
         TabIndex        =   42
         Tag             =   "m"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox Text33 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71160
         MaxLength       =   20
         TabIndex        =   41
         Tag             =   "m"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71160
         TabIndex        =   40
         Tag             =   "m"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "o"
         Top             =   3960
         Width           =   2412
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71760
         MaxLength       =   16
         TabIndex        =   17
         Tag             =   "o"
         Top             =   3240
         Width           =   2412
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Upload"
         Height          =   492
         Left            =   -68880
         TabIndex        =   22
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71040
         TabIndex        =   20
         Tag             =   "m"
         Top             =   1320
         Width           =   6375
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71040
         TabIndex        =   21
         Tag             =   "m"
         Top             =   2040
         Width           =   6375
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "m"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   15
         Tag             =   "m"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71760
         MaxLength       =   8
         TabIndex        =   16
         Tag             =   "o"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "o"
         Top             =   1080
         Width           =   1572
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "o"
         Top             =   1800
         Width           =   1572
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Tag             =   "o"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   -72960
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "m"
         Top             =   720
         Width           =   1452
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -72960
         MaxLength       =   255
         TabIndex        =   2
         Tag             =   "m"
         Top             =   1920
         Width           =   9375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   -72960
         TabIndex        =   1
         Tag             =   "m"
         Top             =   1320
         Width           =   3132
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72960
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "m"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "m"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72000
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "m"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   -69480
         TabIndex        =   10
         Tag             =   "m"
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   -72960
         MaxLength       =   255
         TabIndex        =   3
         Tag             =   "o"
         Top             =   2400
         Width           =   9375
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   -72960
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "m"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   -69480
         TabIndex        =   9
         Tag             =   "m"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label31 
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
         Left            =   2880
         TabIndex        =   133
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label78 
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
         Left            =   -72120
         TabIndex        =   131
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label77 
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72960
         TabIndex        =   130
         Top             =   2640
         Width           =   855
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
         Left            =   -66840
         TabIndex        =   129
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label75 
         Caption         =   "Workplace Status"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   128
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label73 
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
         Left            =   -67200
         TabIndex        =   127
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label72 
         Caption         =   "Shift Time"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   126
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label71 
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
         Left            =   -72120
         TabIndex        =   125
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label70 
         Caption         =   "Duty Shift"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   124
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label62 
         Caption         =   "TIN No"
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
         Left            =   -68520
         TabIndex        =   97
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label66 
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
         Left            =   -71760
         TabIndex        =   96
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label65 
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
         Left            =   -71760
         TabIndex        =   95
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label64 
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
         Left            =   -71760
         TabIndex        =   94
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label61 
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
         Left            =   -72960
         TabIndex        =   93
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label60 
         Caption         =   "Account No"
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
         Left            =   -72960
         TabIndex        =   92
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label59 
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
         Left            =   -72960
         TabIndex        =   91
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label54 
         Caption         =   "Voter ID"
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
         Left            =   -73080
         TabIndex        =   90
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label53 
         Caption         =   "Driving License  No"
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
         TabIndex        =   89
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Emp Photo"
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
         Left            =   -70320
         TabIndex        =   88
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3495
         Left            =   -69840
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label47 
         Caption         =   "Educational Qualification"
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
         Left            =   -72840
         TabIndex        =   87
         Top             =   1320
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label48 
         Caption         =   "Technical Qualification"
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
         Left            =   -72840
         TabIndex        =   86
         Top             =   2040
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label49 
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
         Left            =   -71520
         TabIndex        =   85
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label50 
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
         Left            =   -71520
         TabIndex        =   84
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label39 
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
         Height          =   252
         Left            =   -72480
         TabIndex        =   83
         Top             =   1080
         Width           =   132
      End
      Begin VB.Label Label44 
         Caption         =   "PAN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -73320
         TabIndex        =   82
         Top             =   1080
         Width           =   852
      End
      Begin VB.Label Label46 
         Caption         =   "Aadhaar No"
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
         TabIndex        =   81
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label51 
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
         Left            =   -72480
         TabIndex        =   80
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label52 
         Caption         =   "Passport No"
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
         TabIndex        =   79
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Mobile No"
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
         Left            =   1680
         TabIndex        =   78
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Alt Mobile No"
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
         Left            =   1320
         TabIndex        =   77
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Email ID"
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
         Left            =   1800
         TabIndex        =   76
         Top             =   2520
         Width           =   975
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
         Left            =   -74400
         TabIndex        =   75
         Top             =   720
         Width           =   1095
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
         Left            =   -74520
         TabIndex        =   74
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Address"
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
         TabIndex        =   73
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "DOB"
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
         Left            =   -73920
         TabIndex        =   72
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   " dd      mm      yyyy"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   71
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Gender"
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
         TabIndex        =   70
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label25 
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
         Left            =   -73320
         TabIndex        =   69
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label27 
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
         Left            =   -73320
         TabIndex        =   67
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label28 
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
         Left            =   -73320
         TabIndex        =   66
         Top             =   4440
         Width           =   135
      End
      Begin VB.Label Label29 
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
         Left            =   -73320
         TabIndex        =   65
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label32 
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
         Left            =   -69840
         TabIndex        =   64
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label Label38 
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
         Left            =   -73320
         TabIndex        =   63
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label55 
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
         Left            =   -73320
         TabIndex        =   62
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label56 
         Caption         =   "PIN Code"
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
         Left            =   -73800
         TabIndex        =   61
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label57 
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
         Left            =   -69840
         TabIndex        =   60
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label58 
         Caption         =   "Caste"
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
         Left            =   -70560
         TabIndex        =   59
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Religion"
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
         Left            =   -70800
         TabIndex        =   68
         Top             =   3720
         Width           =   855
      End
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15480
      TabIndex        =   52
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON1 As ADODB.Connection
Dim rs, RS1, rs2 As ADODB.Recordset
Dim rsDesig, rsDept As ADODB.Recordset
Dim sql As String
Dim DOB, DOJ, DOR As Date
Dim sdob, sdoj, sdor As String
Dim dobEmpty, dojEmpty, dorEmpty As Boolean
Dim fempcode, fempcode1 As String
Dim str As Integer
Dim EmpCategory As String
Dim epfFormat As String


Public Function AddRec()
rs.AddNew
rs.fields(0) = Text1.Text
rs.fields(1) = Text2.Text
rs.fields(2) = Text3.Text
rs.fields(4) = Text28.Text
rs.fields(5) = sdob
rs.fields(6) = Combo9.Text
rs.fields(7) = Text29.Text
rs.fields(8) = Text21.Text
rs.fields(12) = Text24.Text
rs.fields(13) = Text25.Text
rs.fields(18) = Text22.Text
rs.fields(19) = Text23.Text
rs.fields(20) = Label7.Caption
rs.fields(21) = EmpCategory
rs.fields(22) = Combo6.Text
rs.fields(23) = Combo5.Text
rs.fields(24) = Combo1.Text
rs.fields(25) = Combo2.Text
rs.fields(26) = sdoj
rs.fields(32) = Combo3.Text
rs.fields(33) = Combo4.Text
rs.fields(34) = Text32.Text
rs.fields(35) = Text33.Text
rs.fields(36) = Text34.Text
rs.fields(37) = Combo7.Text
rs.fields(38) = Text36.Text
rs.fields(39) = Text38.Text
rs.fields(40) = Combo8.Text
sdor = ""
Call Optional_EmptyFields_Update
rs.Update

For Each Ctrl In Me.Controls
If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is ComboBox Then
Ctrl.Text = ""
End If
Next
Check1.Value = 0
Check2.Value = 0
Label7.Caption = ""
Image1.Picture = LoadPicture("")
SSTab1.Tab = 0
Text1.SetFocus

MsgBox "Record Saved Successfully !!"
Call UpdateUserActivity("ADD", "Employee", Text1.Text)

'rs.Refresh
'RS1.Open "Select * from Employee", con, adOpenKeyset, adLockPessimistic
End Function


Public Function Disable()
For Each Ctrl In Me.Controls
If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is ComboBox Then
If Ctrl.Text <> "" Then
Ctrl.Enabled = False
Else
Ctrl.Enabled = True
End If
End If
Next
Check1.Enabled = False
Check2.Enabled = False
End Function


Private Function EmptyAllFields()
For Each Ctrl In Me.Controls
If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is ComboBox Then
If Ctrl.Text <> "" Then
Ctrl.Text = ""
End If
End If
Next
Image1.Picture = LoadPicture("")
Check1.Value = 0
Check2.Value = 0
End Function


Public Function ProperDateCheck(Field_Name As String, Txt1 As TextBox, Txt2 As TextBox, Txt3 As TextBox) As Boolean

Dim mDays(1, 10) As Integer
mDays(0, 0) = 1
mDays(0, 1) = 3
mDays(0, 2) = 4
mDays(0, 3) = 5
mDays(0, 4) = 6
mDays(0, 5) = 7
mDays(0, 6) = 8
mDays(0, 7) = 9
mDays(0, 8) = 10
mDays(0, 9) = 11
mDays(0, 10) = 12
mDays(1, 0) = 31
mDays(1, 1) = 31
mDays(1, 2) = 30
mDays(1, 3) = 31
mDays(1, 4) = 30
mDays(1, 5) = 31
mDays(1, 6) = 31
mDays(1, 7) = 30
mDays(1, 8) = 31
mDays(1, 9) = 30
mDays(1, 10) = 31
Dim dName(11) As String
dName(0) = "January"
dName(1) = "March"
dName(2) = "April"
dName(3) = "May"
dName(4) = "June"
dName(5) = "July"
dName(6) = "August"
dName(7) = "September"
dName(8) = "October"
dName(9) = "November"
dName(10) = "December"

Dim d, m, Y As Integer
Dim dTxtCtrl, mTxtCtrl, yTxtCtrl As TextBox
Dim flag As Boolean
flag = True

d = Val(Txt1.Text)
m = Val(Txt2.Text)
Y = Val(Txt3.Text)
Set dTxtCtrl = Txt1
Set mTxtCtrl = Txt2
Set yTxtCtrl = Txt3

If d = 0 And m = 0 And Y = 0 Then
flag = False
GoTo FuncTerminate
End If
'1,3,5,7,8,10,12 -- 31 days
'4,6,9,11 -- 30 days
'2 -- 28 days
Dim i, j As Integer
Dim isLeap As Boolean
Dim rm, rm1, rm2 As Integer

rm = Y / 100
rm1 = Y Mod 4
rm2 = Y Mod 400
If rm = Fix(rm) Then
If rm2 = 0 Then
isLeap = True
Else
isLeap = False
End If
Else
If rm1 = 0 Then
isLeap = True
Else
isLeap = False
End If
End If

If Not (m > 0 And m <= 12) Then
flag = False
MsgBox "Invalid entry for " & Field_Name & vbNewLine & "Can't have more than 12 months"
mTxtCtrl.SetFocus
GoTo FuncTerminate
End If

If m = 2 Then
'Non-Leap Year Fbruary No Of Days Check
If isLeap = False Then
If Not (d > 0 And d <= 28) Then
flag = False
MsgBox "Invalid entry for " & Field_Name & vbNewLine & "No of days must be between 1 and 28 for the month 2 February"
dTxtCtrl.SetFocus
GoTo FuncTerminate
End If
'Leap Year February No Of Days Check
Else
If Not (d > 0 And d <= 29) Then
flag = False
MsgBox "Invalid entry for " & Field_Name & vbNewLine & "No of days must be between 1 and 29 for the month 2 February"
dTxtCtrl.SetFocus
GoTo FuncTerminate
End If
End If

Else
For i = 0 To 10
If m = mDays(0, i) Then
If Not (d > 0 And d <= mDays(1, i)) Then
flag = False
MsgBox "Invalid entry for " & Field_Name & vbNewLine & "No of days must be between 1 and " & mDays(1, i) & " for the month " & mDays(0, i) & " " & dName(i)
dTxtCtrl.SetFocus
GoTo FuncTerminate
End If
End If
Next
End If

FuncTerminate:
ProperDateCheck = flag
End Function

Private Function IsMandatoryFieldsEmpty() As Boolean

Dim isEmpty As Boolean
isEmpty = False

'PERSONAL INFO
'EmpID
If Text1.Text = "" Then
MsgBox "Emp Code Cannot Be Left Empty !"
SSTab1.Tab = 0
Text1.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'EmpName
If Text2.Text = "" Then
MsgBox "Emp Name Cannot Be Left Empty !"
SSTab1.Tab = 0
Text2.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Address
If Text3.Text = "" Then
MsgBox "Address Cannot Be Left Empty !"
SSTab1.Tab = 0
Text3.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'PIN
If Text28.Text = "" Then
MsgBox "PIN Cannot Be Left Empty !"
SSTab1.Tab = 0
Text28.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'DOB
If (Text6.Text = "" And Text7.Text = "" And Text8.Text = "") Or (Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "") Then
MsgBox "DOB Cannot Be Left Empty !"
SSTab1.Tab = 0
Text6.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Gender
If Combo9.Text = "" Then
MsgBox "Gender Cannot Be Left Empty !"
SSTab1.Tab = 0
Combo9.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Caste
If Text29.Text = "" Then
MsgBox "Caste Cannot Be Left Empty !"
SSTab1.Tab = 0
Text29.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Religion
If Text21.Text = "" Then
MsgBox "Religion Cannot Be Left Empty !"
SSTab1.Tab = 0
Text21.SetFocus
isEmpty = True
GoTo FunctionStop
End If

'CONTACT INFO
'Mobile No
If Text4.Text = "" Then
MsgBox "Mobile No Cannot Be Left Empty !"
SSTab1.Tab = 1
Text4.SetFocus
isEmpty = True
GoTo FunctionStop
End If

'IDENTIFICATION INFO
'PAN No
If Text24.Text = "" Then
MsgBox "PAN No Cannot Be Left Empty !"
SSTab1.Tab = 2
Text24.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Aadhaar
If Text25.Text = "" Then
MsgBox "Aadhaar Cannot Be Left Empty !"
SSTab1.Tab = 2
Text25.SetFocus
isEmpty = True
GoTo FunctionStop
End If

'QUALIFICATION INFO
'Educational Qualification
If Text22.Text = "" Then
MsgBox "Educational Qualification Cannot Be Left Empty !"
SSTab1.Tab = 3
Text22.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Technical Qualification
If Text23.Text = "" Then
MsgBox "Technical Qualification Cannot Be Left Empty !"
SSTab1.Tab = 3
Text23.SetFocus
isEmpty = True
GoTo FunctionStop
End If

'OFFICIAL INFO
'Emp Category
If Check1.Value = 0 And Check2.Value = 0 Then
MsgBox "Emp Category Cannot Be Left Empty !"
SSTab1.Tab = 5
Check1.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Employment Type
If Combo6.Text = "" Then
MsgBox "Employment Type Cannot Be Left Empty !"
SSTab1.Tab = 5
Combo6.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Employee Status
If Combo5.Text = "" Then
MsgBox "Employee Status Cannot Be Left Empty !"
SSTab1.Tab = 5
Combo5.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Designation
If Combo1.Text = "" Then
MsgBox "Designation Cannot Be Left Empty !"
SSTab1.Tab = 5
Combo1.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Department
If Combo2.Text = "" Then
MsgBox "Department Cannot Be Left Empty !"
SSTab1.Tab = 5
Combo2.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'DOJ
If (Text9.Text = "" And Text10.Text = "" And Text10.Text = "") Or (Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "") Then
MsgBox "DOJ Cannot Be Left Empty !"
SSTab1.Tab = 5
Text9.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'UAN No
If Text19.Text <> "" Then
If Text18.Text = "" Then
MsgBox "Since EPF No Is Given, UAN No Cannot Be Left Empty !"
SSTab1.Tab = 5
Text18.SetFocus
isEmpty = True
GoTo FunctionStop
End If
End If

'BANK DETAILS
'Bank Name
If Text32.Text = "" Then
MsgBox "Bank Name Cannot Be Left Empty !"
SSTab1.Tab = 6
Text32.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Account No
If Text33.Text = "" Then
MsgBox "Account No Cannot Be Left Empty !"
SSTab1.Tab = 6
Text33.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'IFSC Code
If Text34.Text = "" Then
MsgBox "IFSC Code Cannot Be Left Empty !"
SSTab1.Tab = 6
Text34.SetFocus
isEmpty = True
GoTo FunctionStop
End If

'DUTY INFO
'Duty Shift
If Combo7.Text = "" Then
MsgBox "Duty Shift Cannot Be Left Empty !"
SSTab1.Tab = 7
Combo7.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Shift Time
If Text36.Text = "" Then
MsgBox "Shift Time Cannot Be Left Empty !"
SSTab1.Tab = 7
Text36.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Branch
If Text38.Text = "" Then
MsgBox "Branch Cannot Be Left Empty !"
SSTab1.Tab = 7
Text38.SetFocus
isEmpty = True
GoTo FunctionStop
End If
'Workplace Status
If Combo8.Text = "" Then
MsgBox "Workplace Status Cannot Be Left Empty !"
SSTab1.Tab = 7
Combo8.SetFocus
isEmpty = True
GoTo FunctionStop
End If

FunctionStop:
IsMandatoryFieldsEmpty = isEmpty
Exit Function

End Function

Public Function MandatoryFieldsEmpty() As Boolean

Dim flag As Boolean
If Not (Check1.Value = 1 Or Check2.Value = 1) Then
MsgBox "Mandatory fields can't be left empty !!"
Else
EmpCategory = EmpCategory
End If

For Each Ctrl In Me.Controls
If Ctrl.Tag = "m" Then
If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is ComboBox Then
If Ctrl.Text = "" Then
Exit For
MsgBox "Mandatory fields can't be left empty !!"
flag = True
GoTo FunctionStop
End If
End If
End If
Next

FunctionStop:
MandatoryFieldsEmpty = flag
End Function

Public Function ContactFieldsCheck() As Boolean
Dim flag As Boolean
Dim fields(3) As String

fields(0) = Text4.Text
fields(1) = Text16.Text
fields(2) = Text17.Text

Dim i, ctr As Integer
ctr = 0
For i = 0 To 2
If fields(i) <> "" Then
ctr = ctr + 1
End If
Next

If ctr < 1 Then
flag = False
MsgBox "Atleast one of the fields: Mobile No, Alt Mobile No, Email should be filled !!"
Else
flag = True
End If

ContactFieldsCheck = flag
End Function

Public Function EPFNoChange()

fempcode = Text1.Text
rs.MoveFirst
Do Until rs.EOF
If rs.fields(0) = fempcode Then
Exit Do
Else
rs.MoveNext
End If
Loop
If EmpCategory = "Experienced" Then
If Text19.Text <> rs.fields(17) Then
Combo3.Text = Combo3.List(1)
Combo4.Text = Combo4.List(1)
Else
Combo3.Text = Combo3.List(0)
Combo4.Text = Combo4.List(0)
End If
End If

End Function

Private Sub Check1_Click()
If Check1.Value = 1 Then
EmpCategory = "Fresher"
Check2.Value = 0
Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Label40.Visible = False
Combo3.Text = Combo3.List(2)
Combo4.Text = Combo4.List(2)
Else
Label40.Visible = False
Combo3.Text = ""
Combo4.Text = ""
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
EmpCategory = "Experienced"
Check1.Value = 0
Text18.Enabled = True
Text19.Enabled = True
Text20.Enabled = True
Else
Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False
End If
End Sub

Private Sub Combo1_LostFocus()
'Exiting sub when text is empty to avoid unncessary checking
If Combo1.Text = "" Then
Exit Sub
End If

Dim i As Integer
Dim match As Boolean
match = False

For i = 0 To Combo1.ListCount - 1
If Combo1.Text = Combo1.List(i) Then
match = True
Exit For
End If
Next i

If Not match Then
MsgBox "Invalid Selection for Designation. Please select from given options !!", vbExclamation
Combo1.Text = ""
Combo1.SetFocus
End If
End Sub

Private Sub Combo2_LostFocus()
'Exiting sub when text is empty to avoid unncessary checking
If Combo2.Text = "" Then
Exit Sub
End If

Dim i As Integer
Dim match As Boolean
match = False

For i = 0 To Combo2.ListCount - 1
If Combo2.Text = Combo2.List(i) Then
match = True
Exit For
End If
Next i

If Not match Then
MsgBox "Invalid Selection for Department. Please select from given options !!", vbExclamation
Combo2.Text = ""
Combo2.SetFocus
End If
End Sub

Private Sub Combo5_LostFocus()
'Exiting sub when text is empty to avoid unncessary checking
If Combo5.Text = "" Then
Exit Sub
End If

Dim i As Integer
Dim match As Boolean
match = False

For i = 0 To Combo5.ListCount - 1
If Combo5.Text = Combo5.List(i) Then
match = True
Exit For
End If
Next i

If Not match Then
MsgBox "Invalid Selection for Employee Status. Please select from given options !!", vbExclamation
Combo5.Text = ""
Combo5.SetFocus
End If
End Sub

Private Sub Combo6_LostFocus()
'Exiting sub when text is empty to avoid unncessary checking
If Combo6.Text = "" Then
Exit Sub
End If

Dim i As Integer
Dim match As Boolean
match = False

For i = 0 To Combo6.ListCount - 1
If Combo6.Text = Combo6.List(i) Then
match = True
Exit For
End If
Next i

If Not match Then
MsgBox "Invalid Selection for Employment Type. Please select from given options !!", vbExclamation
Combo6.Text = ""
Combo6.SetFocus
End If
End Sub

Private Sub Combo7_LostFocus()
'Exiting sub when text is empty to avoid unncessary checking
If Combo7.Text = "" Then
Exit Sub
End If

Dim i As Integer
Dim match As Boolean
match = False

For i = 0 To Combo7.ListCount - 1
If Combo7.Text = Combo7.List(i) Then
match = True
Exit For
End If
Next i

If Not match Then
MsgBox "Invalid Selection for Duty Shift. Please select from given options !!", vbExclamation
Combo7.Text = ""
Combo7.SetFocus
End If
End Sub

Private Sub Combo8_LostFocus()
'Exiting sub when text is empty to avoid unncessary checking
If Combo8.Text = "" Then
Exit Sub
End If

Dim i As Integer
Dim match As Boolean
match = False

For i = 0 To Combo8.ListCount - 1
If Combo8.Text = Combo8.List(i) Then
match = True
Exit For
End If
Next i

If Not match Then
MsgBox "Invalid Selection for Workplace Status. Please select from given options !!", vbExclamation
Combo8.Text = ""
Combo8.SetFocus
End If
End Sub

Private Sub Combo9_LostFocus()
'Exiting sub when text is empty to avoid unncessary checking
If Combo9.Text = "" Then
Exit Sub
End If

Dim i As Integer
Dim match As Boolean
match = False

For i = 0 To Combo9.ListCount - 1
If Combo9.Text = Combo9.List(i) Then
match = True
Exit For
End If
Next i

If Not match Then
MsgBox "Invalid Selection for Gender. Please select from given options !!", vbExclamation
Combo9.Text = ""
Combo9.SetFocus
End If
End Sub

Private Sub Command1_Click()

If IsMandatoryFieldsEmpty() = True Then
Exit Sub
End If

'If ContactFieldsCheck() = False Then
'Exit Sub
'End If

'Checking for proper date entry of DOB
If ProperDateCheck("DOB", Text6, Text7, Text8) = True Then
DOB = DateSerial(CInt(Text8.Text), CInt(Text7.Text), CInt(Text6.Text))
sdob = DOB
Else: Exit Sub
End If
'Checking for valid DOB...when DOB is not empty
Dim current_date As Date
current_date = DateSerial(Year(Date), Month(Date), Day(Date))
If DOB > current_date Then
MsgBox "DOB can't be more than Today's Date...Please give proper date !!"
SSTab1.Tab = 0
Text6.SetFocus
Exit Sub
End If

'Checking for proper date entry for DOJ
If ProperDateCheck("DOJ", Text9, Text10, Text11) = True Then
DOJ = DateSerial(CInt(Text11.Text), CInt(Text10.Text), CInt(Text9.Text))
sdoj = DOJ
Else: Exit Sub
End If
'Checking for valid DOJ...when DOJ is not empty
If DOJ < DOB Then
MsgBox "DOJ can't be less than DOB...Enter proper date !!"
SSTab1.Tab = 5
Text9.SetFocus
Exit Sub
End If
'Checking for minimum 18 years difference between DOB and DOJ
If IsEmployeeAdult = False Then
MsgBox "Employee is a Minor" & vbNewLine & "DOJ must be 18 years more than DOB !!"
SSTab1.Tab = 5
Text9.SetFocus
Exit Sub
End If

fempcode = Text1.Text
If rs.RecordCount > 0 Then
rs.MoveFirst
Do Until rs.EOF
fempcode1 = fempcode
If rs.fields(0) = fempcode1 Then
str = 1
MsgBox "Can't save a record with duplicate EmpCode !!"
Exit Sub
Else
rs.MoveNext
str = 0
End If
Loop

If str = 0 Then
Call AddRec
End If

Else
Call AddRec
End If


End Sub


Private Sub Command2_Click()

fempcode = Text1.Text
'Text1.Enabled = False
Dim str As Boolean
If rs.RecordCount = 0 Then
str = False
Else
rs.MoveFirst
Do Until rs.EOF
If rs.fields(0) = fempcode Then
str = True
Call EmptyAllFields
Text1.Text = rs.fields(0)
Text2.Text = rs.fields(1)
Text3.Text = rs.fields(2)
If Not (IsNull(rs.fields(3))) Then
Text27.Text = rs.fields(3)
End If
Text28.Text = rs.fields(4)
Text6.Text = Left(rs.fields(5), 2)
Text7.Text = Mid(rs.fields(5), 4, 2)
Text8.Text = Right(rs.fields(5), 4)
Combo9.Text = rs.fields(6)
Text29.Text = rs.fields(7)
Text21.Text = rs.fields(8)
If Not (IsNull(rs.fields(9))) Then
Text4.Text = rs.fields(9)
End If
If Not (IsNull(rs.fields(10))) Then
Text16.Text = rs.fields(10)
End If
If Not (IsNull(rs.fields(11))) Then
Text17.Text = rs.fields(11)
End If
Text24.Text = rs.fields(12)
Text25.Text = rs.fields(13)
If Not (IsNull(rs.fields(14))) Then
Text26.Text = rs.fields(14)
End If
If Not (IsNull(rs.fields(15))) Then
Text30.Text = rs.fields(15)
End If
If Not (IsNull(rs.fields(16))) Then
Text31.Text = rs.fields(16)
End If
If Not (IsNull(rs.fields(17))) Then
Text35.Text = rs.fields(17)
End If
Text22.Text = rs.fields(18)
Text23.Text = rs.fields(19)
If Not (IsNull(rs.fields(20))) Then
Image1.Picture = LoadPicture(rs.fields(20))
Label7.Caption = rs.fields(20)
End If
If rs.fields(21) = "Fresher" Then
Check1.Value = 1
Else
Check2.Value = 1
End If
Combo6.Text = rs.fields(22)
Combo5.Text = rs.fields(23)
Combo1.Text = rs.fields(24)
Combo2.Text = rs.fields(25)
Text9.Text = Left(rs.fields(26), 2)
Text10.Text = Mid(rs.fields(26), 4, 2)
Text11.Text = Right(rs.fields(26), 4)
If Not (IsNull(rs.fields(27))) Then
Text12.Text = Left(rs.fields(27), 2)
Text13.Text = Mid(rs.fields(27), 4, 2)
Text14.Text = Right(rs.fields(27), 4)
End If
If Not (IsNull(rs.fields(28))) Then
Text5.Text = rs.fields(28)
End If
If Not (IsNull(rs.fields(29))) Then
Text19.Text = rs.fields(29)
End If
If Not (IsNull(rs.fields(30))) Then
Text18.Text = rs.fields(30)
End If
If Not (IsNull(rs.fields(31))) Then
Text20.Text = rs.fields(31)
End If
Combo3.Text = rs.fields(32)
Combo4.Text = rs.fields(33)
Text32.Text = rs.fields(34)
Text33.Text = rs.fields(35)
Text34.Text = rs.fields(36)
Combo7.Text = rs.fields(37)
Text36.Text = rs.fields(38)
Text38.Text = rs.fields(39)
Combo8.Text = rs.fields(40)
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text5.Enabled = True
'Call Disable
'If Label7.Caption = "" Then
'Command6.Enabled = True
'Else
'Command6.Enabled = False
'End If
Command4.Enabled = True
SSTab1.Tab = 0
Call UpdateUserActivity("SEARCH", "Employee", Text1.Text)
Exit Do

Else
rs.MoveNext
str = False
End If
Loop
End If

If str = False Then
MsgBox "No record with the EmpCode found !!"
Call EmptyAllFields
SSTab1.Tab = 0
Text1.SetFocus
Exit Sub
End If

End Sub

Public Function Optional_EmptyFields_Update()
'Code for updating optional fields which are empty
Dim updt_rsIndex(10) As Integer
updt_rsIndex(0) = 3
updt_rsIndex(1) = 9
updt_rsIndex(2) = 10
updt_rsIndex(3) = 11
updt_rsIndex(4) = 20
updt_rsIndex(5) = 27
updt_rsIndex(6) = 28
updt_rsIndex(7) = 29
updt_rsIndex(8) = 30
updt_rsIndex(9) = 31

Dim updt_OptFields(10) As String
updt_OptFields(0) = Text27.Text
updt_OptFields(1) = Text4.Text
updt_OptFields(2) = Text16.Text
updt_OptFields(3) = Text17.Text
updt_OptFields(4) = Label7.Caption
updt_OptFields(5) = sdor
updt_OptFields(6) = Text5.Text
updt_OptFields(7) = Text19.Text
updt_OptFields(8) = Text18.Text
updt_OptFields(9) = Text20.Text

Dim i As Integer
For i = 0 To UBound(updt_OptFields())
If Not (updt_OptFields(i) = "") Then
rs.fields(updt_rsIndex(i)) = updt_OptFields(i)
End If
Next
End Function


Private Function CheckProperResignationUpdate() As Boolean

Dim empStatus, DOR As String
empStatus = Combo5.Text

If empStatus = "Resigned" Or empStatus = "Terminated" Then
DOR = Text12.Text & Text13.Text & Text14.Text
If DOR = "" And Text5.Text = "" Then
CheckProperResignationUpdate = False
Else
CheckProperResignationUpdate = True
End If
End If

End Function


Private Function IsEmployeeAdult() As Boolean
DOJ = DateSerial(CInt(Text11.Text), CInt(Text10.Text), CInt(Text9.Text))
DOB = DateSerial(CInt(Text8.Text), CInt(Text7.Text), CInt(Text6.Text))

Dim Years, Age As Integer
Years = DOJ - DOB
Age = Years / 365

If Age >= 18 Then
IsEmployeeAdult = True
Else
IsEmployeeAdult = False
End If

End Function


Private Sub Command3_Click()
If SSTab1.Tab <> SSTab1.TabsPerRow - 1 Then
SSTab1.Tab = SSTab1.Tab + 1
End If
End Sub

Private Sub Command4_Click()
Text1.Enabled = True

'Checking for valid DOR...when DOR is not empty
If Text12.Text <> "" And Text13.Text <> "" And Text14.Text <> "" Then
If ProperDateCheck("DOR", Text12, Text13, Text14) = True Then
DOR = DateSerial(CInt(Text14.Text), CInt(Text13.Text), CInt(Text12.Text))
DOJ = DateSerial(CInt(Text11.Text), CInt(Text10.Text), CInt(Text9.Text))
sdor = DOJ
Else: Exit Sub
End If
End If

'Checking for valid DOJ...when DOR is not empty
If Text12.Text <> "" And Text13.Text <> "" And Text14.Text <> "" Then
DOR = DateSerial(CInt(Text14.Text), CInt(Text13.Text), CInt(Text12.Text))
DOJ = DateSerial(CInt(Text11.Text), CInt(Text10.Text), CInt(Text9.Text))
If DOR < DOJ Then
MsgBox "DOR can't be less than DOJ...Enter proper date !!"
Text9.SetFocus
Exit Sub
End If
End If

'Checking for minimum 18 years difference between DOB and DOJ
If IsEmployeeAdult = False Then
MsgBox "Employee is a Minor" & vbNewLine & "DOJ must be 18 years more than DOB !!"
SSTab1.Tab = 5
Text9.SetFocus
Exit Sub
End If

'Checking if DOR is given when Employee Status is changed to "Resigned" or Terminated"
If CheckProperResignationUpdate = False Then
MsgBox "Since Employee Status is updated to " & Combo5.Text & vbNewLine & "DOR and Reason of Resignation cannot be left empty !!"
If Not Text5.Text = "" Then
SSTab1.Tab = 5
Text9.SetFocus
Else
SSTab1.Tab = 5
Text5.SetFocus
End If
Exit Sub
End If

Call Optional_EmptyFields_Update
MsgBox "Field(s) updated successfully !!"
Call UpdateUserActivity("UPDATE", "Employee", Text1.Text)
For Each Ctrl In Me.Controls
If Ctrl.Tag = "m" Or Ctrl.Tag = "o" Then
Ctrl.Enabled = False
End If
Next
Command4.Enabled = False

End Sub

Private Sub Command5_Click()
'Dim i As Integer
'Dim Ctrls(23) As Object
'i = 0
'For Each ctrl In Me.Controls
'If i <= 23 Then
'If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
'Set Ctrls(i) = ctrl
'i = i + 1
'Print ctrl.TabIndex
'End If
'End If
'Next

'i = 0
'ctr = 0
'For Each ctrl In Me.Controls
'If ctr <= 16 Then
'Exit For
'Else
'If ctrl.Index = ctr + 1 Then
'Print ctrl.Name
'End If
'End If
'ctr = ctr + 1
'Next

'Set Ctrls(0) = Text1
'Set Ctrls(1) = Text2
'Set Ctrls(2) = Text15
'Set Ctrls(3) = Text21
'Set Ctrls(4) = Combo1
'For i = 0 To 4
'Ctrls(i).Text = Ctrls(i).Name
'Next

Unload Me
End

End Sub

Private Sub Command6_Click()
Load Form3
Form3.Show
End Sub

Private Sub Command7_Click()
If SSTab1.Tab <> 0 Then
SSTab1.Tab = SSTab1.Tab - 1
End If
End Sub

Private Sub Form_Load()
'Set Con = New ADODB.Connection
'Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Infosys.mdb;Persist Security Info=False"
'Con.Open
'Conn = GetDbConn
Call OpenDbConn

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "Employee", gCon, adOpenDynamic, adLockOptimistic, adCmdTable

'Set rs2 = New ADODB.Recordset
'rs2.Open "Salary", CON, adOpenDynamic, adLockOptimistic, adCmdTable

'Set con1 = New ADODB.connection
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Infosys.mdb;Persist Security Info=False"
'con1.Open
'Set RS1 = New ADODB.Recordset
'RS1.CursorLocation = adUseClient
'RS1.Open "Select * from Employee", con, adOpenKeyset, adLockPessimistic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh


Set rsDesig = New ADODB.Recordset
rsDesig.Open "select * from Designation", gCon, adOpenStatic, adLockReadOnly
rsDesig.MoveFirst
Do Until rsDesig.EOF
With Combo1
.AddItem rsDesig.fields(0)
End With
rsDesig.MoveNext
Loop

Set rsDept = New ADODB.Recordset
rsDept.Open "select * from Department", gCon, adOpenStatic, adLockReadOnly
rsDept.MoveFirst
Do Until rsDept.EOF
With Combo2
.AddItem rsDept.fields(0)
End With
rsDept.MoveNext
Loop

With Combo3
.AddItem "Need To Transfer"
.AddItem "Already Transferred"
.AddItem "Not Issued"
End With

With Combo4
.AddItem "Need To Transfer"
.AddItem "Already Transferred"
.AddItem "Not Issued"
End With

With Combo5
.AddItem "Active"
.AddItem "On Leave"
.AddItem "Resigned"
.AddItem "Terminated"
End With

With Combo6
.AddItem "Full Time"
.AddItem "Part Time"
End With

With Combo7
.AddItem "Day"
.AddItem "Night"
.AddItem "Rotational"
End With

With Combo8
.AddItem "On-Site"
.AddItem "Hybrid"
.AddItem "Work From Home"
End With

With Combo9
.AddItem "Male"
.AddItem "Female"
.AddItem "Trans"
End With

epfFormat = "100 "

Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text5.Enabled = False
Command4.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False

Label40.Visible = False

SSTab1.Tab = 0

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text19_LostFocus()

If Text19.Text <> "" Then

Text18.Tag = "m"
Label40.Visible = True
If Left(Text19.Text, 4) = epfFormat Then
Combo3.Text = Combo3.List(1)
Else
Combo3.Text = Combo3.List(0)
End If

Else
Text18.Tag = "o"
Text18.Enabled = False
Combo3.Text = Combo3.List(2)

End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or _
(KeyAscii >= 97 And KeyAscii <= 122) Or _
(KeyAscii = 8) Or _
(KeyAscii = 32)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text20_LostFocus()
If Text20.Text = "" Then
Combo4.Text = Combo4.List(2)
Else
Combo4.Text = Combo4.List(1)
End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or _
(KeyAscii >= 97 And KeyAscii <= 122) Or _
(KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text29_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or _
(KeyAscii >= 97 And KeyAscii <= 122) Or _
(KeyAscii = 8) Or _
(KeyAscii = 32)) Then
KeyAscii = 0
End If
If (KeyAscii >= 97 And KeyAscii <= 122) Then
KeyAscii = KeyAscii - 32
End If
End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8)) Then
KeyAscii = 0
End If
End Sub
