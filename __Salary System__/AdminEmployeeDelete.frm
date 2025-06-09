VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form101 
   Caption         =   "Admin Employee Table Delete Form"
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
   Icon            =   "AdminEmployeeDelete.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   18060
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2820
      Left            =   120
      TabIndex        =   131
      Top             =   6480
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   4974
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
      Caption         =   "Employee Table Data"
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
      Left            =   12840
      TabIndex        =   120
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
      TabIndex        =   119
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
      TabIndex        =   52
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
      Height          =   3135
      Left            =   15960
      TabIndex        =   50
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
         Height          =   495
         Left            =   240
         TabIndex        =   49
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
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
         Top             =   1440
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
         TabIndex        =   47
         Top             =   480
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   57
      Top             =   120
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   8
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
      TabPicture(0)   =   "AdminEmployeeDelete.frx":1084A
      Tab(0).ControlEnabled=   -1  'True
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
      TabPicture(1)   =   "AdminEmployeeDelete.frx":10866
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(1)=   "Label20"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "Label31"
      Tab(1).Control(4)=   "Text17"
      Tab(1).Control(5)=   "Text16"
      Tab(1).Control(6)=   "Text4"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Identification Info"
      TabPicture(2)   =   "AdminEmployeeDelete.frx":10882
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label52"
      Tab(2).Control(1)=   "Label51"
      Tab(2).Control(2)=   "Label46"
      Tab(2).Control(3)=   "Label44"
      Tab(2).Control(4)=   "Label39"
      Tab(2).Control(5)=   "Label53"
      Tab(2).Control(6)=   "Label54"
      Tab(2).Control(7)=   "Label62"
      Tab(2).Control(8)=   "Text26"
      Tab(2).Control(9)=   "Text25"
      Tab(2).Control(10)=   "Text24"
      Tab(2).Control(11)=   "Text30"
      Tab(2).Control(12)=   "Text31"
      Tab(2).Control(13)=   "Text35"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "&Qualification Info"
      TabPicture(3)   =   "AdminEmployeeDelete.frx":1089E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label50"
      Tab(3).Control(1)=   "Label49"
      Tab(3).Control(2)=   "Label48"
      Tab(3).Control(3)=   "Label47"
      Tab(3).Control(4)=   "Text23"
      Tab(3).Control(5)=   "Text22"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "P&hoto Upload"
      TabPicture(4)   =   "AdminEmployeeDelete.frx":108BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image1"
      Tab(4).Control(1)=   "Label6"
      Tab(4).Control(2)=   "Command6"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "O&fficial Info"
      TabPicture(5)   =   "AdminEmployeeDelete.frx":108D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Bank Details"
      TabPicture(6)   =   "AdminEmployeeDelete.frx":108F2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label59"
      Tab(6).Control(1)=   "Label60"
      Tab(6).Control(2)=   "Label61"
      Tab(6).Control(3)=   "Label64"
      Tab(6).Control(4)=   "Label65"
      Tab(6).Control(5)=   "Label66"
      Tab(6).Control(6)=   "Text32"
      Tab(6).Control(7)=   "Text33"
      Tab(6).Control(8)=   "Text34"
      Tab(6).ControlCount=   9
      TabCaption(7)   =   "&Duty Details"
      TabPicture(7)   =   "AdminEmployeeDelete.frx":1090E
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
         Left            =   2040
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
         TabIndex        =   97
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
            TabIndex        =   98
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
            TabIndex        =   122
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
            TabIndex        =   121
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
            TabIndex        =   118
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
            TabIndex        =   117
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
            TabIndex        =   116
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   109
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
            TabIndex        =   108
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
            TabIndex        =   107
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
            TabIndex        =   106
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
            TabIndex        =   105
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
            TabIndex        =   104
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
            TabIndex        =   103
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
            TabIndex        =   102
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
            TabIndex        =   101
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
            TabIndex        =   100
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
            TabIndex        =   99
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
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "o"
         Top             =   1080
         Width           =   1572
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71760
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "o"
         Top             =   1800
         Width           =   1572
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71760
         TabIndex        =   13
         Tag             =   "o"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "m"
         Top             =   720
         Width           =   1452
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   2
         Tag             =   "m"
         Top             =   1920
         Width           =   9375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   2040
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
         Left            =   2040
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
         Left            =   2520
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
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "m"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   5520
         TabIndex        =   10
         Tag             =   "m"
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   3
         Tag             =   "o"
         Top             =   2400
         Width           =   9375
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "m"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   5520
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
         Left            =   -72120
         TabIndex        =   132
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         Left            =   -73320
         TabIndex        =   77
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
         Left            =   -73680
         TabIndex        =   76
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
         Left            =   -73200
         TabIndex        =   75
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
         Left            =   600
         TabIndex        =   74
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
         Left            =   480
         TabIndex        =   73
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
         Left            =   600
         TabIndex        =   72
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
         Left            =   1080
         TabIndex        =   71
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
         Left            =   2040
         TabIndex        =   70
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
         Left            =   840
         TabIndex        =   69
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
         Left            =   1680
         TabIndex        =   68
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
         Left            =   1680
         TabIndex        =   66
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
         Left            =   1680
         TabIndex        =   65
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
         Left            =   1680
         TabIndex        =   64
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
         Left            =   5160
         TabIndex        =   63
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
         Left            =   1680
         TabIndex        =   62
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
         Left            =   1680
         TabIndex        =   61
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
         Left            =   1200
         TabIndex        =   60
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
         Left            =   5160
         TabIndex        =   59
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
         Left            =   4440
         TabIndex        =   58
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
         Left            =   4200
         TabIndex        =   67
         Top             =   3720
         Width           =   855
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
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
      Left            =   13320
      TabIndex        =   51
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "Form101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim fempcode As String
Dim empFound As Integer


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


Private Function Display_Record()
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
End Function


Private Sub Command2_Click()

fempcode = Text1.Text
'Text1.Enabled = False
If rs.RecordCount = 0 Then
empFound = False
Else
rs.MoveFirst
Do Until rs.EOF
If rs.fields(0) = fempcode Then
empFound = True
Call EmptyAllFields
Call Display_Record
Call UpdateUserActivity("SEARCH", "Employee", Text1.Text)
Exit Do

Else
rs.MoveNext
empFound = False
End If
Loop
End If

If empFound = False Then
MsgBox "No record with the EmpCode " & fempcode & " found !!", vbCritical
Call EmptyAllFields
SSTab1.Tab = 0
Text1.SetFocus
Exit Sub
End If

End Sub


Private Sub Command4_Click()

If empFound = True Then
Dim response As Integer
reponse = MsgBox("Are you sure you want to delete this record ?", vbYesNo, "Confirmation")
If reponse = vbYes Then
rs.Delete
MsgBox "Record Successfully Deleted !!", vbInformation
Call UpdateUserActivity("DELETE", "Employee", Text1.Text)
End If
End If

End Sub


Private Sub Command5_Click()
Unload Me
End
End Sub


Private Sub Form_Load()

Call OpenDbConn

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "Employee", gCon, adOpenDynamic, adLockOptimistic, adCmdTable

Set DataGrid1.DataSource = rs
DataGrid1.Refresh

SSTab1.Tab = 0

End Sub
