VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6624
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6624
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2412
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   9492
      _ExtentX        =   16743
      _ExtentY        =   4255
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   3132
   End
   Begin VB.Frame Frame2 
      Caption         =   "Record Operation"
      Height          =   3852
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   2772
      Begin VB.CommandButton Command6 
         Caption         =   "Close"
         Height          =   492
         Left            =   480
         TabIndex        =   11
         Top             =   3240
         Width           =   1572
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete"
         Height          =   492
         Left            =   600
         TabIndex        =   10
         Top             =   2640
         Width           =   1332
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Update"
         Height          =   492
         Left            =   600
         TabIndex        =   9
         Top             =   2040
         Width           =   1332
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   492
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   1332
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         Height          =   492
         Left            =   600
         TabIndex        =   7
         Top             =   840
         Width           =   1332
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   492
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry Form"
      Height          =   3852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6732
      Begin VB.TextBox Text6 
         Height          =   372
         Left            =   2280
         TabIndex        =   19
         Top             =   3360
         Width           =   4332
      End
      Begin VB.TextBox Text5 
         Height          =   372
         Left            =   2280
         TabIndex        =   17
         Top             =   2760
         Width           =   2412
      End
      Begin VB.TextBox Text4 
         Height          =   372
         Left            =   2280
         TabIndex        =   15
         Top             =   2160
         Width           =   2412
      End
      Begin VB.TextBox Text3 
         Height          =   372
         Left            =   2280
         TabIndex        =   13
         Top             =   1560
         Width           =   4212
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   2412
      End
      Begin VB.Label Label6 
         Caption         =   "EmpPhoto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   1212
      End
      Begin VB.Label Label5 
         Caption         =   "MobileNo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "EmpName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "EmpCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1212
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

