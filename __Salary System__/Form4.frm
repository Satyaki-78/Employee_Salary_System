VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6045
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10380
   LinkTopic       =   "Form4"
   ScaleHeight     =   6045
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add Department"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5160
      TabIndex        =   2
      Text            =   "Enter Department Name..."
      Top             =   2160
      Width           =   4815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add Department From Here"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Departments"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con11 As ADODB.Connection
Dim rs11 As ADODB.Recordset


Private Sub Form_Load()
Set con11 = New ADODB.Connection
con11.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\SalarySystem\Infosys.mdb;Persist Security Info=False"
con11.Open
Set rs11 = New ADODB.Recordset
'rs11.Open "Designation", con11, adOpenDynamic, adLockOptimistic, adCmdTable
rs11.Open "select * from Designation", con11, adOpenStatic, adLockReadOnly
rs11.MoveFirst
Do Until rs11.EOF
With Combo1
.AddItem rs11.fields(0)
End With
rs11.MoveNext
Loop

'MsgBox rs11.RecordCount

'If rs.RecordCount > 0 Then
'rs.MoveFirst
'Do Until rs.EOF
'With Combo1
'.AddItem rs.Fields(0)
'End With
'Loop
'Else
'MsgBox "No Records Found !!"
'End If




MsgBox rs11.RecordCount
MsgBox Adodc1.Recordset.RecordCount
End Sub
