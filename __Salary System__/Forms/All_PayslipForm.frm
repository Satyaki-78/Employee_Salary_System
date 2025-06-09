VERSION 5.00
Begin VB.Form Form62 
   Caption         =   "Form8"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   LinkTopic       =   "Form8"
   ScaleHeight     =   3660
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Payslip"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2880
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Payment Month"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Payment Year"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "Form62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim PayPeriod As String
PayPeriod = Combo1.Text & " " & Text2.Text

If SalaryDataExists(PayPeriod) = False Then
MsgBox "Cannot Generate Report !!" & vbNewLine & "No Record Found For Payment Period '" & PayPeriod & "' !!"
Exit Sub
End If

Load DataEnvironment3

Unload DataReport1
Set DataReport1 = Nothing
Set DataReport1 = New DataReport1

DataEnvironment3.Commands("Command1").CommandText = "SELECT * FROM Salary WHERE Pay_Period = '" & PayPeriod & "'"
DataEnvironment3.Commands("Command1").Execute
Set DataReport1.DataSource = Nothing
DataReport1.DataMember = ""
Set DataReport1.DataSource = DataEnvironment3
DataReport1.DataMember = "Command1"
DataReport1.Refresh
DataReport1.Show

Unload DataEnvironment3

Call UpdateUserActivity("ALL PAYSLIP GENERATE", "Salary", PayPeriod)

End Sub


Private Function SalaryDataExists(PayPeriod As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT Pay_Period FROM Salary", gCon, adOpenStatic, adLockReadOnly

Dim flag As Boolean

If rs.RecordCount = 0 Then
SalaryDataExists = False
Exit Function
End If

rs.MoveFirst
Do Until rs.EOF
If rs("Pay_Period") = PayPeriod Then
flag = True
Exit Do
Else
flag = False
rs.MoveNext
End If
Loop
SalaryDataExists = flag
End Function


Private Sub Form_Load()

With Combo1
.AddItem "January"
.AddItem "February"
.AddItem "March"
.AddItem "April"
.AddItem "May"
.AddItem "June"
.AddItem "July"
.AddItem "August"
.AddItem "September"
.AddItem "October"
.AddItem "November"
.AddItem "December"
End With

End Sub
