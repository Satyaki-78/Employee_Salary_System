VERSION 5.00
Begin VB.Form Form61 
   Caption         =   "Form7"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   12.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   4440
   ScaleWidth      =   7200
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
      TabIndex        =   3
      Top             =   3120
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
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
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
      Left            =   2880
      TabIndex        =   2
      Top             =   2160
      Width           =   930
   End
   Begin VB.TextBox Text1 
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
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   2655
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
      Left            =   480
      TabIndex        =   6
      Top             =   1320
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
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter Employee ID"
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
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function RecordFound()

End Function


Private Sub Command1_Click()

Dim payPeriod As String
payPeriod = Combo1.Text & " " & Text2.Text

If SalaryDataExists(Text1.Text, payPeriod) = False Then
MsgBox "Cannot Generate Report !!" & vbNewLine & "No Record Found For Employee ID '" & Text1.Text & "' and Payment Period '" & payPeriod & "' !!"
Exit Sub
End If

Load DataEnvironment3

Unload DataReport1
Set DataReport1 = Nothing
Set DataReport1 = New DataReport1

DataEnvironment3.Commands("Command1").CommandText = "SELECT * FROM Salary WHERE EmpCode = '" & Text1.Text & "' AND Pay_Period = '" & payPeriod & "'"
DataEnvironment3.Commands("Command1").Execute
Set DataReport1.DataSource = Nothing
DataReport1.DataMember = ""
Set DataReport1.DataSource = DataEnvironment3
DataReport1.DataMember = "Command1"
DataReport1.Refresh
DataReport1.Show

Unload DataEnvironment3

Call UpdateUserActivity("INDIVIDUAL PAYSLIP GENERATE", "Salary", Text1.Text & ", " & payPeriod)

End Sub


Private Function SalaryDataExists(empID As String, payPeriod As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT EmpCode, Pay_Period FROM Salary", gCon, adOpenStatic, adLockReadOnly

Dim flag As Boolean

If rs.RecordCount = 0 Then
SalaryDataExists = False
Exit Function
End If

rs.MoveFirst
Do Until rs.EOF
If rs("EmpCode") = empID And rs("Pay_Period") = payPeriod Then
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



