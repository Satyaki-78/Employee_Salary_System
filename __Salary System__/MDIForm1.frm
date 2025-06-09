VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Infosys Salary System"
   ClientHeight    =   6375
   ClientLeft      =   105
   ClientTop       =   735
   ClientWidth     =   10605
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":1084A
   WindowState     =   2  'Maximized
   Begin VB.Menu System 
      Caption         =   "System"
      Begin VB.Menu UserLogin 
         Caption         =   "User Login"
         Shortcut        =   ^U
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu AdminLogin 
         Caption         =   "Admin Login"
         Shortcut        =   ^A
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu DataEntry 
      Caption         =   "Data Entry"
      Begin VB.Menu EmployeeData 
         Caption         =   "Employee Data"
         Shortcut        =   ^E
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu SalaryData 
         Caption         =   "Salary Data"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu MDEntry 
      Caption         =   "Admin Data Operation"
      Begin VB.Menu UserSignin 
         Caption         =   "New User Sign In"
      End
      Begin VB.Menu sep121 
         Caption         =   "-"
      End
      Begin VB.Menu AdminSignin 
         Caption         =   "New Admin Sign In"
      End
      Begin VB.Menu sep123 
         Caption         =   "-"
      End
      Begin VB.Menu DesigEntryForm 
         Caption         =   "Designation Update  Form"
      End
      Begin VB.Menu sep13 
         Caption         =   "-"
      End
      Begin VB.Menu DeptEntryForm 
         Caption         =   "Department Update  Form"
      End
      Begin VB.Menu sep14 
         Caption         =   "-"
      End
      Begin VB.Menu EmpDataDelete 
         Caption         =   "Employee Data Delete Form"
      End
      Begin VB.Menu sep15 
         Caption         =   "-"
      End
      Begin VB.Menu SalaryDataDelete 
         Caption         =   "Salary Data Delete Form"
      End
   End
   Begin VB.Menu Report 
      Caption         =   "Payslip"
      Begin VB.Menu IndividualReport 
         Caption         =   "Individual"
      End
      Begin VB.Menu sep16 
         Caption         =   "-"
      End
      Begin VB.Menu AllReport 
         Caption         =   "All"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AdminLogin_Click()
Load Form77
Form77.Show vbModal
End Sub

Private Sub AdminSignin_Click()
Load Form88
Form88.Show vbModal
End Sub

Private Sub AllReport_Click()
Load Form62
Form62.Show vbModal
End Sub

Private Sub DeptEntryForm_Click()
Load Form4
Form4.Show vbModal
End Sub

Private Sub DesigEntryForm_Click()
Load Form5
Form5.Show vbModal
End Sub

Private Sub EmpDataDelete_Click()
Load Form101
Form101.Show
End Sub

Private Sub EmployeeData_Click()
Load Form1
Form1.Show
End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub

Private Sub Login_Click()
Load Form7
Form7.Show vbModal
End Sub

Private Sub IndividualReport_Click()
Load Form61
Form61.Show vbModal
End Sub

Private Sub MDIForm_Load()
Call SetDbConn
End Sub

Private Sub SalaryData_Click()
Load Form2
Form2.Show
End Sub

Private Sub SignIn_Click()
Load Form8
Form8.Show vbModal
End Sub

Private Sub SalaryDataDelete_Click()
Load Form201
Form201.Show
End Sub

Private Sub UserLogin_Click()
Load Form78
Form78.Show vbModal
End Sub

Private Sub UserSignin_Click()
Load Form89
Form89.Show vbModal
End Sub
