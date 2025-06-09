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
      Begin VB.Menu Login 
         Caption         =   "User Login"
         Shortcut        =   ^U
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu Adminlogin 
         Caption         =   "Admin Login"
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
      Begin VB.Menu Salarydata 
         Caption         =   "Salary Data"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu MDEntry 
      Caption         =   "Admin Data Operation"
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
      Begin VB.Menu EmpDataUpdate 
         Caption         =   "Employee Data Delete Form"
      End
      Begin VB.Menu sep15 
         Caption         =   "-"
      End
      Begin VB.Menu Salarydataupdate 
         Caption         =   "Salary Data Delete Form"
      End
   End
   Begin VB.Menu Report 
      Caption         =   "Report"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DeptEntryForm_Click()
Load Form4
Form4.Show
End Sub

Private Sub DesigEntryForm_Click()
Load Form5
Form5.Show
End Sub

Private Sub EmpDataUpdate_Click()
Load Form102
Form102.Show
End Sub

Private Sub EmployeeData_Click()
Load Form1
Form1.Show
End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub

Private Sub Report_Click()
DataReport1.Show
End Sub

Private Sub SalaryData_Click()
Load Form2
Form2.Show
End Sub
