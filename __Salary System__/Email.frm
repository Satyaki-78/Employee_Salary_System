VERSION 5.00
Begin VB.Form Form99 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send Email"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim objMessage As Object
    Dim objConfig As Object
    Dim fields As Object
    Dim attachmentPath As String
    Dim emailSubject As String
    Dim emailBody As String

    ' Create CDO message object
    Set objMessage = CreateObject("CDO.Message")
    Set objConfig = CreateObject("CDO.Configuration")

    ' Configure SMTP server settings
    Set fields = objConfig.fields

    With fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  ' Send using SMTP
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "satyakid78@gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "ppwt xlxj yhkl sacg"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With

    ' Assign configuration to message
    Set objMessage.Configuration = objConfig

    ' Set email details (check for optional fields)
    emailSubject = ""  'Change "" to provide a subject (or leave empty)
    emailBody = ""  'Change "" to provide body text (or leave empty)
    attachmentPath = "D:\Salary System\SalarySystem11\Image\1.jpg" ' Change or leave empty if no attachment is needed

    With objMessage
        .To = "satyakid828@gmail.com"
        .From = "satyakid78@gmail.com"
        .Subject = emailSubject
        .TextBody = emailBody
        .AddAttachment attachmentPath
        .Send  'Send email
    End With

    ' Cleanup
    Set objMessage = Nothing
    Set objConfig = Nothing
    Set fields = Nothing

    MsgBox "Email sent successfully!", vbInformation, "Success"
End Sub
