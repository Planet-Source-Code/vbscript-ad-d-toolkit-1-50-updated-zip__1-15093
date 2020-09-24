VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4575
   Begin VB.TextBox txtRegCode 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User's Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User's Company"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User's Email Address"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label lblRegCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Number"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2895
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================================
' This form allows the user to register the program.  This
' section declares any global variables used in the form.
'============================================================

Private Sub Form_Load()
    '============================================================
    ' This is run whenever the form is opened.  This gradiates
    ' the form background, and populates the fields with default
    ' data.
    '============================================================
    DoGradient Me
    txtName.Text = "Your Name"
    txtCompany.Text = "Your Company"
    txtEmail.Text = "your@email.com"
End Sub

Private Sub cmdReg_Click()
    '============================================================
    ' This saves the user data to the registry, reruns the RegID
    ' check and closes the registration box.
    '============================================================
    SaveSetting "AD&D Tools", "Info", "User", txtName.Text
    SaveSetting "AD&D Tools", "Info", "Company", txtCompany.Text
    SaveSetting "AD&D Tools", "Info", "Email", txtEmail.Text
    SaveSetting "AD&D Tools", "Info", "RegID", txtRegCode.Text
    SaveSetting "AD&D Tools", "Info", "RunOnce", "1"
    Unload Me
    Call Main
    SendMail
End Sub

Function SendMail()
    '============================================================
    ' This code asks the user if they would like to email their
    ' registration information back to the author.
    '============================================================
    Dim Message As String, ToUser As String
    Dim FromUser As String, Subject As String
    Dim SendReg As Integer
    Dim Name, Email, Company, RegCode
    SendReg = MsgBox("Send registration information to author?", vbInformation + vbYesNo, "Complete Registration")
    Select Case SendReg
        Case 6      ' Yes
            Name = GetSetting("AD&D Tools", "Info", "User")
            Email = GetSetting("AD&D Tools", "Info", "Email")
            Company = GetSetting("AD&D Tools", "Info", "Company")
            RegCode = GetSetting("AD&D Tools", "Info", "RegID")
            ToUser = "vbscript@sturm.org"
            FromUser = Email
            Subject = "The following person registered the program."
            Message = "Name: " & Name & vbCrLf & "Company: " & Company _
                & vbCrLf & "Email: " & Email & vbCrLf & "Reg Code: " & RegCode
            frmReadMe.rtfReadMe.Text = "The following message will be sent via Email:" & vbCrLf & vbCrLf _
                & "To: " & ToUser & vbCrLf _
                & "From: " & FromUser & vbCrLf _
                & "Subject: " & Subject & vbCrLf & vbCrLf _
                & "Body: " & vbCrLf & Message
            frmReadMe.Show
            MsgBox "Thank you for registering." & vbCrLf & "Your email has been sent.", vbOKOnly + vbExclamation, "Thank You!"
        Case 7      ' No
            MsgBox "Thank you for registering." & vbCrLf & "No mail has been sent.", vbOKOnly + vbExclamation, "Thank You!"
        Case Else   ' Treated as a No
            MsgBox "Thank you for registering." & vbCrLf & "No mail has been sent.", vbOKOnly + vbExclamation, "Thank You!"
    End Select
End Function

Private Sub txtName_GotFocus()
    '============================================================
    ' Selects all text in the textbox
    '============================================================
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtCompany_GotFocus()
    '============================================================
    ' Selects all text in the textbox
    '============================================================
    txtCompany.SelStart = 0
    txtCompany.SelLength = Len(txtCompany.Text)
End Sub

Private Sub txtEmail_GotFocus()
    '============================================================
    ' Selects all text in the textbox
    '============================================================
    txtEmail.SelStart = 0
    txtEmail.SelLength = Len(txtEmail.Text)
End Sub

Private Sub txtRegCode_GotFocus()
    '============================================================
    ' Selects all text in the textbox
    '============================================================
    txtRegCode.SelStart = 0
    txtRegCode.SelLength = Len(txtRegCode.Text)
End Sub

