VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReadMe 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmReadMe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtfReadMe 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8705
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReadMe.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReadMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================================
' This form does nothing more than display a textfile.  It
' displays either the update.txt file or the readme.txt file
' depending on which menu item is chosen.
'============================================================

Private Sub Form_Load()
    '============================================================
    ' This runs when the form is loaded
    '============================================================
    DoGradient Me
End Sub

Private Sub rtfReadMe_DblClick()
    '============================================================
    ' This closes the form
    '============================================================
    Unload Me
End Sub

Public Sub GetText(FName As String)
    '============================================================
    ' Reads in the correct text file and displays it in the RTF
    ' Box.  It also sets the appropriate window title.
    '============================================================
    Select Case FName
        Case "ReadMe"
            Me.Caption = "Read Me File"
            rtfReadMe.LoadFile App.Path & "\readme.txt"
        Case "Update"
            Me.Caption = "Update History"
            rtfReadMe.LoadFile App.Path & "\update.txt"
        Case Else
            Unload Me
    End Select
End Sub

