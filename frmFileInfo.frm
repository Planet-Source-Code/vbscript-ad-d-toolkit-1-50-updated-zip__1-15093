VERSION 5.00
Begin VB.Form frmFileInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application File Properties"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmFileInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   5175
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Copyright"
      Top             =   0
      Width           =   3855
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00808080&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3910
      TabIndex        =   0
      Top             =   50
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "CompanyName"
      Top             =   240
      Width           =   3855
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "ProductName"
      Top             =   480
      Width           =   5175
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Title"
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Version"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Trademarks"
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Path"
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Description"
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   885
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmFileInfo.frx":030A
      Top             =   1680
      Width           =   5175
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Text1.Text = App.Comments
    Text2.Text = App.CompanyName
    Text4.Text = App.FileDescription
    Text6.Text = App.LegalCopyright
    Text7.Text = App.LegalTrademarks
    Text8.Text = App.Major & "." & App.Minor & "." & App.Revision
    Text9.Text = App.Path
    Text10.Text = App.ProductName
    Text11.Text = App.Title
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
