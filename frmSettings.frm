VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Frame frameDBCount 
      Caption         =   "Character Count"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
      Begin VB.TextBox txtCharCount 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCharCount2 
         Caption         =   "characters in the database."
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblCharCount1 
         Caption         =   "There are currently"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame frameDatabase 
      Caption         =   "Current Database"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4455
      Begin VB.TextBox txtDatabase 
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frameIcons 
      Caption         =   "Icon Size"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton optSmall 
         Caption         =   "Use Small Icons (16x16)"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optLarge 
         Caption         =   "Use Large Icons (32x32)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================================
' Sets up global variables used through out the form
'============================================================
Dim db As Database
Dim rs As Recordset

Private Sub Form_Load()
    '============================================================
    ' This function runs when the form is loaded.  It gets the
    ' count of the number of records in the database and displays
    ' the path to the current database.
    '============================================================
    Dim Count As Integer
    Set db = OpenDatabase(DBPath)
    Set rs = db.OpenRecordset("tblCharInfo", dbOpenDynaset)
    txtDatabase.Text = DBPath
    rs.MoveFirst
    Do Until rs.EOF
        Count = Count + 1
        rs.MoveNext
    Loop
    txtCharCount.Text = Count
End Sub

Private Sub cmdSave_Click()
    '============================================================
    ' Saves the information and reloads the main form after
    ' closing itself.
    '============================================================
    Unload Me
    frmMain.RefreshForm
End Sub

Private Sub cmdClose_Click()
    '============================================================
    ' Discards any changes and closes the form
    '============================================================
    Unload Me
End Sub
Private Sub optLarge_Click()
    '============================================================
    ' This sets the icon size of the toolbar in the registry to
    ' 0, which equals Large (32x32).  It also makes sure the
    ' value of the optSmall control is set to false, or off.
    '============================================================
    optSmall.Value = False
    optLarge.Value = True
    SaveSetting "AD&D Tools", "Settings", "IconSize", 0
End Sub

Private Sub optSmall_Click()
    '============================================================
    ' This sets the icon size of the toolbar in the registry to
    ' 1, which equals Small (16x16).  It also makes sure the
    ' value of the optLarge control is set to false, or off.
    '============================================================
    optSmall.Value = True
    optLarge.Value = False
    SaveSetting "AD&D Tools", "Settings", "IconSize", 1
End Sub

