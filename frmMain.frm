VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Deckyon's AD&D Toolkit"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10980
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3784
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4692
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":614C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6466
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6780
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":70CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":91BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AC02
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B236
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B550
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B86A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB84
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BE9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C1B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   120
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbTools 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgLarge"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Settings"
            Description     =   "Settings"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Description     =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Character"
                  Text            =   "Character"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Thief"
                  Text            =   "Rogue Stats"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Turning"
                  Text            =   "Turning The Undead"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dice"
            Description     =   "Roll The Dice"
            Object.ToolTipText     =   "Roll the Dice"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ViewPC"
            Description     =   "View PC"
            Object.ToolTipText     =   "View PC"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ViewNPC"
            Description     =   "View NPC"
            Object.ToolTipText     =   "View NPC"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddPC"
            Description     =   "Add PC"
            Object.ToolTipText     =   "Add New PC"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddNPC"
            Description     =   "Add NPC"
            Object.ToolTipText     =   "Generate/Add new NPC"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cascade"
            Description     =   "Cascade"
            Object.ToolTipText     =   "Cascade Windows"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileV"
            Description     =   "Tile Vertically"
            Object.ToolTipText     =   "Tile Vertically"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileH"
            Description     =   "Tile Horizontally"
            Object.ToolTipText     =   "Tile Horizontally"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Arrange"
            Description     =   "Arrange"
            Object.ToolTipText     =   "Arrange Windows"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Updates"
            Description     =   "Updates"
            Object.ToolTipText     =   "Updates"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Readme"
            Description     =   "Readme"
            Object.ToolTipText     =   "Read Me"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Description     =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Description     =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6990
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13732
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "2/9/2001"
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "4:36 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
         Begin VB.Menu mnuFilePrintChar 
            Caption         =   "Print &Character"
         End
         Begin VB.Menu mnuFilePrintThief 
            Caption         =   "Print &Thieving Skills"
         End
         Begin VB.Menu mnuFilePrintTurn 
            Caption         =   "Print Turning &Undead"
         End
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnuFileSpace01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsDice 
         Caption         =   "&Dice Roller"
      End
      Begin VB.Menu mnuToolsSpace01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsPC 
         Caption         =   "PC Tools"
         Begin VB.Menu mnuToolsPCView 
            Caption         =   "&PC Viewer"
         End
         Begin VB.Menu mnuToolsPCAdd 
            Caption         =   "&Add PC"
         End
      End
      Begin VB.Menu mnuToolsSpace02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsNPC 
         Caption         =   "NPC Tools"
         Begin VB.Menu mnuToolsNPCView 
            Caption         =   "&NPC Viewer"
         End
         Begin VB.Menu mnuToolsNPCAdd 
            Caption         =   "&Generate NPC"
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowTileH 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowSpace01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowWindows 
         Caption         =   "&Windows"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpUpdates 
         Caption         =   "&Updates"
      End
      Begin VB.Menu mnuHelpReadMe 
         Caption         =   "&Read Me"
      End
      Begin VB.Menu mnuHelpSpace00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpRegister 
         Caption         =   "R&egistration"
      End
      Begin VB.Menu mnuHelpSpace01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================================
' This section sets up any global variables used throughout
' the form
'============================================================
Dim db As Database
Dim rs As Recordset

Private Sub MDIForm_Load()
    '============================================================
    ' This section is run whenever the program is launched.
    '============================================================
    Dim IconSize As Integer, IsReg As Integer
    Set db = OpenDatabase(DBPath)
    Set rs = db.OpenRecordset("tblCharInfo", dbOpenDynaset)
    RefreshForm
    IsReg = GetSetting("AD&D Tools", "Info", "RunOnce")
    If IsReg = 1 Then mnuHelpRegister.Enabled = False
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
    '============================================================
    ' This sub is used to open the common dialog box for opening
    ' the database, or switching databases.
    '
    ' The first thing it does is get the current path from the
    ' registry setting.  It then sets up the default properties
    ' for the common dialog box.  It sets the filter to show only
    ' .mdb and .dat files, and then it saves the new selection in
    ' in the registry.  Finally, it reloads the form.
    '============================================================
    Dim sFile As String
    sFile = GetSetting("AD&D Tools", "Settings", "Path")
    With dlgOpen
        .DialogTitle = "Open Database"
        .CancelError = False
        .Filter = "Database Files (*.dat,*.mdb)|*.dat;*.mdb|"
        .Filter = .Filter + "Access Databases (*.mdb)|*.mdb|"
        .Filter = .Filter + "Dat Files (*.dat)|*.dat|"
        .Filter = .Filter + "All Files (*.*)|*.*"
        .InitDir = sFile
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SaveSetting "AD&D Tools", "Settings", "Path", sFile
End Sub

Private Sub mnuFileExit_Click()
    '============================================================
    ' Closes the program
    '============================================================
    Unload Me
    End
End Sub

Private Sub mnuEditSettings_Click()
    '============================================================
    ' Opens the settings form.
    '============================================================
    Dim IconSize
    IconSize = GetSetting("AD&D Tools", "Settings", "IconSize")
    Select Case IconSize
        Case 0
            frmSettings.optLarge.Value = True
            frmSettings.optSmall.Value = False
        Case 1
            frmSettings.optSmall.Value = True
            frmSettings.optLarge.Value = False
    End Select
    frmSettings.Show
End Sub

Private Sub mnuFilePrintChar_Click()
    '============================================================
    ' Prints Character Sheets
    '============================================================
    rptCharacter.Show
End Sub

Private Sub mnuFilePrintThief_Click()
    '============================================================
    ' Prints Character's Rogue Skills
    '============================================================
    rptThief.Show
End Sub

Private Sub mnuFilePrintTurn_Click()
    '============================================================
    ' Prints Character's Stats in Turning the Undead
    '============================================================
    rptTurning.Show
End Sub

Private Sub mnuFileProperties_Click()
    '============================================================
    ' Display Database Properties
    '============================================================
    frmFileInfo.Show
End Sub

Private Sub mnuToolsDice_Click()
    '============================================================
    ' Opens the Dice tool
    '============================================================
    frmDice.Show
End Sub

Private Sub mnuToolsPCView_Click()
    '============================================================
    ' Show the PC character view window.
    '============================================================
    frmPCView.Show
End Sub

Private Sub mnuToolsPCAdd_Click()
    '============================================================
    ' Add a new character to the database
    '============================================================
    frmPChar.Show
End Sub

Private Sub mnuToolsNPCView_Click()
    '============================================================
    ' Show the NPC character view window.
    '============================================================
    frmNPChar.Show
    frmNPChar.cmdGenerate_Click
End Sub

Private Sub mnuToolsNPCAdd_Click()
    '============================================================
    ' Generate a new NPC
    '============================================================
    frmNPChar.Show
End Sub

Private Sub mnuWindowCascade_Click()
    '============================================================
    ' Arranges the open windows in a cascaded fashion
    '============================================================
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileV_Click()
    '============================================================
    ' Tiles the open windows in a vertical view
    '============================================================
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileH_Click()
    '============================================================
    ' Tiles the open windows in a horizontal view
    '============================================================
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowArrange_Click()
    '============================================================
    ' Arranges the window icons
    '============================================================
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuHelpContents_Click()
    '============================================================
    ' Opens the Help file
    '============================================================
    Dim nRun
    nRun = Shell("hh.exe " & App.Path & "\ad&dtk_help.chm", vbMaximizedFocus)
End Sub

Private Sub mnuHelpUpdates_Click()
    '============================================================
    ' Displays the Updates.txt file
    '============================================================
    frmReadMe.GetText "Update"
    frmReadMe.Show
    frmReadMe.GetText "Update"
End Sub

Private Sub mnuHelpReadMe_Click()
    '============================================================
    ' Displays the ReadMe.txt file
    '============================================================
    frmReadMe.GetText "ReadMe"
    frmReadMe.Show
    frmReadMe.GetText "ReadMe"
End Sub

Private Sub mnuHelpRegister_Click()
    '============================================================
    ' Displays the Registration information.
    '============================================================
    frmRegister.txtName.Text = GetSetting("AD&D Tools", "Info", "User")
    frmRegister.txtCompany.Text = GetSetting("AD&D Tools", "Info", "Company")
    frmRegister.txtEmail.Text = GetSetting("AD&D Tools", "Info", "Email")
    frmRegister.txtRegCode.Text = GetSetting("AD&D Tools", "Info", "RegID")
    frmRegister.Show
End Sub

Private Sub mnuHelpAbout_Click()
    '============================================================
    ' Shows the About Form
    '============================================================
    frmAbout.Show
End Sub

Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
    '============================================================
    ' This redirects the click of a toolbar button to preform the
    ' same function from the menu.
    '============================================================
    On Error Resume Next
    Select Case Button.Key
        Case "Open"
            mnuFileOpen_Click
        Case "Settings"
            mnuEditSettings_Click
        Case "Exit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrintChar_Click
        Case "Dice"
            mnuToolsDice_Click
        Case "ViewPC"
            mnuToolsPCView_Click
        Case "ViewNPC"
            mnuToolsNPCView_Click
        Case "AddPC"
            mnuToolsPCAdd_Click
        Case "AddNPC"
            mnuToolsNPCAdd_Click
        Case "Cascade"
            mnuWindowCascade_Click
        Case "TileH"
            mnuWindowTileH_Click
        Case "TileV"
            mnuWindowTileV_Click
        Case "Icons"
            mnuWindowArrange_Click
        Case "Updates"
            mnuHelpUpdates_Click
        Case "Readme"
            mnuHelpReadMe_Click
        Case "Help"
            mnuHelpContents_Click
        Case "About"
            mnuHelpAbout_Click
    End Select
End Sub

Private Sub tbTools_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    '============================================================
    ' This is used for the print dropdown menu from the print
    ' button.  It preforms the menu action that corresponds to
    ' the button function.
    '============================================================
    Select Case ButtonMenu.Key
        Case "Character"
            mnuFilePrintChar_Click
        Case "Turning"
            mnuFilePrintTurn_Click
        Case "Thief"
            mnuFilePrintThief_Click
    End Select
End Sub

Public Function RefreshForm()
    '============================================================
    ' This refreshes the icons and the count of the database
    ' records
    '============================================================
    Dim ToolSize, Count
    ToolSize = GetSetting("AD&D Tools", "Settings", "IconSize")
    Select Case ToolSize
        Case 0
            tbTools.ImageList = imgLarge
            SetIcons
        Case 1
            tbTools.ImageList = imgSmall
            SetIcons
    End Select
    
    rs.MoveFirst
    Do Until rs.EOF
        Count = Count + 1
        rs.MoveNext
    Loop

    If Count > 1 Then
        sbInfo.Panels(1).Text = "There are currently " & Count & " characters in the database."
    Else
        sbInfo.Panels(1).Text = "There is currently " & Count & " character in the database."
    End If
End Function

Public Function SetIcons()
    '============================================================
    ' Displays the icons on the toolbar
    '============================================================
    tbTools.Buttons.Item(1).Image = 1
    tbTools.Buttons.Item(2).Image = 2
    tbTools.Buttons.Item(3).Image = 3
    tbTools.Buttons.Item(5).Image = 4
    tbTools.Buttons.Item(7).Image = 5
    tbTools.Buttons.Item(9).Image = 6
    tbTools.Buttons.Item(10).Image = 8
    tbTools.Buttons.Item(12).Image = 7
    tbTools.Buttons.Item(13).Image = 9
    tbTools.Buttons.Item(15).Image = 11
    tbTools.Buttons.Item(16).Image = 12
    tbTools.Buttons.Item(17).Image = 13
    tbTools.Buttons.Item(18).Image = 14
    tbTools.Buttons.Item(20).Image = 17
    tbTools.Buttons.Item(21).Image = 17
    tbTools.Buttons.Item(23).Image = 15
    tbTools.Buttons.Item(24).Image = 16
End Function
