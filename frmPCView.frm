VERSION 5.00
Begin VB.Form frmPCView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Character Viewer/Editor"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmPCView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   7815
   Begin VB.CommandButton cmdReset 
      Caption         =   "&r"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      ToolTipText     =   "Reset Information"
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&s"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      ToolTipText     =   "Save/Update Character"
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&c"
      Default         =   -1  'True
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      ToolTipText     =   "Close Character Viewer"
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Go to First Record"
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Go to Previous Record"
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Go To Next record"
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Go To Last Record"
      Top             =   7680
      Width           =   375
   End
   Begin VB.Frame frameTurn 
      Caption         =   "Undead Turning"
      Height          =   1095
      Left            =   120
      TabIndex        =   67
      Top             =   6480
      Width           =   7575
      Begin VB.TextBox txtVampire 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4680
         TabIndex        =   56
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtWraith 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   53
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtShadow 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSkeleton 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   47
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSpecial 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6000
         TabIndex        =   59
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtZombie 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   48
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtGhoul 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   49
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtWight 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   51
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtGhast 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   52
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtMummy 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   54
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtSpectre 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtGhost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4680
         TabIndex        =   57
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtLiche 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4680
         TabIndex        =   58
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSpecial 
         Caption         =   "Special:"
         Height          =   255
         Left            =   5280
         TabIndex        =   80
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblVampire 
         Caption         =   "Vampire:"
         Height          =   255
         Left            =   3960
         TabIndex        =   79
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblGhost 
         Caption         =   "Ghost"
         Height          =   255
         Left            =   3960
         TabIndex        =   78
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLiche 
         Caption         =   "Liche:"
         Height          =   255
         Left            =   3960
         TabIndex        =   77
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblWraith 
         Caption         =   "Wraith:"
         Height          =   255
         Left            =   2760
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblMummy 
         Caption         =   "Mummy:"
         Height          =   255
         Left            =   2760
         TabIndex        =   75
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblSpectre 
         Caption         =   "Spectre:"
         Height          =   255
         Left            =   2760
         TabIndex        =   74
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblShaodw 
         Caption         =   "Shadow:"
         Height          =   255
         Left            =   1440
         TabIndex        =   73
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblWight 
         Caption         =   "Wight:"
         Height          =   255
         Left            =   1440
         TabIndex        =   72
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblGhast 
         Caption         =   "Ghast:"
         Height          =   255
         Left            =   1440
         TabIndex        =   71
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblSkeleton 
         Caption         =   "Skeleton:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblZombie 
         Caption         =   "Zombie:"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblGhoul 
         Caption         =   "Ghoul:"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame frameThief 
      Caption         =   "Rogue Skills"
      Height          =   855
      Left            =   120
      TabIndex        =   66
      Top             =   5640
      Width           =   7575
      Begin VB.TextBox txtPockets 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6960
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSilent 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   43
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtTraps 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtClimb 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   39
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtNoise 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   40
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtHide 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   42
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtLocks 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   44
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtRead 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6960
         TabIndex        =   46
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblRead 
         Caption         =   "Read Languages:"
         Height          =   255
         Left            =   5640
         TabIndex        =   88
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblPockets 
         Caption         =   "Pick Pockets:"
         Height          =   255
         Left            =   5640
         TabIndex        =   87
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblMove 
         Caption         =   "Move Silently:"
         Height          =   255
         Left            =   3960
         TabIndex        =   86
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblLocks 
         Caption         =   "Open Locks:"
         Height          =   255
         Left            =   3960
         TabIndex        =   85
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTraps 
         Caption         =   "Find/Remove Traps:"
         Height          =   255
         Left            =   1800
         TabIndex        =   84
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblHide 
         Caption         =   "Hide In Shadows:"
         Height          =   255
         Left            =   1800
         TabIndex        =   83
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblClimb 
         Caption         =   "Climb Walls:"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDetect 
         Caption         =   "Detect Noise:"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame frameNotes 
      Caption         =   "Character History/Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   65
      Top             =   4440
      Width           =   7575
      Begin VB.TextBox txtNotes 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame frameMagic 
      Height          =   1575
      Left            =   3960
      TabIndex        =   64
      Top             =   2880
      Width           =   3735
      Begin VB.TextBox txtMagic1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtMagic2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   34
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtMagic3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtMagic4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtMagic5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblMagic 
         Caption         =   "Magic Items:"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frameItems 
      Height          =   1575
      Left            =   120
      TabIndex        =   63
      Top             =   2880
      Width           =   3735
      Begin VB.TextBox txtArmor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtHelm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtShield 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   30
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtWeapon1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   31
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtWeapon2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   32
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblWeapons 
         Caption         =   "Weapons:"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblShield 
         Caption         =   "Shield:"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblHelm 
         Caption         =   "Helm:"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblArmor 
         Caption         =   "Armor:"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame frameSave 
      Height          =   1335
      Left            =   3960
      TabIndex        =   62
      Top             =   1560
      Width           =   3735
      Begin VB.TextBox txtPetr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPara 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtBreath 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRod 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSpell 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblPetr 
         Caption         =   "Petrification:"
         Height          =   255
         Left            =   2040
         TabIndex        =   104
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPara 
         Caption         =   "Paralyzation:"
         Height          =   255
         Left            =   2040
         TabIndex        =   103
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblBreath 
         Caption         =   "Breath"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRod 
         Caption         =   "Rod:"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblSpell 
         Caption         =   "Spell:"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame frameAbilities 
      Height          =   1335
      Left            =   120
      TabIndex        =   61
      Top             =   1560
      Width           =   3735
      Begin VB.TextBox txtCon 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtStr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtStrAdj 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtDex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtWis 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtCha 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   22
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lvlWis 
         Caption         =   "Wisdom:"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblCon 
         Caption         =   "Constitution:"
         Height          =   255
         Left            =   2160
         TabIndex        =   98
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCha 
         Caption         =   "Charisma:"
         Height          =   255
         Left            =   2160
         TabIndex        =   97
         Top             =   960
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   1800
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Label lblStrAdj 
         Caption         =   "Strength/Adj:"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblInt 
         Caption         =   "Intelligence:"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblDex 
         Caption         =   "Dexterity:"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame framePersonal 
      Height          =   1575
      Left            =   120
      TabIndex        =   60
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtLanguage 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5160
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtHP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5160
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtAC 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5160
         TabIndex        =   15
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtRace 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtClass 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtAlignment 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   115
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblName 
         Caption         =   "Character Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   114
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblRace 
         Caption         =   "Race:"
         Height          =   255
         Left            =   240
         TabIndex        =   113
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblClass 
         Caption         =   "Class:"
         Height          =   255
         Left            =   240
         TabIndex        =   112
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   111
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblAlign 
         Caption         =   "Alignment:"
         Height          =   255
         Left            =   240
         TabIndex        =   110
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblAge 
         Caption         =   "Age:"
         Height          =   255
         Left            =   3840
         TabIndex        =   109
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lblHP 
         Caption         =   "Hit Points:"
         Height          =   255
         Left            =   3840
         TabIndex        =   108
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblPlayer 
         Caption         =   "Player Name:"
         Height          =   255
         Left            =   3840
         TabIndex        =   107
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblLanguage 
         Caption         =   "Main Language:"
         Height          =   255
         Left            =   3840
         TabIndex        =   106
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label lblAC 
         AutoSize        =   -1  'True
         Caption         =   "Armor Class:"
         Height          =   255
         Left            =   3840
         TabIndex        =   105
         Top             =   1200
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmPCView"
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
Dim rsCharInfo As Recordset
Dim rsCharThief As Recordset
Dim rsCharTurn As Recordset

Private Sub Form_Load()
    '============================================================
    ' This section runs when the form loads.  It is responsible
    ' for populating the fields from the database.
    '============================================================
    Set db = OpenDatabase(DBPath)
    Set rsCharInfo = db.OpenRecordset("tblCharInfo", dbOpenDynaset)
    Set rsCharThief = db.OpenRecordset("tblCharThief", dbOpenDynaset)
    Set rsCharTurn = db.OpenRecordset("tblCharTurn", dbOpenDynaset)
    rsCharInfo.MoveFirst
    GetData
End Sub

Private Sub cmdFirst_Click()
    '============================================================
    ' Moves to the first record in the database
    '============================================================
    rsCharInfo.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click()
    '============================================================
    ' Moves to the previous record in the table
    '============================================================
    rsCharInfo.MovePrevious
    GetData
End Sub

Private Sub cmdNext_Click()
    '============================================================
    ' Moves to the next record in the table
    '============================================================
    rsCharInfo.MoveNext
    GetData
End Sub

Private Sub cmdLast_Click()
    '============================================================
    ' Moves to the last record in the table
    '============================================================
    rsCharInfo.MoveLast
    GetData
End Sub

Private Sub cmdReset_Click()
    '============================================================
    ' Resets any changes made before saving.
    '============================================================
    GetData
End Sub
Private Sub cmdSave_Click()
    '============================================================
    ' Saves changes to the character over the existing entry in
    ' the database
    '============================================================
    With rsCharInfo
        .Edit
            !CharName = txtName.Text
            !Player = txtPlayer.Text
            !Class = txtClass.Text
            !Race = txtRace.Text
            !Language = txtLanguage.Text
            !Alignment = txtAlignment.Text
            !Level = txtLevel.Text
            !Age = txtAge.Text
            !HitPoint = txtHP.Text
            !ArmorClass = txtAC.Text
            !Str = txtStr.Text
            !StrAdj = txtStrAdj.Text
            !Int = txtInt.Text
            !Dex = txtDex.Text
            !Wis = txtWis.Text
            !Con = txtCon.Text
            !Cha = txtCha.Text
            !Breath = txtBreath.Text
            !Rod = txtRod.Text
            !Spell = txtSpell.Text
            !Petr = txtPetr.Text
            !Para = txtPara.Text
            !Armor = txtArmor.Text
            !Helm = txtHelm.Text
            !Shield = txtShield.Text
            !Weapon1 = txtWeapon1.Text
            !Weapon2 = txtWeapon2.Text
            !Magic1 = txtMagic1.Text
            !Magic2 = txtMagic2.Text
            !Magic3 = txtMagic3.Text
            !Magic4 = txtMagic4.Text
            !Magic5 = txtMagic5.Text
            !Notes = txtNotes.Text
        .Update
    End With
    If frameThief.Visible = True Then
        With rsCharThief
            .Edit
                !CharName = txtName.Text
                !Climb = txtClimb.Text
                !Detect = txtNoise.Text
                !Find = txtTraps.Text
                !Hide = txtHide.Text
                !Move = txtSilent.Text
                !Open = txtLocks.Text
                !Pick = txtPockets.Text
                !Read = txtRead.Text
            .Update
        End With
    End If
    If frameTurn.Visible = True Then
        With rsCharTurn
            .Edit
                !CharName = txtName.Text
                !Skeleton = txtSkeleton.Text
                !Zombie = txtZombie.Text
                !Ghoul = txtGhoul.Text
                !Shadow = txtShadow.Text
                !Wight = txtWight.Text
                !Ghast = txtGhast.Text
                !Wraith = txtWraith.Text
                !Mummy = txtMummy.Text
                !Spectre = txtSpectre.Text
                !Vampire = txtVampire.Text
                !Ghost = txtGhost.Text
                !Liche = txtLiche.Text
                !Special = txtSpecial.Text
            .Update
        End With
    End If
    GetData
End Sub

Private Sub cmdClose_Click()
    '============================================================
    ' Closes the form
    '============================================================
    Unload Me
End Sub

Public Sub GetData()
    '============================================================
    ' Fills in the data from the character database
    '============================================================
    On Error Resume Next
    txtName.Text = rsCharInfo.Fields("CharName")
    txtPlayer.Text = rsCharInfo.Fields("Player")
    txtRace.Text = rsCharInfo.Fields("Race")
    txtClass.Text = rsCharInfo.Fields("Class")
    CheckClass
    txtAlignment.Text = rsCharInfo.Fields("Alignment")
    txtLanguage.Text = rsCharInfo.Fields("Language")
    txtLevel.Text = rsCharInfo.Fields("Level")
    txtAge.Text = rsCharInfo.Fields("Age")
    txtHP.Text = rsCharInfo.Fields("HitPoint")
    txtAC.Text = rsCharInfo.Fields("ArmorClass")
    txtStr.Text = rsCharInfo.Fields("Str")
    txtStrAdj.Text = rsCharInfo.Fields("StrAdj")
    txtInt.Text = rsCharInfo.Fields("Int")
    txtDex.Text = rsCharInfo.Fields("Dex")
    txtWis.Text = rsCharInfo.Fields("Wis")
    txtCon.Text = rsCharInfo.Fields("Con")
    txtCha.Text = rsCharInfo.Fields("Cha")
    txtBreath.Text = rsCharInfo.Fields("Breath")
    txtRod.Text = rsCharInfo.Fields("Rod")
    txtSpell.Text = rsCharInfo.Fields("Spell")
    txtPara.Text = rsCharInfo.Fields("Para")
    txtPetr.Text = rsCharInfo.Fields("Petr")
    txtArmor.Text = rsCharInfo.Fields("Armor")
    txtHelm.Text = rsCharInfo.Fields("Helm")
    txtShield.Text = rsCharInfo.Fields("Shield")
    txtWeapon1.Text = rsCharInfo.Fields("Weapon1")
    txtWeapon2.Text = rsCharInfo.Fields("Weapon2")
    txtMagic1.Text = rsCharInfo.Fields("Magic1")
    txtMagic2.Text = rsCharInfo.Fields("Magic2")
    txtMagic3.Text = rsCharInfo.Fields("Magic3")
    txtMagic4.Text = rsCharInfo.Fields("Magic4")
    txtMagic5.Text = rsCharInfo.Fields("Magic5")
    txtNotes.Text = rsCharInfo.Fields("Notes")
End Sub

Public Function CheckClass()
    '============================================================
    ' Checks the character class do determine whether or not to
    ' display the Rogue Skills and/or Turning Stats
    '============================================================
    Select Case txtClass.Text
        Case "Fighter"
            frameThief.Visible = False
            frameTurn.Visible = False
        Case "Paladin"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Ranger"
            frameThief.Visible = False
            frameTurn.Visible = False
        Case "Mage"
            frameThief.Visible = False
            frameTurn.Visible = False
        Case "Illusionist"
            frameThief.Visible = False
            frameTurn.Visible = False
        Case "Cleric"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Druid"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Bard"
            frameThief.Visible = True
            frameTurn.Visible = False
            FillInThief
        Case "Thief"
            frameThief.Visible = True
            frameTurn.Visible = False
            FillInThief
        Case "Fighter/Thief"
            frameThief.Visible = True
            frameTurn.Visible = False
            FillInThief
        Case "Fighter/Cleric"
            frameThief.Visible = False
            frameTurn.Visible = True
        Case "Fighter/Mage"
            frameThief.Visible = False
            frameTurn.Visible = False
        Case "Fighter/Illusionist"
            frameThief.Visible = False
            frameTurn.Visible = False
        Case "Cleric/Ranger"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Cleric/Mage"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Cleric/Illusionist"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Cleric/Druid"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Cleric/Thief"
            frameThief.Visible = True
            frameTurn.Visible = True
            FillInTurn
            FillInThief
        Case "Mage/Thief"
            frameThief.Visible = True
            frameTurn.Visible = False
            FillInThief
        Case "Illusionist/Thief"
            frameThief.Visible = True
            frameTurn.Visible = False
            FillInThief
        Case "Fighter/Mage/Cleric"
            frameThief.Visible = False
            frameTurn.Visible = True
            FillInTurn
        Case "Fighter/Mage/Thief"
            frameThief.Visible = True
            frameTurn.Visible = False
            FillInThief
    End Select
End Function

Public Function FillInTurn()
    '============================================================
    ' Fills in the correct values from the database based on
    ' character name
    '============================================================
    Dim strSearch As String, CharName
    
    CharName = txtName.Text
    strSearch = "[CharName] = '" & CharName & "'"
    
    rsCharTurn.MoveFirst
    
    With rsCharTurn
        .FindFirst strSearch
        txtSkeleton.Text = rsCharTurn.Fields("Skeleton")
        txtZombie.Text = rsCharTurn.Fields("Zombie")
        txtGhoul.Text = rsCharTurn.Fields("Ghoul")
        txtShadow.Text = rsCharTurn.Fields("Shadow")
        txtWight.Text = rsCharTurn.Fields("Wight")
        txtGhast.Text = rsCharTurn.Fields("Ghast")
        txtWraith.Text = rsCharTurn.Fields("Wraith")
        txtMummy.Text = rsCharTurn.Fields("Mummy")
        txtSpectre.Text = rsCharTurn.Fields("Spectre")
        txtVampire.Text = rsCharTurn.Fields("Vampire")
        txtGhost.Text = rsCharTurn.Fields("Ghost")
        txtLiche.Text = rsCharTurn.Fields("Liche")
        txtSpecial.Text = rsCharTurn.Fields("Special")
    End With
End Function

Public Function FillInThief()
    '============================================================
    ' Fills in the correct values from the database based on
    ' character name
    '============================================================
    Dim strSearch As String, CharName
    
    CharName = txtName.Text
    strSearch = "[CharName] = '" & CharName & "'"
    
    rsCharThief.MoveFirst
    
    With rsCharThief
        .FindFirst strSearch
        txtClimb.Text = rsCharThief.Fields("Climb")
        txtNoise.Text = rsCharThief.Fields("Detect")
        txtTraps.Text = rsCharThief.Fields("Find")
        txtHide.Text = rsCharThief.Fields("Hide")
        txtSilent.Text = rsCharThief.Fields("Move")
        txtLocks.Text = rsCharThief.Fields("Open")
        txtPockets.Text = rsCharThief.Fields("Pick")
        txtRead.Text = rsCharThief.Fields("Read")
    End With
End Function
