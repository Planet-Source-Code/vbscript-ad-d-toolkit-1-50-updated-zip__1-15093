VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Character Information"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmPChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   7125
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   4800
      TabIndex        =   106
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   2400
      TabIndex        =   54
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3600
      TabIndex        =   55
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   56
      Top             =   4200
      Width           =   1095
   End
   Begin TabDlg.SSTab tabPC 
      Height          =   4095
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Personal Information"
      TabPicture(0)   =   "frmPChar.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRace"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblClass"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLevel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAlign"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblAge"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblHP"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPlayer"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLanguage"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblAC"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtLevel"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAge"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtHP"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtPlayer"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAC"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbRace"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmbClass"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmbAlignment"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmbLang"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "S&tatistics"
      TabPicture(1)   =   "frmPChar.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1"
      Tab(1).Control(1)=   "lblStrAdj"
      Tab(1).Control(2)=   "lblInt"
      Tab(1).Control(3)=   "lblDex"
      Tab(1).Control(4)=   "lvlWis"
      Tab(1).Control(5)=   "lblCon"
      Tab(1).Control(6)=   "lblCha"
      Tab(1).Control(7)=   "lblBreath"
      Tab(1).Control(8)=   "lblRod"
      Tab(1).Control(9)=   "lblSpell"
      Tab(1).Control(10)=   "lblPetr"
      Tab(1).Control(11)=   "lblPara"
      Tab(1).Control(12)=   "txtStr"
      Tab(1).Control(13)=   "txtStrAdj"
      Tab(1).Control(14)=   "txtInt"
      Tab(1).Control(15)=   "txtDex"
      Tab(1).Control(16)=   "txtWis"
      Tab(1).Control(17)=   "txtCon"
      Tab(1).Control(18)=   "txtCha"
      Tab(1).Control(19)=   "txtBreath"
      Tab(1).Control(20)=   "txtRod"
      Tab(1).Control(21)=   "txtSpell"
      Tab(1).Control(22)=   "txtPetr"
      Tab(1).Control(23)=   "txtPara"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "&Inventory"
      TabPicture(2)   =   "frmPChar.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblWeapons"
      Tab(2).Control(1)=   "lblShield"
      Tab(2).Control(2)=   "lblHelm"
      Tab(2).Control(3)=   "lblArmor"
      Tab(2).Control(4)=   "lblMagic"
      Tab(2).Control(5)=   "lblNotes"
      Tab(2).Control(6)=   "txtArmor"
      Tab(2).Control(7)=   "txtHelm"
      Tab(2).Control(8)=   "txtShield"
      Tab(2).Control(9)=   "txtWeapon1"
      Tab(2).Control(10)=   "txtWeapon2"
      Tab(2).Control(11)=   "txtMagic1"
      Tab(2).Control(12)=   "txtMagic2"
      Tab(2).Control(13)=   "txtMagic3"
      Tab(2).Control(14)=   "txtMagic4"
      Tab(2).Control(15)=   "txtMagic5"
      Tab(2).Control(16)=   "txtNotes"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "&Rogue Skills"
      TabPicture(3)   =   "frmPChar.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblClimb"
      Tab(3).Control(1)=   "lblTraps"
      Tab(3).Control(2)=   "lblHide"
      Tab(3).Control(3)=   "lblMove"
      Tab(3).Control(4)=   "lblRead"
      Tab(3).Control(5)=   "lblLocks"
      Tab(3).Control(6)=   "lblPockets"
      Tab(3).Control(7)=   "lblDetect"
      Tab(3).Control(8)=   "txtClimb"
      Tab(3).Control(9)=   "txtNoise"
      Tab(3).Control(10)=   "txtTraps"
      Tab(3).Control(11)=   "txtHide"
      Tab(3).Control(12)=   "txtSilent"
      Tab(3).Control(13)=   "txtLocks"
      Tab(3).Control(14)=   "txtPockets"
      Tab(3).Control(15)=   "txtRead"
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "&Turning Undead"
      TabPicture(4)   =   "frmPChar.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblSkeleton"
      Tab(4).Control(1)=   "lblZombie"
      Tab(4).Control(2)=   "lblGhoul"
      Tab(4).Control(3)=   "lblShaodw"
      Tab(4).Control(4)=   "lblWight"
      Tab(4).Control(5)=   "lblGhast"
      Tab(4).Control(6)=   "lblWraith"
      Tab(4).Control(7)=   "lblMummy"
      Tab(4).Control(8)=   "lblSpectre"
      Tab(4).Control(9)=   "lblVampire"
      Tab(4).Control(10)=   "lblGhost"
      Tab(4).Control(11)=   "lblLiche"
      Tab(4).Control(12)=   "lblSpecial"
      Tab(4).Control(13)=   "txtSkeleton"
      Tab(4).Control(14)=   "txtZombie"
      Tab(4).Control(15)=   "txtGhoul"
      Tab(4).Control(16)=   "txtShadow"
      Tab(4).Control(17)=   "txtWight"
      Tab(4).Control(18)=   "txtGhast"
      Tab(4).Control(19)=   "txtWraith"
      Tab(4).Control(20)=   "txtMummy"
      Tab(4).Control(21)=   "txtSpectre"
      Tab(4).Control(22)=   "txtVampire"
      Tab(4).Control(23)=   "txtGhost"
      Tab(4).Control(24)=   "txtLiche"
      Tab(4).Control(25)=   "txtSpecial"
      Tab(4).ControlCount=   26
      Begin VB.TextBox txtSpecial 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         TabIndex        =   53
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtLiche 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         TabIndex        =   52
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtGhost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         TabIndex        =   51
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtVampire 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         TabIndex        =   50
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtSpectre 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         TabIndex        =   49
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtMummy 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         TabIndex        =   48
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtWraith 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         TabIndex        =   47
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtGhast 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         TabIndex        =   46
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtWight 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         TabIndex        =   45
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtShadow 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         TabIndex        =   44
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtGhoul 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         TabIndex        =   43
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtZombie 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         TabIndex        =   42
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtSkeleton 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         TabIndex        =   41
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtRead 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         TabIndex        =   40
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtPockets 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         TabIndex        =   39
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtLocks 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         TabIndex        =   38
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtSilent 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         TabIndex        =   37
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtHide 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   36
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtTraps 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   35
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtNoise 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   34
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtClimb 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmbLang 
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Top             =   1140
         Width           =   1935
      End
      Begin VB.ComboBox cmbAlignment 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   2100
         Width           =   1935
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1620
         Width           =   1935
      End
      Begin VB.ComboBox cmbRace 
         Height          =   315
         ItemData        =   "frmPChar.frx":0396
         Left            =   1440
         List            =   "frmPChar.frx":0398
         TabIndex        =   2
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtNotes 
         Height          =   1155
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2820
         Width           =   6855
      End
      Begin VB.TextBox txtMagic5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         TabIndex        =   31
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox txtMagic4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         TabIndex        =   30
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox txtMagic3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         TabIndex        =   29
         Top             =   1380
         Width           =   1815
      End
      Begin VB.TextBox txtMagic2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         TabIndex        =   28
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtMagic1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         TabIndex        =   27
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox txtWeapon2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   26
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox txtWeapon1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   25
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox txtShield 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   24
         Top             =   1380
         Width           =   1815
      End
      Begin VB.TextBox txtHelm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   23
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtArmor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         TabIndex        =   22
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox txtPara 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         TabIndex        =   21
         Top             =   2580
         Width           =   495
      End
      Begin VB.TextBox txtPetr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         TabIndex        =   20
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox txtSpell 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         TabIndex        =   19
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox txtRod 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         TabIndex        =   18
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox txtBreath 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         TabIndex        =   17
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtCha 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   16
         Top             =   3060
         Width           =   495
      End
      Begin VB.TextBox txtCon 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   15
         Top             =   2580
         Width           =   495
      End
      Begin VB.TextBox txtWis 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   14
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox txtDex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   13
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   12
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox txtStrAdj 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72960
         TabIndex        =   11
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtStr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         TabIndex        =   10
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtAC 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   1
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtHP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         Top             =   2100
         Width           =   1935
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   1620
         Width           =   1935
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label lblSpecial 
         Caption         =   "Special:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   105
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblLiche 
         Caption         =   "Liche:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   104
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblGhost 
         Caption         =   "Ghost"
         Height          =   255
         Left            =   -70680
         TabIndex        =   103
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblVampire 
         Caption         =   "Vampire:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   102
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblSpectre 
         Caption         =   "Spectre:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   101
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblMummy 
         Caption         =   "Mummy:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   100
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblWraith 
         Caption         =   "Wraith:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   99
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblGhast 
         Caption         =   "Ghast:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   98
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblWight 
         Caption         =   "Wight:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   97
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblShaodw 
         Caption         =   "Shadow:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   96
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblGhoul 
         Caption         =   "Ghoul:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   95
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblZombie 
         Caption         =   "Zombie:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   94
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblSkeleton 
         Caption         =   "Skeleton:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   93
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblDetect 
         Caption         =   "Detect Noise:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblPockets 
         Caption         =   "Pick Pockets:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   91
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblLocks 
         Caption         =   "Open Locks:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   90
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblRead 
         Caption         =   "Read Languages:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   89
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblMove 
         Caption         =   "Move Silently:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   88
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblHide 
         Caption         =   "Hide In Shadows:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   87
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblTraps 
         Caption         =   "Find/Remove Traps:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   86
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblClimb 
         Caption         =   "Climb Walls:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   85
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblNotes 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Character History and Notes"
         Height          =   255
         Left            =   -74880
         TabIndex        =   84
         Top             =   2580
         Width           =   6855
      End
      Begin VB.Label lblMagic 
         Caption         =   "Magic Items:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   83
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblArmor 
         Caption         =   "Armor:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   82
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblHelm 
         Caption         =   "Helm:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   81
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lblShield 
         Caption         =   "Shield:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   80
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label lblWeapons 
         Caption         =   "Weapons:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label lblPara 
         Caption         =   "Paralyzation:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   78
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label lblPetr 
         Caption         =   "Petrification:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   77
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label lblSpell 
         Caption         =   "Spell:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   76
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblRod 
         Caption         =   "Rod:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   75
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblBreath 
         Caption         =   "Breath"
         Height          =   255
         Left            =   -71280
         TabIndex        =   74
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblCha 
         Caption         =   "Charisma:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   73
         Top             =   3060
         Width           =   975
      End
      Begin VB.Label lblCon 
         Caption         =   "Constitution:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   72
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label lvlWis 
         Caption         =   "Wisdom:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   71
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label lblDex 
         Caption         =   "Dexterity:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   70
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblInt 
         Caption         =   "Intelligence:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblStrAdj 
         Caption         =   "Strength/Adj:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   660
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   -73080
         X2              =   -73200
         Y1              =   660
         Y2              =   900
      End
      Begin VB.Label lblAC 
         AutoSize        =   -1  'True
         Caption         =   "Armor Class:"
         Height          =   195
         Left            =   3720
         TabIndex        =   67
         Top             =   2580
         Width           =   870
      End
      Begin VB.Label lblLanguage 
         AutoSize        =   -1  'True
         Caption         =   "Main Language:"
         Height          =   195
         Left            =   3720
         TabIndex        =   66
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblPlayer 
         Caption         =   "Player Name:"
         Height          =   255
         Left            =   3720
         TabIndex        =   65
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         Caption         =   "Hit Points:"
         Height          =   195
         Left            =   3720
         TabIndex        =   64
         Top             =   2100
         Width           =   720
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   195
         Left            =   3720
         TabIndex        =   63
         Top             =   1620
         Width           =   330
      End
      Begin VB.Label lblAlign 
         AutoSize        =   -1  'True
         Caption         =   "Alignment:"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   2580
         Width           =   435
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "Class:"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label lblRace 
         AutoSize        =   -1  'True
         Caption         =   "Race:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1140
         Width           =   435
      End
      Begin VB.Label lblName 
         Caption         =   "Character Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   660
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================================
' This is where the Player Character is generated and saved.
' This section is for the global variables used throughout
' the form
'============================================================
Dim db As Database
Dim rsAlign As Recordset
Dim rsClass As Recordset
Dim rsRace As Recordset
Dim rsLang As Recordset
Dim rsCharInfo As Recordset
Dim rsCharThief As Recordset
Dim rsCharTurn As Recordset

Private Sub Form_Load()
    '============================================================
    ' This runs when the form loads.  It opens the database and
    ' appropriate tables.  It then populates the drop-down lists
    ' with the data from the database.
    '============================================================
    Set db = OpenDatabase(DBPath)
    Set rsAlign = db.OpenRecordset("tblInfoAlign", dbOpenDynaset)
    Set rsClass = db.OpenRecordset("tblInfoClass", dbOpenDynaset)
    Set rsRace = db.OpenRecordset("tblInfoRace", dbOpenDynaset)
    Set rsLang = db.OpenRecordset("tblInfoLang", dbOpenDynaset)
    Set rsCharInfo = db.OpenRecordset("tblCharInfo", dbOpenDynaset)
    Set rsCharThief = db.OpenRecordset("tblCharThief", dbOpenDynaset)
    Set rsCharTurn = db.OpenRecordset("tblCharTurn", dbOpenDynaset)
    
    rsAlign.MoveFirst
    Do While Not rsAlign.EOF
        cmbAlignment.AddItem (rsAlign.Fields("Alignment"))
        rsAlign.MoveNext
    Loop
    rsClass.MoveFirst
    Do While Not rsClass.EOF
        cmbClass.AddItem (rsClass.Fields("Class"))
        rsClass.MoveNext
    Loop
    rsRace.MoveFirst
    Do While Not rsRace.EOF
        cmbRace.AddItem (rsRace.Fields("Race"))
        rsRace.MoveNext
    Loop
    rsLang.MoveFirst
    Do While Not rsLang.EOF
        cmbLang.AddItem (rsLang.Fields("Language"))
        rsLang.MoveNext
    Loop
    tabPC.TabEnabled(3) = False
    tabPC.TabEnabled(4) = False
    ClearData
End Sub

Private Sub cmdGenerate_Click()
    '============================================================
    ' This is where the character's stats are generated.  It
    ' takes a random number from 1 to how ever many entries there
    ' are in the table.
    '============================================================
    Dim RandNum As Integer
    
    Randomize
    ClearData
    
    RandNum = Int(Rnd() * 29) + 1
    txtLevel.Text = RandNum
    
    RandNum = Int(Rnd() * 6)
    cmbRace.Text = cmbRace.List(RandNum)
    
    RandNum = Int(Rnd() * 9)
    cmbAlignment.Text = cmbAlignment.List(RandNum)
    
    RandNum = Int(Rnd() * 22)
    cmbClass.Text = cmbClass.List(RandNum)
    
    txtAC.Text = 10
End Sub

Private Sub cmdSave_Click()
    '============================================================
    ' This saves the information to a new entry in the database.
    '============================================================
    With rsCharInfo
        .AddNew
            !CharName = txtName.Text
            !Player = txtPlayer.Text
            !Class = cmbClass.Text
            !Race = cmbRace.Text
            !Language = cmbLang.Text
            !Alignment = cmbAlignment.Text
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
    If tabPC.TabEnabled(3) = True Then
        With rsCharThief
            .AddNew
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
    If tabPC.TabEnabled(4) = True Then
        With rsCharTurn
            .AddNew
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
    ClearData
End Sub

Private Sub cmdClear_Click()
    '============================================================
    ' Clears the data, both generated and custom.
    '============================================================
    ClearData
End Sub

Private Sub cmdExit_Click()
    '============================================================
    ' Exits the form
    '============================================================
    Unload Me
End Sub

Private Sub cmbRace_Change()
    '============================================================
    ' Calls the race check whenever the race is changed
    '============================================================
    cmbRace_Click
End Sub

Private Sub cmbRace_Click()
    '============================================================
    ' This sets up the language and age values based on race.
    '============================================================
    Select Case cmbRace.Text
        Case "Human"
            cmbLang.Text = "Regional Human"
            txtAge.Text = 15 + RollDice(1, 4)
        Case "Elf"
            cmbLang.Text = "Elven"
            txtAge.Text = RollDice(5, 6) + 100
        Case "Half-Elf"
            cmbLang.Text = "Common"
            txtAge.Text = RollDice(1, 6) + 15
        Case "Halfling"
            cmbLang.Text = "Halfling"
            txtAge.Text = RollDice(3, 4) + 20
        Case "Dwarf"
            cmbLang.Text = "Dwarven"
            txtAge.Text = RollDice(5, 6) + 40
        Case "Gnome"
            cmbLang.Text = "Gnomish"
            txtAge.Text = RollDice(3, 12) + 60
    End Select
End Sub

Private Sub cmbClass_Change()
    '============================================================
    ' Calls the Class check whenever the class is changed.
    '============================================================
    cmbClass_Click
End Sub

Private Sub cmbClass_Click()
    '============================================================
    ' This sets up the constraints on the character's stats based
    ' on class.  Examples: sets statistics and minimums, gets
    ' saving throws based on level/class, Hit Points, and
    ' race/alignment restrictions.  It also enables/disables the
    ' tabs for Rogue Skils and Turning Undead for appropriate
    ' classes.
    '============================================================
    Select Case cmbClass.Text
        Case "Fighter"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            GetHP
            GetSaving Me
        Case "Paladin"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Text = "Lawful Good"
            cmbAlignment.Enabled = False
            cmbRace.Text = "Human"
            cmbRace.Enabled = False
            cmbLang.Text = "Regional Human"
            GetStats Me
            If txtStr.Text < 12 Then txtStr.Text = 12
            If txtCon.Text < 9 Then txtCon.Text = 9
            If txtWis.Text < 13 Then txtWis.Text = 13
            If txtCha.Text < 17 Then txtCha.Text = 17
            GetTurning Me
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Ranger"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 13 Then txtStr.Text = 13
            If txtDex.Text < 13 Then txtDex.Text = 13
            If txtCon.Text < 14 Then txtCon.Text = 14
            If txtWis.Text < 14 Then txtWis.Text = 14
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Mage"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            txtHP.Text = RollDice(1, 4)
        Case "Illusionist"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            txtHP.Text = RollDice(1, 4)
            GetSaving Me
        Case "Cleric"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtWis.Text < 9 Then txtWis.Text = 9
            GetTurning Me
            txtHP.Text = RollDice(1, 8)
            GetSaving Me
        Case "Druid"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtWis.Text < 12 Then txtWis.Text = 12
            If txtCha.Text < 15 Then txtCha.Text = 15
            GetTurning Me
            txtHP.Text = RollDice(1, 8)
            GetSaving Me
        Case "Bard"
            tabPC.TabEnabled(3) = True
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtDex.Text < 12 Then txtDex.Text = 12
            If txtInt.Text < 13 Then txtInt.Text = 13
            If txtCha.Text < 15 Then txtCha.Text = 15
            GetThief Me
            txtHP.Text = RollDice(1, 6)
            GetSaving Me
        Case "Thief"
            tabPC.TabEnabled(3) = True
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtDex.Text < 9 Then txtDex.Text = 9
            GetThief Me
            txtHP.Text = RollDice(1, 6)
            GetSaving Me
        Case "Fighter/Thief"
            tabPC.TabEnabled(3) = True
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            If txtDex.Text < 9 Then txtDex.Text = 9
            GetThief Me
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Fighter/Cleric"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            If txtWis.Text < 9 Then txtWis.Text = 9
            GetTurning Me
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Fighter/Mage"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            If txtInt.Text < 9 Then txtInt.Text = 9
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Fighter/Illusionist"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            If txtInt.Text < 9 Then txtInt.Text = 9
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Cleric/Ranger"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 13 Then txtStr.Text = 13
            If txtDex.Text < 13 Then txtDex.Text = 13
            If txtCon.Text < 14 Then txtCon.Text = 14
            If txtWis.Text < 14 Then txtWis.Text = 14
            GetTurning Me
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Cleric/Mage"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            If txtWis.Text < 9 Then txtWis.Text = 9
            GetTurning Me
            txtHP.Text = RollDice(1, 8)
            GetSaving Me
        Case "Cleric/Illusionist"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            If txtWis.Text < 9 Then txtWis.Text = 9
            GetTurning Me
            txtHP.Text = RollDice(1, 8)
            GetSaving Me
        Case "Cleric/Druid"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtWis.Text < 12 Then txtWis.Text = 12
            If txtCha.Text < 15 Then txtCha.Text = 15
            GetTurning Me
            txtHP.Text = RollDice(1, 8)
            GetSaving Me
        Case "Cleric/Thief"
            tabPC.TabEnabled(3) = True
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtWis.Text < 9 Then txtWis.Text = 9
            If txtDex.Text < 9 Then txtDex.Text = 9
            GetTurning Me
            GetThief Me
            txtHP.Text = RollDice(1, 8)
            GetSaving Me
        Case "Mage/Thief"
            tabPC.TabEnabled(3) = True
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            If txtDex.Text < 9 Then txtDex.Text = 9
            GetThief Me
            txtHP.Text = RollDice(1, 6)
            GetSaving Me
        Case "Illusionist/Thief"
            tabPC.TabEnabled(3) = True
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            If txtDex.Text < 9 Then txtDex.Text = 9
            GetThief Me
            txtHP.Text = RollDice(1, 6)
            GetSaving Me
        Case "Fighter/Mage/Cleric"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = True
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            If txtInt.Text < 9 Then txtInt.Text = 9
            If txtWis.Text < 9 Then txtWis.Text = 9
            GetTurning Me
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case "Fighter/Mage/Thief"
            tabPC.TabEnabled(3) = True
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            If txtInt.Text < 9 Then txtInt.Text = 9
            If txtDex.Text < 9 Then txtDex.Text = 9
            GetThief Me
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
        Case Else
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
    End Select
End Sub

Public Sub ClearData()
    '============================================================
    ' This resets the form and clears all data.
    '============================================================
    txtName.Text = ""
    txtPlayer.Text = GetSetting("AD&D Tools", "Info", "User")
    cmbClass.Text = ""
    cmbRace.Text = ""
    cmbLang.Text = ""
    cmbAlignment.Text = ""
    txtLevel.Text = ""
    txtAge.Text = ""
    txtHP.Text = ""
    txtAC.Text = ""
    txtStr.Text = ""
    txtStrAdj.Text = ""
    txtInt.Text = ""
    txtDex.Text = ""
    txtWis.Text = ""
    txtCon.Text = ""
    txtCha.Text = ""
    txtBreath.Text = ""
    txtRod.Text = ""
    txtSpell.Text = ""
    txtPetr.Text = ""
    txtPara.Text = ""
    txtArmor.Text = ""
    txtHelm.Text = ""
    txtShield.Text = ""
    txtWeapon1.Text = ""
    txtWeapon2.Text = ""
    txtMagic1.Text = ""
    txtMagic2.Text = ""
    txtMagic3.Text = ""
    txtMagic4.Text = ""
    txtMagic5.Text = ""
    txtNotes.Text = ""
    txtSkeleton.Text = ""
    txtZombie.Text = ""
    txtGhoul.Text = ""
    txtShadow.Text = ""
    txtWight.Text = ""
    txtGhast.Text = ""
    txtWraith.Text = ""
    txtMummy.Text = ""
    txtSpectre.Text = ""
    txtVampire.Text = ""
    txtGhost.Text = ""
    txtLiche.Text = ""
    txtSpecial.Text = ""
    txtClimb.Text = ""
    txtNoise.Text = ""
    txtTraps.Text = ""
    txtHide.Text = ""
    txtSilent.Text = ""
    txtLocks.Text = ""
    txtPockets.Text = ""
    txtRead.Text = ""
    tabPC.TabEnabled(3) = False
    tabPC.TabEnabled(4) = False
    cmbAlignment.Enabled = True
    cmbRace.Enabled = True
End Sub

Public Sub GetHP()
    txtHP.Text = RollDice(1, 10) * txtLevel.Text
End Sub
