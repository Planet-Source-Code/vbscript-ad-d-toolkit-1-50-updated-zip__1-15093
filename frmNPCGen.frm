VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNPChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Information"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmNPCGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   7125
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   4800
      TabIndex        =   103
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin TabDlg.SSTab tabPC 
      Height          =   4095
      Left            =   0
      TabIndex        =   2
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
      TabPicture(0)   =   "frmNPCGen.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAC"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLanguage"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblHP"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAge"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAlign"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLevel"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblClass"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblRace"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbLang"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbAlignment"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbClass"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbRace"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAC"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtHP"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAge"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtLevel"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtName"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "S&tatistics"
      TabPicture(1)   =   "frmNPCGen.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtStr"
      Tab(1).Control(1)=   "txtStrAdj"
      Tab(1).Control(2)=   "txtInt"
      Tab(1).Control(3)=   "txtDex"
      Tab(1).Control(4)=   "txtWis"
      Tab(1).Control(5)=   "txtCon"
      Tab(1).Control(6)=   "txtCha"
      Tab(1).Control(7)=   "txtBreath"
      Tab(1).Control(8)=   "txtRod"
      Tab(1).Control(9)=   "txtSpell"
      Tab(1).Control(10)=   "txtPetr"
      Tab(1).Control(11)=   "txtPara"
      Tab(1).Control(12)=   "Line1"
      Tab(1).Control(13)=   "lblStrAdj"
      Tab(1).Control(14)=   "lblInt"
      Tab(1).Control(15)=   "lblDex"
      Tab(1).Control(16)=   "lvlWis"
      Tab(1).Control(17)=   "lblCon"
      Tab(1).Control(18)=   "lblCha"
      Tab(1).Control(19)=   "lblBreath"
      Tab(1).Control(20)=   "lblRod"
      Tab(1).Control(21)=   "lblSpell"
      Tab(1).Control(22)=   "lblPetr"
      Tab(1).Control(23)=   "lblPara"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "&Inventory"
      TabPicture(2)   =   "frmNPCGen.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtArmor"
      Tab(2).Control(1)=   "txtHelm"
      Tab(2).Control(2)=   "txtShield"
      Tab(2).Control(3)=   "txtWeapon1"
      Tab(2).Control(4)=   "txtWeapon2"
      Tab(2).Control(5)=   "txtMagic1"
      Tab(2).Control(6)=   "txtMagic2"
      Tab(2).Control(7)=   "txtMagic3"
      Tab(2).Control(8)=   "txtMagic4"
      Tab(2).Control(9)=   "txtMagic5"
      Tab(2).Control(10)=   "txtNotes"
      Tab(2).Control(11)=   "lblWeapons"
      Tab(2).Control(12)=   "lblShield"
      Tab(2).Control(13)=   "lblHelm"
      Tab(2).Control(14)=   "lblArmor"
      Tab(2).Control(15)=   "lblMagic"
      Tab(2).Control(16)=   "lblNotes"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "&Rogue Skills"
      TabPicture(3)   =   "frmNPCGen.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtClimb"
      Tab(3).Control(1)=   "txtNoise"
      Tab(3).Control(2)=   "txtTraps"
      Tab(3).Control(3)=   "txtHide"
      Tab(3).Control(4)=   "txtSilent"
      Tab(3).Control(5)=   "txtLocks"
      Tab(3).Control(6)=   "txtPockets"
      Tab(3).Control(7)=   "txtRead"
      Tab(3).Control(8)=   "lblClimb"
      Tab(3).Control(9)=   "lblTraps"
      Tab(3).Control(10)=   "lblHide"
      Tab(3).Control(11)=   "lblMove"
      Tab(3).Control(12)=   "lblRead"
      Tab(3).Control(13)=   "lblLocks"
      Tab(3).Control(14)=   "lblPockets"
      Tab(3).Control(15)=   "lblDetect"
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "&Turning Undead"
      TabPicture(4)   =   "frmNPCGen.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtSkeleton"
      Tab(4).Control(1)=   "txtZombie"
      Tab(4).Control(2)=   "txtGhoul"
      Tab(4).Control(3)=   "txtShadow"
      Tab(4).Control(4)=   "txtWight"
      Tab(4).Control(5)=   "txtGhast"
      Tab(4).Control(6)=   "txtWraith"
      Tab(4).Control(7)=   "txtMummy"
      Tab(4).Control(8)=   "txtSpectre"
      Tab(4).Control(9)=   "txtVampire"
      Tab(4).Control(10)=   "txtGhost"
      Tab(4).Control(11)=   "txtLiche"
      Tab(4).Control(12)=   "txtSpecial"
      Tab(4).Control(13)=   "lblSkeleton"
      Tab(4).Control(14)=   "lblZombie"
      Tab(4).Control(15)=   "lblGhoul"
      Tab(4).Control(16)=   "lblShaodw"
      Tab(4).Control(17)=   "lblWight"
      Tab(4).Control(18)=   "lblGhast"
      Tab(4).Control(19)=   "lblWraith"
      Tab(4).Control(20)=   "lblMummy"
      Tab(4).Control(21)=   "lblSpectre"
      Tab(4).Control(22)=   "lblVampire"
      Tab(4).Control(23)=   "lblGhost"
      Tab(4).Control(24)=   "lblLiche"
      Tab(4).Control(25)=   "lblSpecial"
      Tab(4).ControlCount=   26
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   55
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1620
         Width           =   1935
      End
      Begin VB.TextBox txtHP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2100
         Width           =   1935
      End
      Begin VB.TextBox txtAC 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtStr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtStrAdj 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72960
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox txtDex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox txtWis 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox txtCon 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2580
         Width           =   495
      End
      Begin VB.TextBox txtCha 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3060
         Width           =   495
      End
      Begin VB.TextBox txtBreath 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtRod 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox txtSpell 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox txtPetr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox txtPara 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2580
         Width           =   495
      End
      Begin VB.TextBox txtArmor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox txtHelm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtShield 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1380
         Width           =   1815
      End
      Begin VB.TextBox txtWeapon1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox txtWeapon2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox txtMagic1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox txtMagic2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtMagic3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1380
         Width           =   1815
      End
      Begin VB.TextBox txtMagic4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox txtMagic5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox txtNotes 
         Height          =   1155
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2820
         Width           =   6855
      End
      Begin VB.ComboBox cmbRace 
         Height          =   315
         ItemData        =   "frmNPCGen.frx":0956
         Left            =   1440
         List            =   "frmNPCGen.frx":0958
         TabIndex        =   27
         Top             =   1140
         Width           =   1935
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   1440
         TabIndex        =   26
         Top             =   1620
         Width           =   1935
      End
      Begin VB.ComboBox cmbAlignment 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   2100
         Width           =   1935
      End
      Begin VB.ComboBox cmbLang 
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtClimb 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNoise 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtTraps 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtHide 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtSilent 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtLocks 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtPockets 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtRead 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69600
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtSkeleton 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtZombie 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtGhoul 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtShadow 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtWight 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtGhast 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtWraith 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtMummy 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtSpectre 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtVampire 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtGhost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtLiche 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtSpecial 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "Character Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label lblRace 
         AutoSize        =   -1  'True
         Caption         =   "Race:"
         Height          =   195
         Left            =   120
         TabIndex        =   101
         Top             =   1140
         Width           =   435
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "Class:"
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   2580
         Width           =   435
      End
      Begin VB.Label lblAlign 
         AutoSize        =   -1  'True
         Caption         =   "Alignment:"
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   195
         Left            =   3720
         TabIndex        =   97
         Top             =   1620
         Width           =   330
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         Caption         =   "Hit Points:"
         Height          =   195
         Left            =   3720
         TabIndex        =   96
         Top             =   2100
         Width           =   720
      End
      Begin VB.Label lblLanguage 
         AutoSize        =   -1  'True
         Caption         =   "Main Language:"
         Height          =   195
         Left            =   3720
         TabIndex        =   95
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblAC 
         AutoSize        =   -1  'True
         Caption         =   "Armor Class:"
         Height          =   195
         Left            =   3720
         TabIndex        =   94
         Top             =   2580
         Width           =   870
      End
      Begin VB.Line Line1 
         X1              =   -73080
         X2              =   -73200
         Y1              =   660
         Y2              =   900
      End
      Begin VB.Label lblStrAdj 
         Caption         =   "Strength/Adj:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   93
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblInt 
         Caption         =   "Intelligence:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblDex 
         Caption         =   "Dexterity:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   91
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lvlWis 
         Caption         =   "Wisdom:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   90
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label lblCon 
         Caption         =   "Constitution:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label lblCha 
         Caption         =   "Charisma:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   88
         Top             =   3060
         Width           =   975
      End
      Begin VB.Label lblBreath 
         Caption         =   "Breath"
         Height          =   255
         Left            =   -71280
         TabIndex        =   87
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblRod 
         Caption         =   "Rod:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   86
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblSpell 
         Caption         =   "Spell:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   85
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblPetr 
         Caption         =   "Petrification:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   84
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label lblPara 
         Caption         =   "Paralyzation:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   83
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label lblWeapons 
         Caption         =   "Weapons:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   82
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label lblShield 
         Caption         =   "Shield:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   81
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label lblHelm 
         Caption         =   "Helm:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   80
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lblArmor 
         Caption         =   "Armor:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblMagic 
         Caption         =   "Magic Items:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   78
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblNotes 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Character History and Notes"
         Height          =   255
         Left            =   -74880
         TabIndex        =   77
         Top             =   2580
         Width           =   6855
      End
      Begin VB.Label lblClimb 
         Caption         =   "Climb Walls:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   76
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblTraps 
         Caption         =   "Find/Remove Traps:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   75
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblHide 
         Caption         =   "Hide In Shadows:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   74
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblMove 
         Caption         =   "Move Silently:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   73
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblRead 
         Caption         =   "Read Languages:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   72
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblLocks 
         Caption         =   "Open Locks:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   71
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblPockets 
         Caption         =   "Pick Pockets:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   70
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblDetect 
         Caption         =   "Detect Noise:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSkeleton 
         Caption         =   "Skeleton:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   68
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblZombie 
         Caption         =   "Zombie:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   67
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblGhoul 
         Caption         =   "Ghoul:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   66
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblShaodw 
         Caption         =   "Shadow:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   65
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblWight 
         Caption         =   "Wight:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   64
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblGhast 
         Caption         =   "Ghast:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   63
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblWraith 
         Caption         =   "Wraith:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   62
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblMummy 
         Caption         =   "Mummy:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   61
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblSpectre 
         Caption         =   "Spectre:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   60
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblVampire 
         Caption         =   "Vampire:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   59
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblGhost 
         Caption         =   "Ghost"
         Height          =   255
         Left            =   -70680
         TabIndex        =   58
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblLiche 
         Caption         =   "Liche:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   57
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblSpecial 
         Caption         =   "Special:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   56
         Top             =   3600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmNPChar"
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

Public Sub cmdGenerate_Click()
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
    ' classes.  It also fills in default data for the items in
    ' the Inventory tab.
    '============================================================
    Select Case cmbClass.Text
        Case "Fighter"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
            txtArmor.Text = "Chain Mail"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Longsword"
            txtWeapon2.Text = "Dagger"
            txtMagic1.Text = "Ring of Protection +1"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of Magic Missles"
            txtMagic4.Text = "Bracers of Speed"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "3"
            txtNotes.Text = "This character is a Fighter."
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
            txtArmor.Text = "Plate Mail"
            txtShield.Text = "Shield +1"
            txtHelm.Text = "Great Helm"
            txtWeapon1.Text = "Vorpal Longsword"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of Wishes (3)"
            txtMagic4.Text = "Bracers of Speed"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "-3"
            txtNotes.Text = "This character is a Paladin."
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
            txtArmor.Text = "Leather +2"
            txtShield.Text = "Shield"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Longsword +2"
            txtWeapon2.Text = "Longbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of True Seeing"
            txtMagic4.Text = "Bracers of Archery"
            txtMagic5.Text = "Dagger of Return +1"
            txtAC.Text = "3"
            txtNotes.Text = "This character is a Ranger."
        Case "Mage"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            txtHP.Text = RollDice(1, 4)
            txtArmor.Text = "Robe +2"
            txtShield.Text = "None"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Quarterstaff +1"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Wizardry"
            txtMagic3.Text = "Amulet of Spell Storing"
            txtMagic4.Text = "Cloak of Displacement"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "5"
            txtNotes.Text = "This character is a Mage."
        Case "Illusionist"
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtInt.Text < 9 Then txtInt.Text = 9
            txtHP.Text = RollDice(1, 4)
            GetSaving Me
            txtArmor.Text = "Robe +2"
            txtShield.Text = "None"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Quarterstaff +1"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Wizardry"
            txtMagic3.Text = "Amulet of Spell Storing"
            txtMagic4.Text = "Cloak of Displacement"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "5"
            txtNotes.Text = "This character is an Illusionist."
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
            txtArmor.Text = "Chain Mail +1"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Mace +2"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Speed"
            txtMagic4.Text = "Amulet of Life Stealing"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "1"
            txtNotes.Text = "This character is a Cleric."
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
            txtArmor.Text = "Leather +2"
            txtShield.Text = "None"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Quarterstaff +2"
            txtWeapon2.Text = "Longbow +2"
            txtMagic1.Text = "Ring of Protection +3"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of Speaking to Animals"
            txtMagic4.Text = "Wand of Teleportation (15)"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "3"
            txtNotes.Text = "This character is a Druid."
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
            txtArmor.Text = "Leather +1"
            txtShield.Text = "Shield +1"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Short Sword +1"
            txtWeapon2.Text = "Handheld Crossbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Dexterity"
            txtMagic4.Text = "Dagger of Return +2"
            txtMagic5.Text = "Skeleton Key"
            txtAC.Text = "4"
            txtNotes.Text = "This character is a Bard."
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
            txtArmor.Text = "Leather +1"
            txtShield.Text = "Shield +1"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Short Sword +1"
            txtWeapon2.Text = "Handheld Crossbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Dexterity"
            txtMagic4.Text = "Dagger of Return +2"
            txtMagic5.Text = "Skeleton Key"
            txtAC.Text = "4"
            txtNotes.Text = "This character is a Thief."
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
            txtArmor.Text = "Chain Mail"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Longsword"
            txtWeapon2.Text = "Dagger"
            txtMagic1.Text = "Ring of Protection +1"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of Magic Missles"
            txtMagic4.Text = "Bracers of Speed"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "3"
            txtNotes.Text = "This character is a Fighter/Thief."
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
            txtArmor.Text = "Chain Mail +1"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Mace +2"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Speed"
            txtMagic4.Text = "Amulet of Life Stealing"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "1"
            txtNotes.Text = "This character is a Fighter/Cleric."
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
            txtArmor.Text = "Chain Mail"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Longsword"
            txtWeapon2.Text = "Dagger"
            txtMagic1.Text = "Ring of Protection +1"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of Magic Missles"
            txtMagic4.Text = "Bracers of Speed"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "3"
            txtNotes.Text = "This character is a Fighter/Mage."
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
            txtArmor.Text = "Chain Mail"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Longsword"
            txtWeapon2.Text = "Dagger"
            txtMagic1.Text = "Ring of Protection +1"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of Magic Missles"
            txtMagic4.Text = "Bracers of Speed"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "3"
            txtNotes.Text = "This character is a Fighter/Illusionist."
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
            txtArmor.Text = "Leather +2"
            txtShield.Text = "Shield"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Longsword +2"
            txtWeapon2.Text = "Longbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Amulet of True Seeing"
            txtMagic4.Text = "Bracers of Archery"
            txtMagic5.Text = "Dagger of Return +1"
            txtAC.Text = "3"
            txtNotes.Text = "This character is a Cleric/Ranger."
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
            txtArmor.Text = "Chain Mail +1"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Mace +2"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Speed"
            txtMagic4.Text = "Amulet of Life Stealing"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "1"
            txtNotes.Text = "This character is a Cleric/Mage."
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
            txtArmor.Text = "Chain Mail +1"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Mace +2"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Speed"
            txtMagic4.Text = "Amulet of Life Stealing"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "1"
            txtNotes.Text = "This character is a Cleric/Illusionist."
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
            txtArmor.Text = "Chain Mail +1"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Mace +2"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Speed"
            txtMagic4.Text = "Amulet of Life Stealing"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "1"
            txtNotes.Text = "This character is a Cleric/Druid."
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
            txtArmor.Text = "Leather +1"
            txtShield.Text = "Shield +1"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Short Sword +1"
            txtWeapon2.Text = "Handheld Crossbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Dexterity"
            txtMagic4.Text = "Dagger of Return +2"
            txtMagic5.Text = "Skeleton Key"
            txtAC.Text = "4"
            txtNotes.Text = "This character is a Cleric/Thief."
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
            txtArmor.Text = "Leather +1"
            txtShield.Text = "Shield +1"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Short Sword +1"
            txtWeapon2.Text = "Handheld Crossbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Dexterity"
            txtMagic4.Text = "Dagger of Return +2"
            txtMagic5.Text = "Skeleton Key"
            txtAC.Text = "4"
            txtNotes.Text = "This character is a Mage/Thief."
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
            txtArmor.Text = "Leather +1"
            txtShield.Text = "Shield +1"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Short Sword +1"
            txtWeapon2.Text = "Handheld Crossbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Dexterity"
            txtMagic4.Text = "Dagger of Return +2"
            txtMagic5.Text = "Skeleton Key"
            txtAC.Text = "4"
            txtNotes.Text = "This character is a Illusionist/Thief."
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
            txtArmor.Text = "Chain Mail +1"
            txtShield.Text = "Shield"
            txtHelm.Text = "Helmet"
            txtWeapon1.Text = "Mace +2"
            txtWeapon2.Text = "Dagger +1"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Speed"
            txtMagic4.Text = "Amulet of Life Stealing"
            txtMagic5.Text = "Warp Marble"
            txtAC.Text = "1"
            txtNotes.Text = "This character is a Fighter/Mage/Cleric."
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
            txtArmor.Text = "Leather +1"
            txtShield.Text = "Shield +1"
            txtHelm.Text = "None"
            txtWeapon1.Text = "Short Sword +1"
            txtWeapon2.Text = "Handheld Crossbow +2"
            txtMagic1.Text = "Ring of Protection +2"
            txtMagic2.Text = "Ring of Regeneration"
            txtMagic3.Text = "Bracers of Dexterity"
            txtMagic4.Text = "Dagger of Return +2"
            txtMagic5.Text = "Skeleton Key"
            txtAC.Text = "4"
            txtNotes.Text = "This character is a Fighter/Mage/Thief."
        Case Else
            tabPC.TabEnabled(3) = False
            tabPC.TabEnabled(4) = False
            cmbAlignment.Enabled = True
            cmbRace.Enabled = True
            GetStats Me
            If txtStr.Text < 9 Then txtStr.Text = 9
            txtHP.Text = RollDice(1, 10)
            GetSaving Me
            txtNotes.Text = "This character is a " & cmbClass.Text & "."
    End Select
End Sub

Public Sub ClearData()
    '============================================================
    ' This resets the form and clears all data.
    '============================================================
    txtName.Text = ""
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
