Attribute VB_Name = "modUtilities"
Option Explicit
'============================================================
' This module contains 6 function and the Windows API code
' declarations required by the function.  This section
' contains all the API declarations and any data types needed
' by the following functions.
'============================================================
Public Const mlngWindows95 = 0
Public Const mlngWindowsNT = 1
Public Const mlngWindows2000 = 2
Public Declare Function GetVersion Lib "kernel32" () As Long
Public glngWhichWindows32 As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Public Type POINTAPI
    x As Long
    y As Long
    End Type

Public Type SIZE
    cx As Long
    cy As Long
    End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
    End Type

Public Const WS_EX_LAYERED = &H80000
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const AC_SRC_OVER = &H0
Public Const AC_SRC_ALPHA = &H1
Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
Public Const AC_SRC_NO_ALPHA = &H2
Public Const AC_DST_NO_PREMULT_ALPHA = &H10
Public Const AC_DST_NO_ALPHA = &H20
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public lret As Long

Sub Main()
    '============================================================
    ' The first thing it does is check to see if the program has
    ' run.  If so, it shows the main form.  If not, it sets the
    ' DateRan setting and the RunOnce setting in the registry.
    ' Then, it goes through and sets up the default settings for
    ' the program.  After all that, it shows the main form.  This
    ' function runs the registration process.  Basically the
    ' program expires after 30 days and quits running.
    '============================================================
    Dim RunOnce As String
    Dim DateRan As String
    Dim ExpireDate As String
    Dim Date1 As Date
    Dim Date2 As Date
    Dim ExpDateDiff As Integer
    Dim Register As Integer
    Dim RegID As Variant
    
    RunOnce = GetSetting("AD&D Tools", "Info", "RunOnce")
    DateRan = Format(Date, "dd-mmmm-yyyy")
    
    Select Case RunOnce
        Case 0      ' This is run when the program is unregistered.
            ExpireDate = GetSetting("AD&D Tools", "Info", "FirstRan")
            Date1 = Format(ExpireDate, "m/d/yyyy")
            Date2 = Format(DateRan, "m/d/yyyy")
            ExpDateDiff = DateDiff("d", Date1, Date2)
            If ExpDateDiff > 30 Then
                Register = MsgBox("This program may be used for 30 days." & vbCrLf & "Would you like to register to continue use.", vbCritical + vbYesNo, "Please register")
                Select Case Register
                    Case 6
                        frmRegister.Show
                    Case 7
                        End
                End Select
            Else
                Register = MsgBox("This program may be used for 30 days." & vbCrLf & "Would you like to register to continue use.", vbCritical + vbYesNo, "Please register")
                Select Case Register
                    Case 6
                        frmRegister.Show
                    Case 7
                        frmMain.Show
                        SaveSetting "AD&D Tools", "Info", "LastRan", DateRan
                End Select
            End If
        Case 1      ' This is run when the program has been registered.
            If RegTest = True Then
                frmMain.Show
            Else
                frmRegister.Show
            End If
        Case Else   ' This is run when the registry setting is not found
            SaveSetting "AD&D Tools", "Info", "RunOnce", "0"
            SaveSetting "AD&D Tools", "Info", "FirstRan", DateRan
            SaveSetting "AD&D Tools", "Info", "LastRan", DateRan
            SaveSetting "AD&D Tools", "Settings", "IconSize", 1
            MsgBox "This program may be used for 30 days." & vbCrLf & "You will be prompted to register next" & vbCrLf & "time you run the program.", vbCritical + vbOKOnly, "Please register"
            
            frmMain.Show
    End Select
End Sub

Function RegTest() As Boolean
    '============================================================
    ' This checks the number entered into the registry as the
    ' RegId to what the number is supposed to be and returns a
    ' boolean value.
    '============================================================
    Dim RegNum As String
    
    RegNum = GetSetting("AD&D Tools", "Info", "RegID")
    
    If RegNum = "123456789" Then
        RegTest = True
    Else
        RegTest = False
    End If
End Function

Function CheckLayered(ByVal hWnd As Long) As Boolean
On Error Resume Next
    '============================================================
    ' This checks to see if the window is transparent.
    '============================================================
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (lret And WS_EX_LAYERED) = WS_EX_LAYERED Then
        CheckLayered = True
    Else
        CheckLayered = False
    End If
End Function

Function SetLayered(ByVal hWnd As Long, SetAs As Boolean, bAlpha As Byte)
On Error Resume Next
    '============================================================
    ' This sets the window to be transparent or not based on the
    ' user input.
    '============================================================
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If SetAs = True Then
        lret = lret Or WS_EX_LAYERED
    Else
        lret = lret And Not WS_EX_LAYERED
    End If
    SetWindowLong hWnd, GWL_EXSTYLE, lret
    SetLayeredWindowAttributes hWnd, 0, bAlpha, LWA_ALPHA
End Function

Public Sub DoTrans(FormName As Object)
On Error Resume Next
    '============================================================
    ' This function, when called, turns the program transparent.
    ' Note: This only works in Windows 2000.
    '============================================================
    Dim lngVersion As Long
    lngVersion = GetVersion()
    If lngVersion = 143851525 Then
        SetLayered FormName.hWnd, True, 150
    End If
End Sub

Public Sub DoGradient(FormName As Object)
On Error Resume Next
    '============================================================
    ' This function creates a Black to Blue, top down gradient.
    '============================================================
    Dim i As Integer, y As Integer
    FormName.AutoRedraw = True
    FormName.DrawStyle = 6
    FormName.DrawMode = 13
    FormName.DrawWidth = 13
    FormName.ScaleMode = 3
    FormName.ScaleHeight = 256
    For i = 0 To 510
        FormName.Line (0, y)-(FormName.Width, y + 1), RGB(0, 0, i), BF
        y = y + 1
    Next i
End Sub

Public Function DBPath() As String
On Error Resume Next
    '============================================================
    ' This function gets the path to the database.  It first
    ' checks the registry for the path.  If not found there, it
    ' opens an inputbox with the application path and a database
    ' name as the default.  When this is entered, it sets that
    ' path in the the registry.
    '============================================================
    DBPath = GetSetting("AD&D Tools", "Settings", "Path")
    If DBPath = "" Then
        DBPath = InputBox("Please type the full path to your database.", "Database Not Found", App.Path & "\character.mdb")
        SaveSetting "AD&D Tools", "Settings", "Path", DBPath
    End If
End Function

Public Function RollDice(DiceNum As Integer, Sides As Integer) As Integer
    '============================================================
    ' This function takes the number of sides and the number of
    ' dice and returns a random number based on the input.
    '============================================================
    Dim x As Integer, Roll As Integer
    For x = 1 To DiceNum
        Roll = Int(Rnd() * Sides) + 1
        RollDice = RollDice + Roll
    Next x
End Function

Public Function GetStats(FormName As Object)
    '============================================================
    ' This section generates the main statistics based on a 3d6
    ' method of rolling.
    '============================================================
    FormName.txtStr.Text = RollDice(3, 6)
    If FormName.txtStr.Text = 18 Then
        FormName.txtStrAdj.Text = RollDice(1, 100)
    Else
        FormName.txtStrAdj.Text = 0
    End If
    FormName.txtDex.Text = RollDice(3, 6)
    FormName.txtInt.Text = RollDice(3, 6)
    FormName.txtWis.Text = RollDice(3, 6)
    FormName.txtCon.Text = RollDice(3, 6)
    FormName.txtCha.Text = RollDice(3, 6)
End Function

Public Function GetTurning(FormName As Object)
    '============================================================
    ' Fills in the correct values for the Turning Undead section
    ' of the character sheet/generator
    '============================================================
    Dim db As Database, Level, strSearch
    Dim rs As Recordset
    
    Set db = OpenDatabase(DBPath)
    Set rs = db.OpenRecordset("tblInfoTurn", dbOpenDynaset)
    
    Level = FormName.txtLevel.Text
    
    strSearch = "[Level] = " & Level
    
    With rs
        .FindFirst strSearch
        FormName.txtSkeleton.Text = rs.Fields("Skeleton")
        FormName.txtZombie.Text = rs.Fields("Zombie")
        FormName.txtGhoul.Text = rs.Fields("Ghoul")
        FormName.txtShadow.Text = rs.Fields("Shadow")
        FormName.txtWight.Text = rs.Fields("Wight")
        FormName.txtGhast.Text = rs.Fields("Ghast")
        FormName.txtWraith.Text = rs.Fields("Wraith")
        FormName.txtMummy.Text = rs.Fields("Mummy")
        FormName.txtSpectre.Text = rs.Fields("Spectre")
        FormName.txtVampire.Text = rs.Fields("Vampire")
        FormName.txtGhost.Text = rs.Fields("Ghost")
        FormName.txtLiche.Text = rs.Fields("Liche")
        FormName.txtSpecial.Text = rs.Fields("Special")
    End With
End Function

Public Function GetThief(FormName As Object)
    '============================================================
    ' This fills in the values for the rogue skills on the
    ' character sheet/generator
    '============================================================
    FormName.txtClimb.Text = "60%"
    FormName.txtNoise.Text = "15%"
    FormName.txtTraps.Text = "5%"
    FormName.txtHide.Text = "5%"
    FormName.txtSilent.Text = "10%"
    FormName.txtLocks.Text = "10%"
    FormName.txtPockets.Text = "15%"
    FormName.txtRead.Text = "0%"
End Function

Public Function GetSaving(FormName As Object)
    '============================================================
    ' This gets the correct values for the saving throws based on
    ' level and class.
    '============================================================
    Dim db As Database, Level, strSearch
    Dim rs As Recordset, CharClass
    
    Set db = OpenDatabase(DBPath)
    Set rs = db.OpenRecordset("tblInfoSaving", dbOpenDynaset)
    
    Level = FormName.txtLevel.Text
    CharClass = FormName.cmbClass.Text
    
    strSearch = "[Level] = " & Level & "AND [Class] Like '" & CharClass & "'"
    
    With rs
        .FindFirst strSearch
        FormName.txtBreath.Text = rs.Fields("Breath")
        FormName.txtRod.Text = rs.Fields("Rod")
        FormName.txtSpell.Text = rs.Fields("Spell")
        FormName.txtPetr.Text = rs.Fields("Petr")
        FormName.txtPara.Text = rs.Fields("Para")
    End With
End Function

