VERSION 5.00
Begin VB.Form frmDice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dice Roll"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmDice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtADDBonus 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtADDDice 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdADDRole 
      Caption         =   "Roll Dice"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.ListBox lstADDResults 
      BackColor       =   &H00C0FFFF&
      Height          =   1620
      ItemData        =   "frmDice.frx":0882
      Left            =   2760
      List            =   "frmDice.frx":0884
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdADDReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtADDTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lstADDDice 
      Height          =   840
      ItemData        =   "frmDice.frx":0886
      Left            =   120
      List            =   "frmDice.frx":08AC
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblRoll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblBonus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bonus / Penalty"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Results"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblDiceToRoll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number of Dice to Roll"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "frmDice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================================
' This form performs the dice rolling.  This section is for
' any global variables used in the form.
'============================================================

Private Sub Form_Load()
    '============================================================
    ' This function runs everytime the form is run.  It gradiates
    ' the background and clears out all textboxes/resets
    ' listboxes
    '============================================================
    DoGradient Me
    ClearAll
End Sub

Private Sub cmdADDRole_Click()
    '============================================================
    ' Clears the last results  Calls the function to roll dice
    ' for AD&D
    '============================================================
    Dim Sides, intLoopIndex As Integer, Bonus
    For intLoopIndex = 0 To lstADDDice.ListCount - 1
        If lstADDDice.Selected(intLoopIndex) Then
            Sides = lstADDDice.ItemData(intLoopIndex)
        End If
    Next intLoopIndex
    If txtADDBonus.Text >= 0 Then
        Bonus = "+" & txtADDBonus.Text
    ElseIf txtADDBonus.Text < 0 Then
        Bonus = txtADDBonus.Text
    End If
    lblRoll.Caption = "You are rolling " & txtADDDice.Text & "d" & Sides & " " & Bonus
    lstADDResults.Clear
    RollADDDice
End Sub

Private Sub cmdADDReset_Click()
    '============================================================
    ' Clears out all Textboxes/Resets Listboxes
    '============================================================
    ClearAll
End Sub

Private Sub cmdExit_Click()
    '============================================================
    ' This function clear the data and closes the window
    '============================================================
    ClearAll
    Unload Me
End Sub

Function RollADDDice()
    '============================================================
    ' This tests the user inputed values and verifies that they
    ' are both integer and greater/equal to 1.  If true, calls
    ' the Dice Rolling function, if not, calls the wrror message
    ' function
    '============================================================
    
    If txtADDDice.Text >= "1" And IsNumeric(txtADDDice.Text) Then
        If IsNumeric(txtADDBonus.Text) Then
            ADDRoll
        Else
            MessageBox 1
        End If
    Else
        MessageBox 0
    End If
    
End Function

Function ADDRoll()
    '============================================================
    ' This function actually performs the roll and builds the
    ' listbox of results, keeping a running total as the dice
    ' are generated.
    '============================================================
    Dim Total, Rolls, Roll, Sides, Result, Bonus, Index As Integer
    Dim intLoopIndex As Integer
    
    Randomize  ' Initializes the random number generator
    
    'This loop gets the selected item from the list and sets it
    'as the number of sides on the dice.
    For intLoopIndex = 0 To lstADDDice.ListCount - 1
        If lstADDDice.Selected(intLoopIndex) Then
            Sides = lstADDDice.ItemData(intLoopIndex)
        End If
    Next intLoopIndex
    
    Rolls = Int(txtADDDice.Text)    'Gets the number of dice to roll from the form.
    Bonus = Int(txtADDBonus.Text)   'Gets the number for the bonus/penalty
    
    'The for/next loop runs for the number of dice called for.
    For Roll = 1 To Rolls               'Rolls the dice a certain number of times
        Result = Int(Rnd * Sides) + 1   'Rolls the dice with the correct sides.
        Result = Result + Bonus         'Adds/Subtracts the Bonus/Penalty
        If Result < 0 Then Result = 1   'Checks for Negative numbers.
        lstADDResults.AddItem (Result)  'Adds the result to the list
        Total = Total + Result          'Sums the total of the dice.
    Next Roll
    
    txtADDTotal.Text = Total            'displays the total sum of the dice.
    
End Function

Function MessageBox(Answer As Integer)
    '============================================================
    ' This is the function that displays the error message box
    ' when invalid input is entered.
    '============================================================
    Dim MsgText As String
    Dim MsgType As String
    Dim MsgTitle As String
    MsgType = vbCritical And vbOKOnly
    MsgTitle = "Invalid Input!"
    Select Case Answer
        Case 0
            MsgText = "Please enter an Integer that is Greater than 0."
        Case 1
            MsgText = "Please enter a Positive or Negative Integer. (i.e. 1 or -1)"
        Case Else
            MsgText = "Please enter the correct data type."
    End Select
        
    MsgBox MsgText, MsgType, MsgTitle
End Function

Function ClearAll()
    '============================================================
    ' This function clears user inputed data and sets defaults
    '============================================================
    txtADDDice.Text = "1"           'Resets the number of dice to roll to 1
    txtADDBonus.Text = "0"          'Resets the Bonus/Penalty to 0
    txtADDTotal.Text = ""           'Clears the total
    lstADDResults.Clear             'Clears the Results list
    lstADDDice.Selected(0) = True   'Sets the first item in list as selected.
    lblRoll.Caption = ""            'Resets the Dice Roll text (1d6 +2)
End Function
