VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JRSs Mancala (beta 1)"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextTotalInCups 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   39
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox TextPlayer2sTurns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   33
      Text            =   "0"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox TextPlayer1sTurns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   32
      Text            =   "0"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame FramePlayer2Turns 
      Caption         =   "Player 2s"
      Height          =   615
      Left            =   1920
      TabIndex        =   36
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame FramePlayer1Turns 
      Caption         =   "Player 1s"
      Height          =   615
      Left            =   360
      TabIndex        =   35
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox TextCurrentPlayer 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6120
      TabIndex        =   30
      Text            =   "Player 1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "12"
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "11"
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "10"
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TextPlayer1Mancala 
      Enabled         =   0   'False
      Height          =   735
      Left            =   1200
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox TextPlayer2Mancala 
      Enabled         =   0   'False
      Height          =   735
      Left            =   5520
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox TextPlayer2Cup6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox TextPlayer2Cup5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox TextPlayer2Cup4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox TextPlayer2Cup3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox TextPlayer2Cup2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox TextPlayer2Cup1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox TextPlayer1Cup1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TextPlayer1Cup2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TextPlayer1Cup3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TextPlayer1Cup4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TextPlayer1Cup5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox TextPlayer1Cup6 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Frame FrameTotalTurns 
      Caption         =   "Total Turns per Player"
      Height          =   975
      Left            =   240
      TabIndex        =   34
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label LabelTotalInCups 
      Caption         =   "Total Stones in Cups"
      Height          =   375
      Left            =   4680
      TabIndex        =   40
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label LabelPlayer2sClockwise 
      Caption         =   "--->"
      Height          =   255
      Left            =   4680
      TabIndex        =   38
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label LabelPlayer1sClockwise 
      Caption         =   "<---"
      Height          =   255
      Left            =   2640
      TabIndex        =   37
      Top             =   960
      Width           =   375
   End
   Begin VB.Label LabelCurrentPlayer 
      Caption         =   "Current Player"
      Height          =   495
      Left            =   5400
      TabIndex        =   31
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LabelPlayer2Cups 
      Caption         =   "Player 2s Cups"
      Height          =   255
      Left            =   3120
      TabIndex        =   29
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label LabelPlayer1Cups 
      Caption         =   "Player 1s Cups"
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label LabelPlayer2Mancala 
      Caption         =   "Player 2s Mancala"
      Height          =   495
      Left            =   6240
      TabIndex        =   27
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label LabelPlayer1Mancala 
      Caption         =   "Player 1s Mancala"
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   1800
      Width           =   735
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuFileNew1 
         Caption         =   "New - &1 Player"
      End
      Begin VB.Menu MenuFileNew2 
         Caption         =   "New - &2 Player"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If TextCurrentPlayer.Text = "Player 2" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer1Cup6.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "1" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "2" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "3" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "4" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "5" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "6" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "7" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "8" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  If TextPlayer2Cup6.Text > 0 And TextPlayer1Cup1.Text = 1 Then
   TextPlayer1Mancala = Val(TextPlayer1Mancala) + TextPlayer2Cup6.Text + TextPlayer1Cup1.Text
   TextPlayer2Cup6 = "0"
   TextPlayer1Cup1 = "0"
  End If
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "9" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  If TextPlayer2Cup5.Text > 0 And TextPlayer1Cup2.Text = 1 Then
   TextPlayer1Mancala = Val(TextPlayer1Mancala) + TextPlayer2Cup5.Text + TextPlayer1Cup2.Text
   TextPlayer2Cup5 = "0"
   TextPlayer1Cup2 = "0"
  TextCurrentPlayer.Text = "Player 2"
  End If
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup6.Text = "10" Then
  TextPlayer1Cup6.Text = "0"
  TextPlayer1Mancala = TextPlayer1Mancala + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  If TextPlayer2Cup4.Text > 0 And TextPlayer1Cup3.Text = 1 Then
   TextPlayer1Mancala = Val(TextPlayer1Mancala) + TextPlayer2Cup4.Text + TextPlayer1Cup3.Text
   TextPlayer2Cup4 = "0"
   TextPlayer1Cup3 = "0"
  End If
  TextCurrentPlayer.Text = "Player 2"
 End If
 
 If TextPlayer1Cup1.Text = "0" And TextPlayer1Cup2.Text = "0" And TextPlayer1Cup3.Text = "0" And TextPlayer1Cup4.Text = "0" And TextPlayer1Cup5.Text = "0" And TextPlayer1Cup6.Text = "0" Then
  TextPlayer2Mancala.Text = Val(TextPlayer2Mancala.Text) + TextPlayer2Cup1.Text + TextPlayer2Cup2.Text + TextPlayer2Cup3.Text + TextPlayer2Cup4.Text + TextPlayer2Cup5.Text + TextPlayer2Cup6.Text
 End If
 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command10_Click()
 If TextCurrentPlayer.Text = "Player 1" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer2Cup4.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "1" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "2" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "3" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "4" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "5" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "6" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "7" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "8" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "9" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup4.Text = "10" Then
  TextPlayer2Cup4.Text = "0"
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  If TextPlayer1Cup6.Text > 0 And TextPlayer2Cup1.Text = 1 Then
   TextPlayer2Mancala = Val(TextPlayer2Mancala) + TextPlayer1Cup6.Text + TextPlayer2Cup1.Text
   TextPlayer1Cup6 = "0"
   TextPlayer2Cup1 = "0"
  End If
  TextCurrentPlayer.Text = "Player 1"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command11_Click()
 If TextCurrentPlayer.Text = "Player 1" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer2Cup5.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "1" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  If TextPlayer1Cup1.Text > 0 And TextPlayer2Cup6.Text = 1 Then
   TextPlayer2Mancala = Val(TextPlayer2Mancala) + TextPlayer1Cup1.Text + TextPlayer2Cup6.Text
   TextPlayer1Cup1 = "0"
   TextPlayer2Cup6 = "0"
  End If
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "2" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "3" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "4" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "5" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "6" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "7" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "8" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "9" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup5.Text = "10" Then
  TextPlayer2Cup5.Text = "0"
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 
 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command12_Click()
 If TextCurrentPlayer.Text = "Player 1" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer2Cup6.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "1" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "2" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "3" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "4" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "4" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "5" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "6" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "7" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "8" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  If TextPlayer1Cup6.Text > 0 And TextPlayer2Cup1.Text = 1 Then
   TextPlayer2Mancala = Val(TextPlayer2Mancala) + TextPlayer1Cup6.Text + TextPlayer2Cup1.Text
   TextPlayer1Cup6 = "0"
   TextPlayer2Cup1 = "0"
  End If
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "9" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  If TextPlayer1Cup5.Text > 0 And TextPlayer2Cup2.Text = 1 Then
   TextPlayer2Mancala = Val(TextPlayer2Mancala) + TextPlayer1Cup5.Text + TextPlayer2Cup2.Text
   TextPlayer1Cup5 = "0"
   TextPlayer2Cup2 = "0"
  End If
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup6.Text = "10" Then
  TextPlayer2Cup6.Text = "0"
  TextPlayer2Mancala = TextPlayer2Mancala + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  If TextPlayer1Cup4.Text > 0 And TextPlayer2Cup3.Text = 1 Then
   TextPlayer2Mancala = Val(TextPlayer2Mancala) + TextPlayer1Cup4.Text + TextPlayer2Cup3.Text
   TextPlayer1Cup4 = "0"
   TextPlayer2Cup3 = "0"
  End If
  TextCurrentPlayer.Text = "Player 1"
 End If
 
 If TextPlayer2Cup1.Text = "0" And TextPlayer2Cup2.Text = "0" And TextPlayer2Cup3.Text = "0" And TextPlayer2Cup4.Text = "0" And TextPlayer2Cup5.Text = "0" And TextPlayer2Cup6.Text = "0" Then
  TextPlayer1Mancala.Text = Val(TextPlayer1Mancala.Text) + TextPlayer1Cup1.Text + TextPlayer1Cup2.Text + TextPlayer1Cup3.Text + TextPlayer1Cup4.Text + TextPlayer1Cup5.Text + TextPlayer1Cup6.Text
 End If
 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command2_Click()
 If TextCurrentPlayer.Text = "Player 2" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer1Cup5.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "1" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  If TextPlayer2Cup1.Text > 0 And TextPlayer1Cup6.Text = 1 Then
   TextPlayer1Mancala = Val(TextPlayer1Mancala) + TextPlayer2Cup1.Text + TextPlayer1Cup6.Text
   TextPlayer2Cup1 = "0"
   TextPlayer1Cup6 = "0"
  End If
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "2" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "3" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "4" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "5" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "6" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "7" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "8" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "9" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup5.Text = "10" Then
  TextPlayer1Cup5.Text = "0"
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command3_Click()
 If TextCurrentPlayer.Text = "Player 2" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer1Cup4.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "1" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "2" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "3" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "4" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "5" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "6" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "7" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "8" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "9" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup4.Text = "10" Then
  TextPlayer1Cup4.Text = "0"
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command4_Click()
 If TextCurrentPlayer.Text = "Player 2" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer1Cup3.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "1" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "2" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "3" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "4" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "5" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "6" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "7" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "8" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "9" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup3.Text = "10" Then
  TextPlayer1Cup3.Text = "0"
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command5_Click()
 If TextCurrentPlayer.Text = "Player 2" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer1Cup2.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "1" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "2" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "3" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "4" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "5" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "6" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "7" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "8" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "9" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup2.Text = "10" Then
  TextPlayer1Cup2.Text = "0"
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command6_Click()
 If TextCurrentPlayer.Text = "Player 2" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer1Cup1.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "1" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "2" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "3" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "4" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "5" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "6" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "7" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "8" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "9" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 1" And TextPlayer1Cup1.Text = "10" Then
  TextPlayer1Cup1.Text = "0"
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer1Mancala.Text = TextPlayer1Mancala.Text + 1
  TextPlayer2Cup1.Text = TextPlayer2Cup1.Text + 1
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer1sTurns.Text = TextPlayer1sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command7_Click()
 If TextCurrentPlayer.Text = "Player 1" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer2Cup1.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "1" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "2" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "3" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "4" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "5" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "6" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "7" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "8" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "9" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup1.Text = "10" Then
  TextPlayer2Cup1.Text = "0"
  TextPlayer2Cup2.Text = TextPlayer2Cup2.Text + 1
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command8_Click()
 If TextCurrentPlayer.Text = "Player 1" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer2Cup2.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "1" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "2" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "3" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "4" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "5" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "6" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "7" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "8" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "9" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup2.Text = "10" Then
  TextPlayer2Cup2.Text = "0"
  TextPlayer2Cup3.Text = TextPlayer2Cup3.Text + 1
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Command9_Click()
 If TextCurrentPlayer.Text = "Player 1" Then
  MsgBox "Illegal Move - It's not your turn", vbOKOnly, "Illegal Move"
 End If
 If TextPlayer2Cup3.Text = "0" Then
  MsgBox "Illegal Move - Choose a cup with stones", vbOKOnly, "Illegal Move"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "1" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "2" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "3" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "4" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 2"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "5" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "6" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "7" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "8" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "9" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If
 If TextCurrentPlayer.Text = "Player 2" And TextPlayer2Cup3.Text = "10" Then
  TextPlayer2Cup3.Text = "0"
  TextPlayer2Cup4.Text = TextPlayer2Cup4.Text + 1
  TextPlayer2Cup5.Text = TextPlayer2Cup5.Text + 1
  TextPlayer2Cup6.Text = TextPlayer2Cup6.Text + 1
  TextPlayer2Mancala.Text = TextPlayer2Mancala.Text + 1
  TextPlayer1Cup1.Text = TextPlayer1Cup1.Text + 1
  TextPlayer1Cup2.Text = TextPlayer1Cup2.Text + 1
  TextPlayer1Cup3.Text = TextPlayer1Cup3.Text + 1
  TextPlayer1Cup4.Text = TextPlayer1Cup4.Text + 1
  TextPlayer1Cup5.Text = TextPlayer1Cup5.Text + 1
  TextPlayer1Cup6.Text = TextPlayer1Cup6.Text + 1
  TextPlayer2sTurns.Text = TextPlayer2sTurns.Text + 1
  TextCurrentPlayer.Text = "Player 1"
 End If

 TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Load()
 TextTotalInCups = 48 - (TextPlayer1Mancala.Text + TextPlayer2Mancala.Text)
End Sub

Private Sub MenuFileNew1_Click()
TextCurrentPlayer.Text = "Player 1"
TextPlayer1Cup1.Text = "4"
TextPlayer1Cup2.Text = "4"
TextPlayer1Cup3.Text = "4"
TextPlayer1Cup4.Text = "4"
TextPlayer1Cup5.Text = "4"
TextPlayer1Cup6.Text = "4"
TextPlayer2Cup1.Text = "4"
TextPlayer2Cup2.Text = "4"
TextPlayer2Cup3.Text = "4"
TextPlayer2Cup4.Text = "4"
TextPlayer2Cup5.Text = "4"
TextPlayer2Cup6.Text = "4"
TextPlayer1Mancala.Text = "0"
TextPlayer2Mancala.Text = "0"
TextPlayer1sTurns = "0"
TextPlayer2sTurns = "0"
TextTotalInCups = 48 - (Val(TextPlayer1Mancala.Text) + Val(TextPlayer2Mancala.Text))
End Sub

Private Sub MenuHelpAbout_Click()
 MsgBox "JRSs Mancala vBeta1 - Copyright 2003 http://jamesrskemp.net/", vbOKOnly, "About JRSs Mancala"
End Sub
