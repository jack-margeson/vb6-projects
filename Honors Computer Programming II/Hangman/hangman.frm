VERSION 5.00
Begin VB.Form frmHangman 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hangman"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Noto Sans Hebrew"
      Size            =   36
      Charset         =   1
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdPlayAgain 
      Caption         =   "Play Again 2 Players!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   1380
      TabIndex        =   36
      Top             =   4620
      Width           =   1095
   End
   Begin VB.PictureBox picNoose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   60
      Picture         =   "hangman.frx":0000
      ScaleHeight     =   3735
      ScaleWidth      =   4815
      TabIndex        =   35
      Top             =   60
      Width           =   4815
      Begin VB.Line lnLeftArm 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   1920
         X2              =   1680
         Y1              =   1620
         Y2              =   2100
      End
      Begin VB.Line lnRightArm 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   2040
         X2              =   2220
         Y1              =   1560
         Y2              =   2040
      End
      Begin VB.Line lnRightLeg 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   2040
         X2              =   2220
         Y1              =   2160
         Y2              =   2700
      End
      Begin VB.Line lnLeftLeg 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   1920
         X2              =   1800
         Y1              =   2160
         Y2              =   2700
      End
      Begin VB.Line lnHead 
         BorderWidth     =   50
         Visible         =   0   'False
         X1              =   1980
         X2              =   1980
         Y1              =   960
         Y2              =   1020
      End
      Begin VB.Line lnBody 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   1980
         X2              =   1980
         Y1              =   1380
         Y2              =   2100
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   5220
      Width           =   2235
   End
   Begin VB.CommandButton cmdPlayAgain 
      Caption         =   "Play Again 1 Player!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   32
      Top             =   4620
      Width           =   1095
   End
   Begin VB.ListBox lstdashes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   9840
      TabIndex        =   3
      Top             =   2160
      Width           =   1995
   End
   Begin VB.ListBox lstletters 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   7740
      TabIndex        =   2
      Top             =   2160
      Width           =   1995
   End
   Begin VB.ListBox lstwords 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   5640
      TabIndex        =   0
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label lblWins 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   34
      Top             =   5160
      Width           =   1995
   End
   Begin VB.Label lblLives 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   31
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   25
      Left            =   3540
      TabIndex        =   30
      Top             =   7980
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   24
      Left            =   2880
      TabIndex        =   29
      Top             =   7980
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   23
      Left            =   2220
      TabIndex        =   28
      Top             =   7980
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   22
      Left            =   1560
      TabIndex        =   27
      Top             =   7980
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   21
      Left            =   900
      TabIndex        =   26
      Top             =   7980
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   20
      Left            =   4200
      TabIndex        =   25
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   19
      Left            =   3540
      TabIndex        =   24
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   18
      Left            =   2880
      TabIndex        =   23
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   17
      Left            =   2220
      TabIndex        =   22
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   16
      Left            =   1560
      TabIndex        =   21
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   15
      Left            =   900
      TabIndex        =   20
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   14
      Left            =   240
      TabIndex        =   19
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   13
      Left            =   4200
      TabIndex        =   18
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   12
      Left            =   3540
      TabIndex        =   17
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   11
      Left            =   2880
      TabIndex        =   16
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   10
      Left            =   2220
      TabIndex        =   15
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "j"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   1560
      TabIndex        =   14
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   900
      TabIndex        =   13
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   4200
      TabIndex        =   11
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   3540
      TabIndex        =   10
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   2880
      TabIndex        =   9
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   2220
      TabIndex        =   8
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   1560
      TabIndex        =   7
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   900
      TabIndex        =   6
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lblletters 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label lbldashes 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   60
      TabIndex        =   4
      Top             =   3780
      Width           =   4785
   End
   Begin VB.Label lblrword 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   1680
      Width           =   1995
   End
End
Attribute VB_Name = "frmHangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public player As Single

Dim Path As String
Dim strline As String
Dim rword As String
Dim wsf As String
Dim lc As String
Dim check As Single
Dim debugger As Single

Dim flag As Single
Dim lives As Single
Dim wins As Single

Dim arrwords(60000) As String
Dim arrletters(22) As String
Dim arrdashes(22) As String

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdPlayAgain_Click(Index As Integer)
Dim i As Single
lives = 6
lblLives.Caption = "Lives: " & lives
cmdPlayAgain(0).Visible = False
cmdPlayAgain(1).Visible = False
player = Index
EnableUI

lnHead.Visible = False
lnBody.Visible = False
lnLeftArm.Visible = False
lnRightArm.Visible = False
lnLeftLeg.Visible = False
lnRightLeg.Visible = False

Erase arrletters
Erase arrdashes
lstletters.Clear
lstdashes.Clear
wsf = ""
lbldashes.Caption = ""

If player = 0 Then
    Randomize
    rword = arrwords(Int((Rnd * 58112) + 1))
    lblrword.Caption = rword
ElseIf player = 1 Then
    Dim answer As String
    answer = InputBox("Enter word:", "Choose your word, player 1!", "skyline")
    answer = Trim(answer)
    rword = answer
    lblrword.Caption = rword
End If

For i = 1 To Len(rword)
    arrletters(i) = Mid(rword, i, 1)
    lstletters.AddItem arrletters(i)
Next i
For i = 1 To Len(rword)
    arrdashes(i) = "-"
    lstdashes.AddItem arrdashes(i)
    wsf = wsf + "-"
    lbldashes.Caption = wsf
Next i
End Sub

Private Sub Form_Activate()
Dim i As Single
Path = App.Path & "\dictionary.txt"
lives = 6
wins = 0
lblLives.Caption = "Lives: " & lives
lblWins = "Wins: " & wins
cmdPlayAgain(0).Visible = False
cmdPlayAgain(1).Visible = False
EnableUI

Open Path For Input As #1

Do While Not EOF(1)
    Line Input #1, strline
    strline = Trim(strline)
    i = i + 1
    arrwords(i) = strline
    lstwords.AddItem arrwords(i)
Loop

Close #1

If player = 0 Then
    Randomize
    rword = arrwords(Int((Rnd * 58112) + 1))
    lblrword.Caption = rword
ElseIf player = 1 Then
    Dim answer As String
    answer = InputBox("Enter word:", "Choose your word, player 1!", "skyline")
    answer = Trim(answer)
    rword = answer
    lblrword.Caption = rword
End If

For i = 1 To Len(rword)
    arrletters(i) = Mid(rword, i, 1)
    lstletters.AddItem arrletters(i)
Next i
For i = 1 To Len(rword)
    arrdashes(i) = "-"
    lstdashes.AddItem arrdashes(i)
    wsf = wsf + "-"
    lbldashes.Caption = wsf
Next i
End Sub

Private Sub lblletters_Click(Index As Integer)
lc = Chr(Index + 97)
wsf = ""
flag = 0
Dim i As Single
For i = 1 To Len(rword)
    If lc = arrletters(i) Then
        arrdashes(i) = lc
        flag = 1
    End If
    wsf = wsf + arrdashes(i)
Next i
lbldashes.Caption = wsf

If flag = 0 Then
    lives = lives - 1
    lblLives.Caption = "Lives: " & lives
    If lives = 5 Then
        lnHead.Visible = True
    ElseIf lives = 4 Then
        lnBody.Visible = True
    ElseIf lives = 3 Then
        lnLeftArm.Visible = True
    ElseIf lives = 2 Then
        lnRightArm.Visible = True
    ElseIf lives = 1 Then
        lnLeftLeg.Visible = True
    ElseIf lives = 0 Then
        lnRightLeg.Visible = True
        GameOver
    End If
End If

If Not InStr(wsf, "-") > 0 Then
    GameWin
End If

lblletters(Index).Enabled = False
lblletters(Index).BackColor = &H80000000
End Sub

Private Sub GameOver()
MsgBox ("Game over... The word was " & rword & ".")
DisableUI
End Sub

Private Sub GameWin()
MsgBox ("You won! The word was " & rword & ".")
wins = wins + 1
lblWins = "Wins: " & wins
DisableUI
End Sub

Private Sub EnableUI()
Dim i As Single
For i = 0 To 25
    lblletters(i).Enabled = True
    lblletters(i).BackColor = &HC0FFFF
Next i
End Sub

Private Sub DisableUI()
Dim i As Single
For i = 0 To 25
    lblletters(i).Enabled = False
    lblletters(i).BackColor = &H80000000
    cmdPlayAgain(0).Visible = True
    cmdPlayAgain(1).Visible = True
Next i
End Sub

Private Sub lblWins_Click()
debugger = debugger + 1
If debugger >= 8 Then
    frmHangman.Width = 12135
End If
End Sub
