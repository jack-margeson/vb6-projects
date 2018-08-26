VERSION 5.00
Begin VB.Form TicTacToe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9135
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "tictactoe.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBlinking 
      Interval        =   1
      Left            =   4260
      Top             =   480
   End
   Begin VB.Timer tmrAI 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4260
      Top             =   3180
   End
   Begin VB.TextBox txtPlayer2Wins 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2220
      TabIndex        =   14
      Top             =   2460
      Width           =   1995
   End
   Begin VB.TextBox txtPlayer1Wins 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2460
      Width           =   1995
   End
   Begin VB.TextBox txtTurn 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox txtPlayer2Name 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2220
      MaxLength       =   8
      TabIndex        =   9
      Top             =   1500
      Width           =   1995
   End
   Begin VB.TextBox txtPlayer1Name 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1500
      Width           =   1995
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3780
      Width           =   1995
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3780
      Width           =   1995
   End
   Begin VB.CommandButton cmdAI 
      Caption         =   "vs. AI"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1995
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play 1v1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   5940
      TabIndex        =   23
      Top             =   1860
      Width           =   1635
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   7680
      TabIndex        =   22
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   4560
      TabIndex        =   21
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   5940
      TabIndex        =   20
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5940
      TabIndex        =   19
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   7680
      TabIndex        =   18
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   4560
      TabIndex        =   17
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   16
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      TabIndex        =   15
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblPlayer2Wins 
      Caption         =   "Player 2 Wins:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2220
      TabIndex        =   12
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Label lblPlayer1Wins 
      Caption         =   "Player 1 Wins:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Line ln7 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   8880
      Y1              =   720
      Y2              =   4200
   End
   Begin VB.Line ln8 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4620
      X2              =   8880
      Y1              =   4140
      Y2              =   780
   End
   Begin VB.Line ln6 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   8280
      X2              =   8280
      Y1              =   720
      Y2              =   4260
   End
   Begin VB.Line ln5 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   6780
      X2              =   6780
      Y1              =   720
      Y2              =   4260
   End
   Begin VB.Line ln4 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   5280
      X2              =   5280
      Y1              =   720
      Y2              =   4260
   End
   Begin VB.Line ln3 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   8880
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Line ln2 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   8880
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line ln1 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4680
      X2              =   8880
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line lnGrid4 
      BorderColor     =   &H8000000B&
      BorderWidth     =   4
      X1              =   4560
      X2              =   9000
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line lnGrid3 
      BorderColor     =   &H8000000B&
      BorderWidth     =   4
      X1              =   4560
      X2              =   9000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line lnGrid2 
      BorderColor     =   &H8000000B&
      BorderWidth     =   4
      X1              =   7620
      X2              =   7620
      Y1              =   720
      Y2              =   4260
   End
   Begin VB.Line lnGrid1 
      BorderColor     =   &H8000000B&
      BorderWidth     =   4
      X1              =   5880
      X2              =   5880
      Y1              =   720
      Y2              =   4260
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Tic-Tac-Toe"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4620
      MousePointer    =   4  'Icon
      TabIndex        =   3
      Top             =   0
      Width           =   4395
   End
   Begin VB.Label lblPlayer2 
      Caption         =   "Player 2:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2220
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblPlayer1 
      Caption         =   "Player 1:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label lblMove 
      Caption         =   "Who's turn?"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2415
   End
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GameActive As Boolean
Dim Player1Turn As Boolean
Dim lbl1Filled As Boolean
Dim lbl2Filled As Boolean
Dim lbl3Filled As Boolean
Dim lbl4Filled As Boolean
Dim lbl5Filled As Boolean
Dim lbl6Filled As Boolean
Dim lbl7Filled As Boolean
Dim lbl8Filled As Boolean
Dim lbl9Filled As Boolean
Dim Player1WinCount As Single
Dim Player2WinCount As Single
Dim LastWinner As Single
Dim PlayingAI As Boolean

Private Sub cmdClear_Click()
FullClear
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdPlay_Click()
    If GameActive = False Then
        GameActive = True
        PlayingAI = False
        tmrAI.Enabled = False
        GameClear
        EnableLabel
            If LastWinner = 0 Then
                Player1Turn = True
            ElseIf LastWinner = 1 Then
                Player1Turn = False
            ElseIf LastWinner = 2 Then
                Player1Turn = True
            Else
            MsgBox ("Something has gone fatally wrong. This shouldn't happen...")
            End If
            If txtPlayer1Name.Text = "" Then
                txtPlayer1Name.Text = "X"
            End If
            If txtPlayer2Name.Text = "" Then
                txtPlayer2Name.Text = "O"
            End If
            If txtPlayer1Name.Text = txtPlayer2Name.Text Then
                MsgBox ("How 'bout you guys choose different names from each other?")
                GameActive = False
                FullClear
            End If
            If txtPlayer1Name.Text = "Jack" Then
                txtPlayer1Wins.Text = "9999999999"
                MsgBox ("You think you can beat me?!")
            ElseIf txtPlayer2Name.Text = "Jack" Then
                txtPlayer2Wins.Text = "9999999999"
                MsgBox ("You think you can beat me?!")
            End If
        If Not txtPlayer1Name.Text = "" And Not txtPlayer2Name.Text = "" Then
            If Player1Turn = True Then
                MsgBox ("The game has started! Player 1 - you start first.")
                txtTurn.Text = txtPlayer1Name.Text
            ElseIf Player1Turn = False Then
                MsgBox ("The game has started! Player 2 - you start first.")
                txtTurn.Text = txtPlayer2Name.Text
            End If
        End If
    Else
        MsgBox ("Hey, there's already a game in progress!")
    End If
End Sub

Private Sub cmdAI_Click()
    If GameActive = False Then
        GameActive = True
        PlayingAI = True
        tmrAI.Enabled = True
        GameClear
        EnableLabel
        txtPlayer2Name.Text = "AI"
            If LastWinner = 0 Then
                Player1Turn = True
            ElseIf LastWinner = 1 Then
                Player1Turn = False
            ElseIf LastWinner = 2 Then
                Player1Turn = True
            Else
            MsgBox ("Something has gone fatally wrong. This shouldn't happen...")
            End If
            If txtPlayer1Name.Text = "" Then
                txtPlayer1Name.Text = "X"
            End If
            If txtPlayer1Name.Text = txtPlayer2Name.Text Then
                MsgBox ("How 'bout you guys choose different names from each other?")
                GameActive = False
                FullClear
            End If
            If txtPlayer1Name.Text = "Jack" Then
                txtPlayer1Wins.Text = "9999999999"
                MsgBox ("You think you can beat me?!")
            ElseIf txtPlayer2Name.Text = "Jack" Then
                txtPlayer2Wins.Text = "9999999999"
                MsgBox ("You think you can beat me?!")
            End If
        If Not txtPlayer1Name.Text = "" And Not txtPlayer2Name.Text = "" Then
            If Player1Turn = True Then
                MsgBox ("The game has started! Player 1 - you start first.")
                txtTurn.Text = txtPlayer1Name.Text
            ElseIf Player1Turn = False Then
                MsgBox ("The game has started! Player 2 - you start first.")
                txtTurn.Text = txtPlayer2Name.Text
            End If
        End If
    Else
        MsgBox ("Hey, there's already a game in progress!")
    End If
End Sub
Private Sub Form_Load()
TicTacToe.Visible = True
txtPlayer1Name.SetFocus
lbl1.ForeColor = vbWhite
lbl2.ForeColor = vbWhite
lbl3.ForeColor = vbWhite
lbl4.ForeColor = vbWhite
lbl5.ForeColor = vbWhite
lbl6.ForeColor = vbWhite
lbl7.ForeColor = vbWhite
lbl8.ForeColor = vbWhite
lbl9.ForeColor = vbWhite
End Sub

Private Sub lbl1_Click()
    If GameActive = True And lbl1Filled = False Then
        If Player1Turn = True And lbl1Filled = False Then
            lbl1.Caption = "X"
            Player1Turn = False
            lbl1Filled = True
            txtTurn.Text = txtPlayer2Name.Text
            CheckWin
        Else
            lbl1.Caption = "O"
            Player1Turn = True
            lbl1Filled = True
            txtTurn.Text = txtPlayer1Name.Text
            CheckWin
        End If
        Else
            If GameActive = False Then
                MsgBox ("The game has not yet started! Please click play.")
            End If
            If lbl1Filled = True Then
                MsgBox ("This spot is already filled!")
            End If
    End If
End Sub

Private Sub lbl2_Click()
    If GameActive = True And lbl2Filled = False Then
        If Player1Turn = True And lbl2Filled = False Then
            lbl2.Caption = "X"
            Player1Turn = False
            lbl2Filled = True
            txtTurn.Text = txtPlayer2Name.Text
            CheckWin
        Else
            lbl2.Caption = "O"
            Player1Turn = True
            lbl2Filled = True
            txtTurn.Text = txtPlayer1Name.Text
            CheckWin
        End If
        Else
            If GameActive = False Then
                MsgBox ("The game has not yet started! Please click play.")
            End If
            If lbl2Filled = True Then
                MsgBox ("This spot is already filled!")
            End If
    End If
End Sub

Private Sub lbl3_Click()
    If GameActive = True And lbl3Filled = False Then
        If Player1Turn = True And lbl3Filled = False Then
            lbl3.Caption = "X"
            Player1Turn = False
            lbl3Filled = True
            txtTurn.Text = txtPlayer2Name.Text
            CheckWin
        Else
            lbl3.Caption = "O"
            Player1Turn = True
            lbl3Filled = True
            txtTurn.Text = txtPlayer1Name.Text
            CheckWin
        End If
        Else
            If GameActive = False Then
                MsgBox ("The game has not yet started! Please click play.")
            End If
            If lbl3Filled = True Then
                MsgBox ("This spot is already filled!")
            End If
    End If
End Sub

Private Sub lbl4_Click()
If GameActive = True And lbl4Filled = False Then
    If Player1Turn = True And lbl4Filled = False Then
        lbl4.Caption = "X"
        Player1Turn = False
        lbl4Filled = True
        txtTurn.Text = txtPlayer2Name.Text
        CheckWin
    Else
        lbl4.Caption = "O"
        Player1Turn = True
        lbl4Filled = True
        txtTurn.Text = txtPlayer1Name.Text
        CheckWin
    End If
    Else
        If GameActive = False Then
            MsgBox ("The game has not yet started! Please click play.")
        End If
        If lbl4Filled = True Then
            MsgBox ("This spot is already filled!")
        End If
End If
End Sub

Private Sub lbl5_Click()
If GameActive = True And lbl5Filled = False Then
    If Player1Turn = True And lbl5Filled = False Then
        lbl5.Caption = "X"
        Player1Turn = False
        lbl5Filled = True
        txtTurn.Text = txtPlayer2Name.Text
        CheckWin
    Else
        lbl5.Caption = "O"
        Player1Turn = True
        lbl5Filled = True
        txtTurn.Text = txtPlayer1Name.Text
        CheckWin
    End If
    Else
        If GameActive = False Then
            MsgBox ("The game has not yet started! Please click play.")
        End If
        If lbl5Filled = True Then
            MsgBox ("This spot is already filled!")
        End If
End If
End Sub

Private Sub lbl6_Click()
If GameActive = True And lbl6Filled = False Then
    If Player1Turn = True And lbl6Filled = False Then
        lbl6.Caption = "X"
        Player1Turn = False
        lbl6Filled = True
        txtTurn.Text = txtPlayer2Name.Text
        CheckWin
    Else
        lbl6.Caption = "O"
        Player1Turn = True
        lbl6Filled = True
        txtTurn.Text = txtPlayer1Name.Text
        CheckWin
    End If
    Else
        If GameActive = False Then
            MsgBox ("The game has not yet started! Please click play.")
        End If
        If lbl6Filled = True Then
            MsgBox ("This spot is already filled!")
        End If
End If
End Sub

Private Sub lbl7_Click()
If GameActive = True And lbl7Filled = False Then
    If Player1Turn = True And lbl7Filled = False Then
        lbl7.Caption = "X"
        Player1Turn = False
        lbl7Filled = True
        txtTurn.Text = txtPlayer2Name.Text
        CheckWin
    Else
        lbl7.Caption = "O"
        Player1Turn = True
        lbl7Filled = True
        txtTurn.Text = txtPlayer1Name.Text
        CheckWin
    End If
    Else
        If GameActive = False Then
            MsgBox ("The game has not yet started! Please click play.")
        End If
        If lbl7Filled = True Then
            MsgBox ("This spot is already filled!")
        End If
End If
End Sub

Private Sub lbl8_Click()
If GameActive = True And lbl8Filled = False Then
    If Player1Turn = True And lbl8Filled = False Then
        lbl8.Caption = "X"
        Player1Turn = False
        lbl8Filled = True
        txtTurn.Text = txtPlayer2Name.Text
        CheckWin
    Else
        lbl8.Caption = "O"
        Player1Turn = True
        lbl8Filled = True
        txtTurn.Text = txtPlayer1Name.Text
        CheckWin
    End If
    Else
        If GameActive = False Then
            MsgBox ("The game has not yet started! Please click play.")
        End If
        If lbl8Filled = True Then
            MsgBox ("This spot is already filled!")
        End If
End If
End Sub

Private Sub lbl9_Click()
If GameActive = True And lbl9Filled = False Then
    If Player1Turn = True And lbl9Filled = False Then
        lbl9.Caption = "X"
        Player1Turn = False
        lbl9Filled = True
        txtTurn.Text = txtPlayer2Name.Text
        CheckWin
    Else
        lbl9.Caption = "O"
        Player1Turn = True
        lbl9Filled = True
        txtTurn.Text = txtPlayer1Name.Text
        CheckWin
    End If
    Else
        If GameActive = False Then
            MsgBox ("The game has not yet started! Please click play.")
        End If
        If lbl9Filled = True Then
            MsgBox ("This spot is already filled!")
        End If
End If
End Sub

Sub CheckWin()
If lbl1.Caption = "X" And lbl2.Caption = "X" And lbl3.Caption = "X" Then
    Player1Win
    ln1.Visible = True
ElseIf lbl1.Caption = "O" And lbl2.Caption = "O" And lbl3.Caption = "O" Then
    Player2Win
    ln1.Visible = True
ElseIf lbl4.Caption = "X" And lbl5.Caption = "X" And lbl6.Caption = "X" Then
    Player1Win
    ln2.Visible = True
ElseIf lbl4.Caption = "O" And lbl5.Caption = "O" And lbl6.Caption = "O" Then
    Player2Win
    ln2.Visible = True
ElseIf lbl7.Caption = "X" And lbl8.Caption = "X" And lbl9.Caption = "X" Then
    Player1Win
    ln3.Visible = True
ElseIf lbl7.Caption = "O" And lbl8.Caption = "O" And lbl9.Caption = "O" Then
    Player2Win
    ln3.Visible = True
ElseIf lbl1.Caption = "X" And lbl4.Caption = "X" And lbl7.Caption = "X" Then
    Player1Win
    ln4.Visible = True
ElseIf lbl1.Caption = "O" And lbl4.Caption = "O" And lbl7.Caption = "O" Then
    Player2Win
    ln4.Visible = True
ElseIf lbl2.Caption = "X" And lbl5.Caption = "X" And lbl8.Caption = "X" Then
    Player1Win
    ln5.Visible = True
ElseIf lbl2.Caption = "O" And lbl5.Caption = "O" And lbl8.Caption = "O" Then
    Player2Win
    ln5.Visible = True
ElseIf lbl3.Caption = "X" And lbl6.Caption = "X" And lbl9.Caption = "X" Then
    Player1Win
    ln6.Visible = True
ElseIf lbl3.Caption = "O" And lbl6.Caption = "O" And lbl9.Caption = "O" Then
    Player2Win
    ln6.Visible = True
ElseIf lbl1.Caption = "X" And lbl5.Caption = "X" And lbl9.Caption = "X" Then
    Player1Win
    ln7.Visible = True
ElseIf lbl1.Caption = "O" And lbl5.Caption = "O" And lbl9.Caption = "O" Then
    Player2Win
    ln7.Visible = True
ElseIf lbl3.Caption = "X" And lbl5.Caption = "X" And lbl7.Caption = "X" Then
    Player1Win
    ln8.Visible = True
ElseIf lbl3.Caption = "O" And lbl5.Caption = "O" And lbl7.Caption = "O" Then
    Player2Win
    ln8.Visible = True
ElseIf lbl1Filled = True And lbl2Filled = True And lbl3Filled = True And lbl4Filled = True And lbl5Filled = True And lbl6Filled = True And lbl7Filled = True And lbl8Filled = True And lbl9Filled = True Then
    GameTie
End If
End Sub

Sub EnableLabel()
lbl1.Enabled = True
lbl2.Enabled = True
lbl3.Enabled = True
lbl4.Enabled = True
lbl5.Enabled = True
lbl6.Enabled = True
lbl7.Enabled = True
lbl8.Enabled = True
lbl9.Enabled = True
End Sub

Sub DisableLabel()
lbl1.Enabled = False
lbl2.Enabled = False
lbl3.Enabled = False
lbl4.Enabled = False
lbl5.Enabled = False
lbl6.Enabled = False
lbl7.Enabled = False
lbl8.Enabled = False
lbl9.Enabled = False
End Sub

Sub Player1Win()
GameActive = False
DisableLabel
txtTurn.Text = ""
MsgBox ("Player 1 wins!")
Player1WinCount = Player1WinCount + 1
txtPlayer1Wins.Text = Player1WinCount
txtTurn.ForeColor = &H80000016
txtTurn.BackColor = &H80000016
LastWinner = 1
If PlayingAI = False Then
    cmdPlay.Caption = "Rematch"
    lblMove.Caption = "Winner!"
ElseIf PlayingAI = True Then
    cmdAI.Caption = "Rematch AI"
    lblMove.Caption = "Winner!"
End If
txtTurn.Text = txtPlayer1Name.Text
End Sub

Sub Player2Win()
GameActive = False
DisableLabel
txtTurn.Text = ""
MsgBox ("Player 2 wins!")
Player2WinCount = Player2WinCount + 1
txtPlayer2Wins.Text = Player2WinCount
txtTurn.ForeColor = &H80000016
txtTurn.BackColor = &H80000016
LastWinner = 2
If PlayingAI = False Then
    cmdPlay.Caption = "Rematch"
    lblMove.Caption = "Winner!"
ElseIf PlayingAI = True Then
    cmdAI.Caption = "Rematch AI"
    lblMove.Caption = "Winner!"
End If
txtTurn.Text = txtPlayer2Name.Text
End Sub

Sub GameTie()
GameActive = False
DisableLabel
If PlayingAI = False Then
    cmdPlay.Caption = "Rematch"
    lblMove.Caption = "Tie Game!"
ElseIf PlayingAI = True Then
    cmdAI.Caption = "Rematch AI"
    lblMove.Caption = "Tie Game!"
End If
txtTurn.Text = ""
MsgBox ("Game is a tie!")
End Sub

Sub GameClear()
ln1.Visible = False
ln2.Visible = False
ln3.Visible = False
ln4.Visible = False
ln5.Visible = False
ln6.Visible = False
ln7.Visible = False
ln8.Visible = False

lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""

lbl1Filled = False
lbl2Filled = False
lbl3Filled = False
lbl4Filled = False
lbl5Filled = False
lbl6Filled = False
lbl7Filled = False
lbl8Filled = False
lbl9Filled = False

lblMove.Caption = "Who's turn?"
cmdPlay.Caption = "Play 1v1"
cmdAI.Caption = "vs. AI"
txtTurn.Text = ""
End Sub

Sub FullClear()
GameClear

txtTurn.Text = ""
txtPlayer1Name.Text = ""
txtPlayer2Name.Text = ""
txtPlayer1Wins.Text = ""
txtPlayer2Wins.Text = ""
Player1WinCount = 0
Player2WinCount = 0
LastWinner = 0
cmdPlay.Caption = "Play"

PlayingAI = False
tmrAI.Enabled = False

EnableLabel
GameActive = False
txtPlayer1Name.SetFocus
End Sub

Private Sub tmrAI_Timer()
Dim RNG As Integer
If Player1Turn = False And GameActive = True And PlayingAI = True Then
    RNG = Int((9 - 1 + 1) * Rnd + 1)
    If RNG = 1 Then
        If lbl1Filled = False Then
            lbl1_Click
        End If
    ElseIf RNG = 2 Then
        If lbl2Filled = False Then
            lbl2_Click
        End If
    ElseIf RNG = 3 Then
        If lbl3Filled = False Then
            lbl3_Click
        End If
    ElseIf RNG = 4 Then
        If lbl4Filled = False Then
            lbl4_Click
        End If
    ElseIf RNG = 5 Then
        If lbl5Filled = False Then
            lbl5_Click
        End If
    ElseIf RNG = 6 Then
        If lbl6Filled = False Then
            lbl6_Click
        End If
    ElseIf RNG = 7 Then
        If lbl7Filled = False Then
            lbl7_Click
        End If
    ElseIf RNG = 8 Then
        If lbl8Filled = False Then
            lbl8_Click
        End If
    ElseIf RNG = 9 Then
        If lbl9Filled = False Then
            lbl9_Click
        End If
    End If
End If
End Sub

Private Sub tmrBlinking_Timer()
Dim Blinking As Integer
Blinking = Second(Now) Mod 2
If GameActive = True Then
    If Blinking = 0 Then
       txtTurn.BackColor = &H80000010
       txtTurn.ForeColor = &H80000016
    Else
       txtTurn.BackColor = &H80000016
       txtTurn.ForeColor = &H80000010
    End If
End If
End Sub

Sub txtPlayer1Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPlayer2Name.SetFocus
    If txtPlayer1Name.Text = "" Then
        txtPlayer1Name.Text = "X"
    End If
End If
End Sub

Sub txtPlayer2Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdPlay_Click
End If
End Sub
Private Sub lblTitle_Click()
MsgBox ("Tic-Tac-Toe by Jack Margeson. Made for Computer Programming 1, Summer 2018.")
End Sub


