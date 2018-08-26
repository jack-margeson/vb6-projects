VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hangman - Menu"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   30
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNoose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   -720
      Picture         =   "menu.frx":0000
      ScaleHeight     =   3735
      ScaleWidth      =   4815
      TabIndex        =   5
      Top             =   900
      Width           =   4815
      Begin VB.Line lnBody 
         BorderWidth     =   5
         X1              =   1980
         X2              =   1980
         Y1              =   1380
         Y2              =   2100
      End
      Begin VB.Line lnHead 
         BorderWidth     =   50
         X1              =   1980
         X2              =   1980
         Y1              =   960
         Y2              =   1020
      End
      Begin VB.Line lnLeftLeg 
         BorderWidth     =   5
         X1              =   1920
         X2              =   1800
         Y1              =   2160
         Y2              =   2700
      End
      Begin VB.Line lnRightLeg 
         BorderWidth     =   5
         X1              =   2040
         X2              =   2220
         Y1              =   2160
         Y2              =   2700
      End
      Begin VB.Line lnRightArm 
         BorderWidth     =   5
         X1              =   2040
         X2              =   2220
         Y1              =   1560
         Y2              =   2040
      End
      Begin VB.Line lnLeftArm 
         BorderWidth     =   5
         X1              =   1920
         X2              =   1680
         Y1              =   1620
         Y2              =   2100
      End
   End
   Begin VB.CommandButton cmdPlay2 
      Caption         =   "2 Players"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   3195
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   6300
      Width           =   3195
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   3195
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "1 Player"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4980
      Width           =   3195
   End
   Begin VB.Label lblTitle 
      Caption         =   "Hangman"
      BeginProperty Font 
         Name            =   "Prestige Elite Std"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Path As String
Public player As Single

Private Sub cmdCredits_Click()
MsgBox ("Made by Jack Margeson. Computer Programming II, 2018.")
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdPlay_Click()
Path = App.Path & "\dictionary.txt"
MsgBox ("Click OK to load the dictionary file from " & Path & ". This may take up to a minute.")
Me.Hide
player = 0
frmHangman.player = player
frmHangman.Show
End Sub

Private Sub cmdPlay2_Click()
Path = App.Path & "\dictionary.txt"
MsgBox ("Click OK to load the dictionary file from " & Path & ". This may take up to a minute.")
Me.Hide
player = 1
frmHangman.player = player
frmHangman.Show
End Sub
