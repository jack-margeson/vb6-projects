VERSION 5.00
Begin VB.Form frmParabola 
   Caption         =   "Parabola"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   120
      ScaleHeight     =   4845
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "frmParabola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p As Single
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim a As Single

Private Sub Form_Load()
p = 0

picGrid.Visible = True
picGrid.Scale (-10, 10)-(10, -10)
picGrid.Line (-10, 0)-(10, 0)
picGrid.Line (0, -10)-(0, 10)

Dim i As Single
For i = -10 To 10 Step 1
    picGrid.Line (i, 0.4)-(i, -0.4)
    picGrid.Line (0.4, i)-(-0.4, i)
Next i
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If p = 0 Then
    x1 = X
    y1 = Y
    picGrid.Circle (x1, y1), 0.25
    p = p + 1
ElseIf p = 1 Then
    x2 = X
    y2 = Y
    picGrid.Circle (x2, y2), 0.25
    p = p - 2
    
    'defines a
    Dim k As Single
    Dim h As Single
    k = y1
    h = x1
    a = (y2 - k) / ((x2 - h) ^ 2)
    MsgBox (a)
    
    'draws lines
    For X = -10 To 10 Step 0.005
        Y = a * ((X - x1) ^ 2) + (y1)
        picGrid.Circle (X, Y), 0.1
    Next X
End If
End Sub
