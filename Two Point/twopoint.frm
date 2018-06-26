VERSION 5.00
Begin VB.Form frmTwoPoint 
   Caption         =   "Two Point"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdThreePoint 
      Caption         =   "Three Point"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9300
      TabIndex        =   23
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtPosition 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5940
      TabIndex        =   22
      Top             =   1200
      Width           =   2835
   End
   Begin VB.TextBox txtY4 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   9540
      TabIndex        =   15
      Text            =   "txtY4"
      Top             =   2820
      Width           =   1035
   End
   Begin VB.TextBox txtX4 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   8340
      TabIndex        =   14
      Text            =   "txtX4"
      Top             =   2820
      Width           =   1035
   End
   Begin VB.TextBox txtY3 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6960
      TabIndex        =   13
      Text            =   "txtY3"
      Top             =   2820
      Width           =   1035
   End
   Begin VB.TextBox txtX3 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5760
      TabIndex        =   12
      Text            =   "txtX3"
      Top             =   2820
      Width           =   1035
   End
   Begin VB.TextBox txtY2 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   9540
      TabIndex        =   11
      Text            =   "txtY2"
      Top             =   2100
      Width           =   1035
   End
   Begin VB.TextBox txtX2 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   8340
      TabIndex        =   10
      Text            =   "txtX2"
      Top             =   2100
      Width           =   1035
   End
   Begin VB.TextBox txtY1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6960
      TabIndex        =   9
      Text            =   "txtY1"
      Top             =   2100
      Width           =   1035
   End
   Begin VB.TextBox txtX1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5760
      TabIndex        =   8
      Text            =   "txtX1"
      Top             =   2100
      Width           =   1035
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7920
      TabIndex        =   4
      Top             =   4320
      Width           =   2835
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4740
      TabIndex        =   3
      Top             =   4320
      Width           =   2835
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7920
      TabIndex        =   2
      Top             =   3600
      Width           =   1395
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4740
      TabIndex        =   1
      Top             =   3600
      Width           =   2835
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   60
      ScaleHeight     =   4845
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      Begin VB.Label lblPoint4Info 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3360
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblPoint3Info 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2220
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblPoint2Info 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1140
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPoint1Info 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   60
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Two Point"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   17
      Top             =   60
      Width           =   6315
   End
   Begin VB.Label lblBrackets 
      Caption         =   $"twopoint.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   5640
      TabIndex        =   16
      Top             =   2040
      Width           =   5115
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Caption         =   "Position:"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4860
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblF2 
      Alignment       =   2  'Center
      Caption         =   "f2(x)"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      TabIndex        =   6
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label lblF1 
      Alignment       =   2  'Center
      Caption         =   "f1(x)"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      TabIndex        =   5
      Top             =   2040
      Width           =   795
   End
End
Attribute VB_Name = "frmTwoPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public x1 As Single
Public y1 As Single
Public x2 As Single
Public y2 As Single
Public x3 As Single
Public y3 As Single
Public x4 As Single
Public y4 As Single
Dim Counter As Integer
Dim GraphingComplete As Boolean

Private Sub cmdClear_Click()
FullClear
End Sub

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdFind_Click()
If Not txtX1.Text = "" And Not txtX2.Text = "" And Not txtX3.Text = "" And Not txtX4.Text = "" And Not txtY1.Text = "" And Not txtY2.Text = "" And Not txtY3.Text = "" And Not txtY4.Text = "" Then
    x1 = Val(txtX1)
    x2 = Val(txtX2)
    x3 = Val(txtX3)
    x4 = Val(txtX4)
    y1 = Val(txtY1)
    y2 = Val(txtY2)
    y3 = Val(txtY3)
    y4 = Val(txtY4)
    
    'Drawing label info boxes.
    lblPoint1Info.Visible = True
    lblPoint1Info = "(" + Format(x1, "fixed") + ", " + Format(y1, "fixed") + ")"
    lblPoint1Info.Top = y1 - 0.5
    lblPoint1Info.Left = x1 + 0.5
    
    lblPoint2Info.Visible = True
    lblPoint2Info = "(" + Format(x2, "fixed") + ", " + Format(y2, "fixed") + ")"
    lblPoint2Info.Top = y2 - 0.5
    lblPoint2Info.Left = x2 + 0.5

    lblPoint3Info.Visible = True
    lblPoint3Info = "(" + Format(x3, "fixed") + ", " + Format(y3, "fixed") + ")"
    lblPoint3Info.Top = y3 - 0.5
    lblPoint3Info.Left = x3 + 0.5
    
    lblPoint4Info.Visible = True
    lblPoint4Info = "(" + Format(x4, "fixed") + ", " + Format(y4, "fixed") + ")"
    lblPoint4Info.Top = y4 - 0.5
    lblPoint4Info.Left = x4 + 0.5

    'Drawing circles and lines.
    picGrid.Circle (x1, y1), 0.25
    picGrid.Circle (x2, y2), 0.25
    picGrid.Line (x1, y1)-(x2, y2), vbGreen
    
    picGrid.Circle (x3, y3), 0.25
    picGrid.Circle (x4, y4), 0.25
    picGrid.Line (x3, y3)-(x4, y4), vbRed
    
    Counter = 5
    GraphingComplete = True
Else
    MsgBox ("Please fill out all coordinate information!")
End If
End Sub

Private Sub cmdResults_Click()
If GraphingComplete = True Then
    frmTwoPoint.Hide
    frmResults.Show
    
    frmResults.x1 = x1
    frmResults.x2 = x2
    frmResults.x3 = x3
    frmResults.x4 = x4
    frmResults.y1 = y1
    frmResults.y2 = y2
    frmResults.y3 = y3
    frmResults.y4 = y4
Else
    MsgBox ("Please finish graphing first!")
End If
End Sub

Private Sub cmdThreePoint_Click()
frmTwoPoint.Hide
frmThreePoint.Show
End Sub

Private Sub Form_Activate()
FullClear
ScaleDraw
frmTwoPoint.BackColor = &H8000000F
End Sub

Sub ScaleDraw()
Dim i As Integer

picGrid.Scale (-10, 10)-(10, -10)
picGrid.Line (-10, 0)-(10, 0)
picGrid.Line (0, -10)-(0, 10)

For i = -10 To 10 Step 1
    picGrid.Line (i, 0.4)-(i, -0.4)
    picGrid.Line (0.4, i)-(-0.4, i)
Next i
End Sub

Sub FullClear()
txtX1.Text = ""
txtY1.Text = ""
txtX2.Text = ""
txtY2.Text = ""
txtX3.Text = ""
txtY3.Text = ""
txtX4.Text = ""
txtY4.Text = ""
txtPosition = ""

lblPoint1Info.Visible = False
lblPoint2Info.Visible = False
lblPoint3Info.Visible = False
lblPoint4Info.Visible = False

x1 = 0
x2 = 0
x3 = 0
x4 = 0
y1 = 0
y2 = 0
y3 = 0
y4 = 0

picGrid.Cls
ScaleDraw
frmTwoPoint.BackColor = &H8000000F
GraphingComplete = False

Counter = 0
End Sub

Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim MousePosition As String
MousePosition = "(" + Format(x, "fixed") + ", " + Format(Y, "fixed") + ")"
txtPosition.Text = MousePosition
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Counter = Counter + 1

If Counter = 1 Then
    x1 = x
    y1 = Y
    picGrid.Circle (x1, y1), 0.25
    txtX1 = x1
    txtY1 = y1
    
    lblPoint1Info.Visible = True
    lblPoint1Info = "(" + Format(x1, "fixed") + ", " + Format(y1, "fixed") + ")"
    lblPoint1Info.Top = y1 - 0.5
    lblPoint1Info.Left = x1 + 0.5
    
    frmTwoPoint.BackColor = vbGreen
ElseIf Counter = 2 Then
    x2 = x
    y2 = Y
    picGrid.Line (x1, y1)-(x2, y2), vbGreen
    picGrid.Circle (x2, y2), 0.25
    txtX2 = x2
    txtY2 = y2
    
    lblPoint2Info.Visible = True
    lblPoint2Info = "(" + Format(x2, "fixed") + ", " + Format(y2, "fixed") + ")"
    lblPoint2Info.Top = y2 - 0.5
    lblPoint2Info.Left = x2 + 0.5
    
    frmTwoPoint.BackColor = vbRed
ElseIf Counter = 3 Then
    x3 = x
    y3 = Y
    picGrid.Circle (x3, y3), 0.25
    txtX3 = x3
    txtY3 = y3
    
    lblPoint3Info.Visible = True
    lblPoint3Info = "(" + Format(x3, "fixed") + ", " + Format(y3, "fixed") + ")"
    lblPoint3Info.Top = y3 - 0.5
    lblPoint3Info.Left = x3 + 0.5
ElseIf Counter = 4 Then
    x4 = x
    y4 = Y
    picGrid.Line (x3, y3)-(x4, y4), vbRed
    picGrid.Circle (x4, y4), 0.25
    txtX4 = x4
    txtY4 = y4
    
    lblPoint4Info.Visible = True
    lblPoint4Info = "(" + Format(x4, "fixed") + ", " + Format(y4, "fixed") + ")"
    lblPoint4Info.Top = y4 - 0.5
    lblPoint4Info.Left = x4 + 0.5
    
    frmTwoPoint.BackColor = &H8000000F
    GraphingComplete = True
End If
End Sub

Private Sub txtX1_Change()
frmTwoPoint.BackColor = vbGreen
End Sub

Private Sub txtX2_Change()
frmTwoPoint.BackColor = vbGreen
End Sub

Private Sub txtX3_Change()
frmTwoPoint.BackColor = vbRed
End Sub

Private Sub txtX4_Change()
frmTwoPoint.BackColor = vbRed
End Sub

Private Sub txtY1_Change()
frmTwoPoint.BackColor = vbGreen
End Sub

Private Sub txtY2_Change()
frmTwoPoint.BackColor = vbGreen
End Sub

Private Sub txtY3_Change()
frmTwoPoint.BackColor = vbRed
End Sub

Private Sub txtY4_Change()
frmTwoPoint.BackColor = vbRed
End Sub
