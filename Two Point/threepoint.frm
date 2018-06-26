VERSION 5.00
Begin VB.Form frmThreePoint 
   Caption         =   "Three Point"
   ClientHeight    =   7230
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTwoPoint 
      Caption         =   "Two Point"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7980
      TabIndex        =   31
      Top             =   6000
      Width           =   2175
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   5640
      ScaleHeight     =   4845
      ScaleWidth      =   4485
      TabIndex        =   30
      Top             =   960
      Width           =   4515
      Begin VB.Label lblPoint1Info 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPoint2Info 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPoint3Info 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   1
      Left            =   60
      Top             =   1980
   End
   Begin VB.PictureBox picObtuse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2880
      Picture         =   "threepoint.frx":0000
      ScaleHeight     =   1245
      ScaleWidth      =   1185
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picAcute 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2880
      Picture         =   "threepoint.frx":0B13
      ScaleHeight     =   1245
      ScaleWidth      =   1185
      TabIndex        =   28
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2880
      Picture         =   "threepoint.frx":160E
      ScaleHeight     =   1245
      ScaleWidth      =   1185
      TabIndex        =   27
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picIsosceles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   1080
      Picture         =   "threepoint.frx":228B
      ScaleHeight     =   1245
      ScaleWidth      =   1185
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picScalene 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   1080
      Picture         =   "threepoint.frx":2E10
      ScaleHeight     =   1245
      ScaleWidth      =   1185
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picEquilateral 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   1080
      Picture         =   "threepoint.frx":3710
      ScaleHeight     =   1245
      ScaleWidth      =   1185
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7980
      TabIndex        =   20
      Top             =   6540
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   19
      Top             =   6540
      Width           =   2175
   End
   Begin VB.TextBox txtArea 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2700
      MaxLength       =   5
      TabIndex        =   16
      Top             =   4380
      Width           =   1575
   End
   Begin VB.TextBox txtPerimeter 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   960
      TabIndex        =   15
      Top             =   4380
      Width           =   1575
   End
   Begin VB.TextBox txtAngleA 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      MaxLength       =   5
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtAngleC 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtAngleB 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1860
      MaxLength       =   5
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtSideB 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1860
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtSideC 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtSideA 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblEquilateralScaleneIsosceles 
      Alignment       =   2  'Center
      Caption         =   "Equilateral"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   23
      Top             =   5460
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblAcuteRightObtuse 
      Alignment       =   2  'Center
      Caption         =   "Obtuse"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2820
      TabIndex        =   22
      Top             =   5460
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblPerimeter 
      Alignment       =   2  'Center
      Caption         =   "Perimeter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   18
      Top             =   4020
      Width           =   1275
   End
   Begin VB.Label lblArea 
      Alignment       =   2  'Center
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2820
      TabIndex        =   17
      Top             =   4020
      Width           =   1275
   End
   Begin VB.Label lblAngleA 
      Alignment       =   2  'Center
      Caption         =   "Angle A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label lblAngleB 
      Alignment       =   2  'Center
      Caption         =   "Angle B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1980
      TabIndex        =   13
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label lblAngleC 
      Alignment       =   2  'Center
      Caption         =   "Angle C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   12
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label lblIsNotTriangle 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "It is not a triangle."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   540
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblIsTriangle 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "It is a triangle!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   540
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblSideC 
      Alignment       =   2  'Center
      Caption         =   "Side C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblSideB 
      Alignment       =   2  'Center
      Caption         =   "Side B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1980
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblSideA 
      Alignment       =   2  'Center
      Caption         =   "Side A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Three Point"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5580
      TabIndex        =   0
      ToolTipText     =   "Click for more info."
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "frmThreePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Most of this code is basically directly copied from Triangle Checker.
Dim SideA As Single
Dim SideB As Single
Dim SideC As Single
Dim AngleA As Single
Dim AngleB As Single
Dim AngleC As Single
Dim Perimeter As Single
Dim SemiPerimeter As Single
Dim Area As Single
Dim LawOfCosines As Single
Dim IsTriangle As Boolean
Dim IsEquilateral As Boolean
Dim IsScalene As Boolean
Dim IsIsosceles As Boolean
Dim IsRight As Boolean
Dim IsAcute As Boolean
Dim IsObtuse As Boolean
Dim Counter As Integer
Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim y1 As Single
Dim y2 As Single
Dim y3 As Single

Private Sub cmdCalculate_Click()
If Not txtSideA.Text = "" And Not txtSideB.Text = "" And Not txtSideC.Text = "" Then
    Calculate
Else
    MsgBox ("Please finish graphing and plotting points!")
End If
End Sub

Private Sub cmdClear_Click()
Clear
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdTwoPoint_Click()
frmThreePoint.Hide
frmTwoPoint.Show
End Sub

Private Sub Form_Activate()
Clear
ScaleDraw
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

Private Sub lblTitle_Click()
MsgBox ("Made by Jack Margeson for Computer Programming 1. Summer 2018.")
End Sub

Private Sub mnuCalculate_Click()
Calculate
End Sub

Private Sub mnuClear_Click()
Clear
End Sub

Private Sub mnuExit_Click()
End
End Sub

Sub Clear()
frmThreePoint.Visible = True
txtSideA.Enabled = True
txtSideB.Enabled = True
txtSideC.Enabled = True
txtSideA.Text = ""
txtSideB.Text = ""
txtSideC.Text = ""
txtAngleA.Text = ""
txtAngleB.Text = ""
txtAngleC.Text = ""
txtPerimeter.Text = ""
txtArea.Text = ""
lblIsTriangle.Visible = False
lblIsNotTriangle.Visible = False
SideA = 0
SideB = 0
SideC = 0
AngleA = 0
AngleB = 0
AngleC = 0
Perimeter = 0
SemiPerimeter = 0
Area = 0
IsTriangle = False
IsEquilateral = False
IsScalene = False
IsIsosceles = False
IsRight = False
IsAcute = False
IsObtuse = False
picEquilateral.Visible = False
picScalene.Visible = False
picIsosceles.Visible = False
picRight.Visible = False
picAcute.Visible = False
picObtuse.Visible = False
lblEquilateralScaleneIsosceles.Visible = False
lblAcuteRightObtuse.Visible = False
txtSideA.SetFocus

'This clears the new stuff not taken from Triangle Checker.
x1 = 0
x2 = 0
x3 = 0
y1 = 0
y2 = 0
y3 = 0
lblPoint1Info.Visible = False
lblPoint2Info.Visible = False
lblPoint3Info.Visible = False
picGrid.Cls
ScaleDraw
Counter = 0
End Sub

Sub Calculate()
SideA = Val(txtSideA)
SideB = Val(txtSideB)
SideC = Val(txtSideC)
    If SideA + SideB > SideC And SideA + SideC > SideB And SideB + SideC > SideA Then
        IsTriangle = True
        lblIsTriangle.Visible = True

        'Perimeter
        Perimeter = SideA + SideB + SideC
        txtPerimeter.Text = Perimeter

        'Semi Perimeter
        SemiPerimeter = Perimeter / 2

        'Area
        Area = Sqr(SemiPerimeter * (SemiPerimeter - SideA) * (SemiPerimeter - SideB) * (SemiPerimeter - SideC))
        txtArea.Text = Area

        'Angles
        Dim Pi As Single
        Pi = 3.14159265358979

        'Angle A
        LawOfCosines = ((SideB ^ 2) + (SideC ^ 2) - (SideA ^ 2)) / (2 * SideB * SideC)
        AngleA = Atn(-LawOfCosines / Sqr(-LawOfCosines * LawOfCosines + 1)) + 2 * Atn(1)
        AngleA = AngleA * 180 / Pi

        'Angle B
        LawOfCosines = ((SideC ^ 2) + (SideA ^ 2) - (SideB ^ 2)) / (2 * SideC * SideA)
        AngleB = Atn(-LawOfCosines / Sqr(-LawOfCosines * LawOfCosines + 1)) + 2 * Atn(1)
        AngleB = AngleB * 180 / Pi

        'Angle C
        AngleC = 180 - (AngleA + AngleB)

        'Rounding
        AngleA = Math.Round(AngleA)
        AngleB = Math.Round(AngleB)
        AngleC = Math.Round(AngleC)
        txtAngleA.Text = AngleA
        txtAngleB.Text = AngleB
        txtAngleC.Text = AngleC
    
        'Acute, Right, Obtuse
        If AngleA = 90 Or AngleB = 90 Or AngleC = 90 Then
            IsRight = True
            picRight.Visible = True
            lblAcuteRightObtuse.Caption = "Right"
            lblAcuteRightObtuse.Visible = True
        ElseIf AngleA > 90 Or AngleB > 90 Or AngleC = 90 Then
            IsObtuse = True
            picObtuse.Visible = True
            lblAcuteRightObtuse.Caption = "Obtuse"
            lblAcuteRightObtuse.Visible = True
        Else
            IsAcute = True
            picAcute.Visible = True
            lblAcuteRightObtuse.Caption = "Acute"
            lblAcuteRightObtuse.Visible = True
        End If
        
        'Equilateral, Scalene, Isosceles
        If SideA = SideB And SideB = SideC Then
            IsEquilateral = True
            IsScalene = False
            IsIsosceles = False
            picEquilateral.Visible = True
            lblEquilateralScaleneIsosceles.Caption = "Equilateral"
            lblEquilateralScaleneIsosceles.Visible = True
        ElseIf Not SideA = SideB And Not SideB = SideC Then
            IsScalene = True
            IsEquilateral = False
            IsIsosceles = False
            picScalene.Visible = True
            lblEquilateralScaleneIsosceles.Caption = "Scalene"
            lblEquilateralScaleneIsosceles.Visible = True
        Else
            IsIsosceles = True
            IsEquilateral = False
            IsScalene = False
            picIsosceles.Visible = True
            lblEquilateralScaleneIsosceles.Caption = "Isosceles"
            lblEquilateralScaleneIsosceles.Visible = True
        End If
    Else
        IsTriangle = False
        lblIsNotTriangle.Visible = True
    End If

    txtSideA.Enabled = False
    txtSideB.Enabled = False
    txtSideC.Enabled = False
    Counter = 4
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Counter = Counter + 1

If Counter = 1 Then
    x1 = x
    y1 = Y
    picGrid.Circle (x1, y1), 0.25
    
    lblPoint1Info.Visible = True
    lblPoint1Info = "(" + Format(x1, "fixed") + ", " + Format(y1, "fixed") + ")"
    lblPoint1Info.Top = y1 - 0.5
    lblPoint1Info.Left = x1 + 0.5
ElseIf Counter = 2 Then
    x2 = x
    y2 = Y
    picGrid.Line (x1, y1)-(x2, y2), vbRed
    picGrid.Circle (x2, y2), 0.25
    
    lblPoint2Info.Visible = True
    lblPoint2Info = "(" + Format(x2, "fixed") + ", " + Format(y2, "fixed") + ")"
    lblPoint2Info.Top = y2 - 0.5
    lblPoint2Info.Left = x2 + 0.5
    
    txtSideA.Text = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
ElseIf Counter = 3 Then
    x3 = x
    y3 = Y
    picGrid.Line (x2, y2)-(x3, y3), vbGreen
    picGrid.Line (x3, y3)-(x1, y1), vbBlue
    picGrid.Circle (x3, y3), 0.25
    
    lblPoint3Info.Visible = True
    lblPoint3Info = "(" + Format(x3, "fixed") + ", " + Format(y3, "fixed") + ")"
    lblPoint3Info.Top = y3 - 0.5
    lblPoint3Info.Left = x3 + 0.5
    
    txtSideB.Text = Sqr((x3 - x2) ^ 2 + (y3 - y2) ^ 2)
    txtSideC.Text = Sqr((x1 - x3) ^ 2 + (y1 - y3) ^ 2)
    Calculate
End If
End Sub
Private Sub tmrAnimation_Timer()
Dim Blinking As Integer
Blinking = Second(Now) Mod 2
If lblIsTriangle.Visible = True Or lblIsNotTriangle.Visible = True Then
    If Blinking = 0 Then
       lblIsTriangle.BackColor = &HC000&
       lblIsNotTriangle.BackColor = &HFF&
    Else
       lblIsTriangle.BackColor = &H80000016
       lblIsNotTriangle.BackColor = &H80000016
    End If
End If
End Sub

Private Sub txtSideA_Change()
If Not txtSideA.Text = "0" Then
    If IsNumeric(txtSideA) = False And Not txtSideA.Text = "" Then
        txtSideA.Text = ""
        MsgBox ("Please enter valid numbers only.")
        txtSideA.SetFocus
    Else
        SideA = Val(txtSideA)
    End If
Else
    txtSideA.Text = ""
    MsgBox ("You can't have a side length of 0.")
    txtSideA.SetFocus
End If
End Sub

Private Sub txtSideB_Change()
If Not txtSideB.Text = "0" Then
    If IsNumeric(txtSideB) = False And Not txtSideB.Text = "" Then
        txtSideB.Text = ""
        MsgBox ("Please enter valid numbers only.")
        txtSideB.SetFocus
    Else
        SideB = Val(txtSideB)
    End If
Else
    txtSideB.Text = ""
    MsgBox ("You can't have a side length of 0.")
    txtSideB.SetFocus
End If
End Sub

Private Sub txtSideC_Change()
If Not txtSideC.Text = "0" Then
    If IsNumeric(txtSideC) = False And Not txtSideC.Text = "" Then
        txtSideC.Text = ""
        MsgBox ("Please enter valid numbers only.")
        txtSideC.SetFocus
    Else
        SideC = Val(txtSideC)
    End If
Else
    txtSideC.Text = ""
    MsgBox ("You can't have a side length of 0.")
    txtSideC.SetFocus
End If
End Sub

Private Sub txtSideA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSideB.SetFocus
End If
End Sub

Private Sub txtSideB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSideC.SetFocus
End If
End Sub

Private Sub txtSideC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Calculate
End If
End Sub

