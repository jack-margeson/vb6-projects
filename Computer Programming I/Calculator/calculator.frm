VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00514744&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox gifCalculator 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   8820
      Picture         =   "calculator.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2370
      TabIndex        =   21
      Top             =   3840
      Width           =   2400
   End
   Begin VB.CommandButton cmdLog 
      BackColor       =   &H00514744&
      Caption         =   "Log1(2)"
      Height          =   675
      Left            =   10140
      MaskColor       =   &H00514744&
      TabIndex        =   20
      ToolTipText     =   "Log Base Number 1 of Number 2"
      Top             =   2820
      Width           =   735
   End
   Begin VB.CommandButton cmdLogE 
      BackColor       =   &H00514744&
      Caption         =   "LogE(1)"
      Height          =   675
      Left            =   7800
      MaskColor       =   &H00514744&
      TabIndex        =   19
      ToolTipText     =   "Log Base E of Number 1"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdSquared 
      BackColor       =   &H00514744&
      Caption         =   "^ 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7800
      MaskColor       =   &H00514744&
      TabIndex        =   18
      ToolTipText     =   "Number 1 squared"
      Top             =   2820
      Width           =   735
   End
   Begin VB.CommandButton cmdSquareRoot 
      BackColor       =   &H00514744&
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8940
      MaskColor       =   &H00514744&
      TabIndex        =   17
      ToolTipText     =   "Square Root"
      Top             =   2820
      Width           =   735
   End
   Begin VB.CommandButton cmdTan 
      BackColor       =   &H00514744&
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6660
      MaskColor       =   &H00514744&
      TabIndex        =   16
      ToolTipText     =   "Tangent"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdCos 
      BackColor       =   &H00514744&
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7800
      MaskColor       =   &H00514744&
      TabIndex        =   15
      ToolTipText     =   "Cosine"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdSin 
      BackColor       =   &H00514744&
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6660
      MaskColor       =   &H00514744&
      TabIndex        =   14
      ToolTipText     =   "Sine"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdPower 
      BackColor       =   &H00514744&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6660
      MaskColor       =   &H00514744&
      TabIndex        =   13
      ToolTipText     =   "Number 1 to the power of Number 2"
      Top             =   2820
      Width           =   735
   End
   Begin VB.CommandButton cmdDivide 
      BackColor       =   &H00514744&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10140
      MaskColor       =   &H00514744&
      TabIndex        =   12
      ToolTipText     =   "Division"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdMultiply 
      BackColor       =   &H00514744&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8940
      MaskColor       =   &H00514744&
      TabIndex        =   11
      ToolTipText     =   "Multiplication"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdSubtract 
      BackColor       =   &H00514744&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7800
      MaskColor       =   &H00514744&
      TabIndex        =   10
      ToolTipText     =   "Subtraction"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00514744&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6660
      MaskColor       =   &H00514744&
      TabIndex        =   9
      ToolTipText     =   "Addition"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3180
      TabIndex        =   8
      ToolTipText     =   "Exits the program."
      Top             =   5460
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Clears the input and answer fields."
      Top             =   5460
      Width           =   2655
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox txtNum2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3300
      TabIndex        =   5
      Top             =   1440
      Width           =   2115
   End
   Begin VB.TextBox txtNum1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   540
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label 
      BackColor       =   &H00B5AA99&
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   2
      Left            =   540
      TabIndex        =   3
      Top             =   2820
      Width           =   2115
   End
   Begin VB.Label Label 
      BackColor       =   &H00B5AA99&
      Caption         =   "Number 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   555
      Index           =   1
      Left            =   3300
      TabIndex        =   2
      Top             =   660
      Width           =   2115
   End
   Begin VB.Label Label 
      BackColor       =   &H00B5AA99&
      Caption         =   "Number 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   555
      Index           =   0
      Left            =   540
      TabIndex        =   1
      Top             =   660
      Width           =   2115
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00B5AA99&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   675
      Left            =   6900
      TabIndex        =   0
      ToolTipText     =   "Click me for info!"
      Top             =   660
      Width           =   3855
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Single
Dim Check1 As Boolean
Dim Check2 As Boolean

If Not txtNum1.Text = "" Then
    Num1 = txtNum1.Text
    Else
    MsgBox ("Please enter a valid number for #1.")
    Check1 = True
End If
If Not txtNum2.Text = "" Then
    Num2 = txtNum2.Text
    Else
    MsgBox ("Please enter a valid number for #2.")
    Check2 = True
End If

If Not Check1 = True And Not Check2 = True Then
    Answer = Num1 + Num2
    txtAnswer.Text = Answer
End If
End Sub

Private Sub cmdClear_Click()
txtNum1.Text = ""
txtNum2.Text = ""
txtAnswer.Text = ""
End Sub

Private Sub cmdCos_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Double
Dim Pi As Double
Dim Radian As Double

Pi = 4 * Atn(1)
Radian = Pi / 180

If Not txtNum1.Text = "" And txtNum2.Text = "" Then
    Num1 = txtNum1.Text
    Num1 = Num1 * Radian
    Num1 = Cos(Num1)
    Answer = Num1
    Else
    txtNum1.Text = ""
    txtNum2.Text = ""
    MsgBox ("Please only enter 1 number!")
End If

txtAnswer.Text = Answer
End Sub

Private Sub cmdDivide_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Single
Dim Check1 As Boolean
Dim Check2 As Boolean

If Not txtNum1.Text = "" Then
    Num1 = txtNum1.Text
    Else
    MsgBox ("Please enter a valid number for #1.")
    Check1 = True
End If
If Not txtNum2.Text = "" Then
    Num2 = txtNum2.Text
    Else
    MsgBox ("Please enter a valid number for #2.")
    Check2 = True
End If
If txtNum2.Text = "0" Then
    MsgBox ("You can't divide by 0!")
    Check2 = True
End If

If Not Check1 = True And Not Check2 = True Then
    Answer = Num1 / Num2
    txtAnswer.Text = Answer
End If
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLog_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Single
Dim Check1 As Boolean
Dim Check2 As Boolean

If Not txtNum1.Text = "" Then
    Num1 = txtNum1.Text
    Else
    MsgBox ("Please enter a valid number for #1.")
    Check1 = True
End If
If Not txtNum2.Text = "" Then
    Num2 = txtNum2.Text
    Else
    MsgBox ("Please enter a valid number for #2.")
    Check2 = True
End If

If Not Check1 = True And Not Check2 = True Then
    Answer = Log(Num2) / Log(Num1)
    txtAnswer.Text = Answer
End If
End Sub

Private Sub cmdLogE_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As Double

If Not txtNum1.Text = "" And txtNum2.Text = "" Then
    Num1 = txtNum1.Text
    Answer = Log(Num1)
    Else
    txtNum1.Text = ""
    txtNum2.Text = ""
    MsgBox ("Please only enter 1 number!")
End If

txtAnswer.Text = Answer
End Sub

Private Sub cmdMultiply_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Single
Dim Check1 As Boolean
Dim Check2 As Boolean

If Not txtNum1.Text = "" Then
    Num1 = txtNum1.Text
    Else
    MsgBox ("Please enter a valid number for #1.")
    Check1 = True
End If
If Not txtNum2.Text = "" Then
    Num2 = txtNum2.Text
    Else
    MsgBox ("Please enter a valid number for #2.")
    Check2 = True
End If

If Not Check1 = True And Not Check2 = True Then
    Answer = Num1 * Num2
    txtAnswer.Text = Answer
End If
End Sub

Private Sub cmdPower_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Single
Dim Check1 As Boolean
Dim Check2 As Boolean

If Not txtNum1.Text = "" Then
    Num1 = txtNum1.Text
    Else
    MsgBox ("Please enter a valid number for #1.")
    Check1 = True
End If
If Not txtNum2.Text = "" Then
    Num2 = txtNum2.Text
    Else
    MsgBox ("Please enter a valid number for #2.")
    Check2 = True
End If

If Not Check1 = True And Not Check2 = True Then
    Answer = Num1 ^ Num2
    txtAnswer.Text = Answer
End If
End Sub

Private Sub cmdSin_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Double
Dim Pi As Double
Dim Radian As Double

Pi = 4 * Atn(1)
Radian = Pi / 180

If Not txtNum1.Text = "" And txtNum2.Text = "" Then
    Num1 = txtNum1.Text
    Num1 = Num1 * Radian
    Num1 = Sin(Num1)
    Answer = Num1
    Else
    txtNum1.Text = ""
    txtNum2.Text = ""
    MsgBox ("Please only enter 1 number!")
End If

txtAnswer.Text = Answer
End Sub

Private Sub cmdSquared_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Single

If Not txtNum1.Text = "" And txtNum2.Text = "" Then
    Num1 = txtNum1.Text
    Num1 = Num1 ^ 2
    Answer = Num1
    Else
    txtNum1.Text = ""
    txtNum2.Text = ""
    MsgBox ("Please only enter 1 number!")
End If

txtAnswer.Text = Answer
End Sub

Private Sub cmdSquareRoot_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Double

If Not txtNum1.Text = "" And txtNum2.Text = "" Then
    Num1 = txtNum1.Text
    Num1 = Sqr(Num1)
    Answer = Num1
    Else
    txtNum1.Text = ""
    txtNum2.Text = ""
    MsgBox ("Please only enter 1 number!")
End If

txtAnswer.Text = Answer
End Sub

Private Sub cmdSubtract_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Single
Dim Check1 As Boolean
Dim Check2 As Boolean

If Not txtNum1.Text = "" Then
    Num1 = txtNum1.Text
    Else
    MsgBox ("Please enter a valid number for #1.")
    Check1 = True
End If
If Not txtNum2.Text = "" Then
    Num2 = txtNum2.Text
    Else
    MsgBox ("Please enter a valid number for #2.")
    Check2 = True
End If

If Not Check1 = True And Not Check2 = True Then
    Answer = Num1 - Num2
    txtAnswer.Text = Answer
End If
End Sub

Private Sub cmdTan_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Double
Dim Pi As Double
Dim Radian As Double

Pi = 4 * Atn(1)
Radian = Pi / 180

If Not txtNum1.Text = "" And txtNum2.Text = "" Then
    Num1 = txtNum1.Text
    Num1 = Num1 * Radian
    Num1 = Tan(Num1)
    Answer = Num1
    Else
    txtNum1.Text = ""
    txtNum2.Text = ""
    MsgBox ("Please only enter 1 number!")
End If

txtAnswer.Text = Answer
End Sub

Private Sub lblTitle_Click()
MsgBox ("Made by Jack Margeson for Computer Programming 1, Summer of 2018.")
End Sub

Private Sub txtNum1_Change()
If Not IsNumeric(txtNum1.Text) And Not txtNum1.Text = "" Then
    txtNum1.Text = ""
    MsgBox ("This isn't a number!")
End If
End Sub

Private Sub txtNum2_Change()
If Not IsNumeric(txtNum2.Text) And Not txtNum2.Text = "" Then
    txtNum2.Text = ""
    MsgBox ("This isn't a number!")
End If
End Sub
