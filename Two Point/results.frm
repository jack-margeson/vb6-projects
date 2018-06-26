VERSION 5.00
Begin VB.Form frmResults 
   Caption         =   "Two Point - Results"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtDistance2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8220
      TabIndex        =   17
      Text            =   "txtDistance2"
      Top             =   1380
      Width           =   1155
   End
   Begin VB.TextBox txtMidpoint2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Text            =   "txtMidpoint2"
      Top             =   1380
      Width           =   1275
   End
   Begin VB.TextBox txtYInt2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Text            =   "txtYInt2"
      Top             =   1380
      Width           =   1275
   End
   Begin VB.TextBox txtSlope2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3660
      TabIndex        =   14
      Text            =   "txtSlope2"
      Top             =   1380
      Width           =   795
   End
   Begin VB.TextBox txtEquation2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Text            =   "txtEquation2"
      Top             =   1380
      Width           =   2055
   End
   Begin VB.TextBox txtDistance1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   8220
      TabIndex        =   12
      Text            =   "txtDistance1"
      Top             =   660
      Width           =   1155
   End
   Begin VB.TextBox txtMidpoint1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Text            =   "txtMidpoint1"
      Top             =   660
      Width           =   1275
   End
   Begin VB.TextBox txtYInt1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Text            =   "txtYInt1"
      Top             =   660
      Width           =   1275
   End
   Begin VB.TextBox txtSlope1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3660
      TabIndex        =   9
      Text            =   "txtSlope1"
      Top             =   660
      Width           =   795
   End
   Begin VB.TextBox txtEquation1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Text            =   "txtEquation1"
      Top             =   660
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return and Clear"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblDistance 
      Alignment       =   2  'Center
      Caption         =   "Distance"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblMidpoint 
      Alignment       =   2  'Center
      Caption         =   "Midpoint"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblYInt 
      Alignment       =   2  'Center
      Caption         =   "Y-Intercept"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   5
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label lblSlope 
      Alignment       =   2  'Center
      Caption         =   "Slope"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblEquation 
      Alignment       =   2  'Center
      Caption         =   "Equation"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblLine2 
      Alignment       =   2  'Center
      Caption         =   "Line 2"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label lblLine1 
      Alignment       =   2  'Center
      Caption         =   "Line 1"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
End
Attribute VB_Name = "frmResults"
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
Dim Slope1 As Single
Dim Slope2 As Single
Dim Yint1 As Single
Dim Yint2 As Single
Dim Distance1 As Single
Dim Distance2 As Single
Dim Middlex1 As Single
Dim Middlex2 As Single
Dim Middley1 As Single
Dim Middley2 As Single
Dim Intersectx As Single
Dim Intersecty As Single
Dim Var1 As Single
Dim Var2 As Single

Private Sub cmdReturn_Click()
frmResults.Hide
frmTwoPoint.Show

Slope1 = 0
Slope2 = 0
Yint1 = 0
Yint2 = 0
Distance1 = 0
Distance2 = 0
Middlex1 = 0
Middlex2 = 0
Middley1 = 0
Middley2 = 0
Intersectx = 0
Intersecty = 0
Var1 = 0
Var2 = 0
End Sub

Private Sub Form_Activate()
'Line 1 results:
If x2 - x1 <> 0 Then
    'Slope
    Slope1 = (y2 - y1) / (x2 - x1)
    txtSlope1.Text = Slope1
    
    'Y Intercept
    Yint1 = y1 - Slope1 * x1
    txtYInt1.Text = Yint1
    
    'Equation
    txtEquation1.Text = "y = " + Format(Slope1, "fixed") + "x + " + Format(Yint1, "fixed")
    
    'Distance
    Distance1 = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    txtDistance1.Text = Distance1
    
    'Midpoint
    Middlex1 = (x1 + x2) / 2
    Middley1 = (y1 + y2) / 2
    txtMidpoint1 = "(" + Format(Middlex1, "fixed") + ", " + Format(Middley1, "fixed") + ")"
Else
    txtSlope1.Text = "N/A"
    txtYInt1.Text = "N/A"
    txtEquation1.Text = "N/A"
    Distance1 = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    txtDistance1.Text = Distance1
    Middlex1 = (x1 + x2) / 2
    Middley1 = (y1 + y2) / 2
    txtMidpoint1 = "(" + Format(Middlex1, "fixed") + ", " + Format(Middley1, "fixed") + ")"
End If

'Line 2 results:
If x4 - x3 <> 0 Then
    'Slope
    Slope2 = (y4 - y3) / (x4 - x3)
    txtSlope2 = Slope2
    
    'Y Intercept
    Yint2 = y3 - Slope2 * x3
    txtYInt2 = Yint2
    
    'Equation
    txtEquation2 = "y = " + Format(Slope2, "fixed") + "x + " + Format(Yint2, "fixed")
    
    'Distance
    Distance2 = Sqr((x4 - x3) ^ 2 + (y4 - y3) ^ 2)
    txtDistance2.Text = Distance2
    
    'Midpoint
    Middlex2 = (x3 + x4) / 2
    Middley2 = (y3 + y4) / 2
    txtMidpoint2 = "(" + Format(Middlex2, "fixed") + ", " + Format(Middley2, "fixed") + ")"
Else
    txtSlope2.Text = "N/A"
    txtYInt2.Text = "N/A"
    txtEquation2.Text = "N/A"
    Distance2 = Sqr((x4 - x3) ^ 2 + (y4 - y3) ^ 2)
    txtDistance2.Text = Distance2
    Middlex2 = (x3 + x4) / 2
    Middley2 = (y3 + y4) / 2
    txtMidpoint2 = "(" + Format(Middlex2, "fixed") + ", " + Format(Middley2, "fixed") + ")"
End If
End Sub
