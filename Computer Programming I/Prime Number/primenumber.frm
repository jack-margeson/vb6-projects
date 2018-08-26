VERSION 5.00
Begin VB.Form PrimeNumber 
   Caption         =   "Prime Number"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumberOfFactors2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4380
      TabIndex        =   23
      Top             =   4680
      Width           =   1635
   End
   Begin VB.TextBox txtNumberOfFactors1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1260
      TabIndex        =   22
      Top             =   4680
      Width           =   1635
   End
   Begin VB.TextBox txtSum2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4140
      TabIndex        =   21
      Top             =   4080
      Width           =   1875
   End
   Begin VB.TextBox txtSum1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1020
      TabIndex        =   20
      Top             =   4080
      Width           =   1875
   End
   Begin VB.TextBox txtLCM 
      Enabled         =   0   'False
      Height          =   435
      Left            =   3900
      TabIndex        =   15
      Top             =   6060
      Width           =   1935
   End
   Begin VB.TextBox txtGCF 
      Enabled         =   0   'False
      Height          =   435
      Left            =   1080
      TabIndex        =   14
      Top             =   6060
      Width           =   1935
   End
   Begin VB.TextBox txtFactors2 
      Enabled         =   0   'False
      Height          =   675
      Left            =   3240
      ScrollBars      =   1  'Horizontal
      TabIndex        =   11
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox txtFactors1 
      Enabled         =   0   'False
      Height          =   675
      Left            =   120
      ScrollBars      =   1  'Horizontal
      TabIndex        =   10
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ListBox lstFactors2 
      Height          =   1815
      Left            =   4260
      TabIndex        =   9
      Top             =   1320
      Width           =   1755
   End
   Begin VB.ListBox lstFactors1 
      Height          =   1815
      Left            =   1140
      TabIndex        =   8
      Top             =   1320
      Width           =   1755
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3180
      TabIndex        =   5
      Top             =   6720
      Width           =   2715
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   300
      TabIndex        =   4
      Top             =   6720
      Width           =   2715
   End
   Begin VB.TextBox txtNumber2 
      Height          =   435
      Left            =   3240
      TabIndex        =   3
      Top             =   660
      Width           =   2775
   End
   Begin VB.TextBox txtNumber1 
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   2835
   End
   Begin VB.Label lblNotPrime2 
      Alignment       =   2  'Center
      Caption         =   "Not Prime"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblNotPrime1 
      Alignment       =   2  'Center
      Caption         =   "Not Prime"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   26
      Top             =   5280
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblPrime2 
      Alignment       =   2  'Center
      Caption         =   "Prime"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   25
      Top             =   5280
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblPrime1 
      Alignment       =   2  'Center
      Caption         =   "Prime"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblNumberOfFactors2 
      Caption         =   "# of Factors"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3300
      TabIndex        =   19
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblNumberOfFactors1 
      Caption         =   "# of Factors"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   180
      TabIndex        =   18
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblSum2 
      Caption         =   "Sum"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   17
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblSum1 
      Caption         =   "Sum"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   16
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblLCM 
      Caption         =   "LCM"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3180
      TabIndex        =   13
      Top             =   6060
      Width           =   735
   End
   Begin VB.Label lblGCF 
      Caption         =   "GCF"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   12
      Top             =   6060
      Width           =   735
   End
   Begin VB.Label lblFactors2 
      Caption         =   "Factors"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   7
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label lblFactors1 
      Caption         =   "Factors"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label lblNumber2 
      Alignment       =   2  'Center
      Caption         =   "Number 2"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3660
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblNumber1 
      Alignment       =   2  'Center
      Caption         =   "Number 1"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "PrimeNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number1 As Long
Dim Number2 As Long
Dim Number1Check As Boolean
Dim Number2Check As Boolean

Private Sub cmdClear_Click()
Clear
txtNumber1.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
lblPrime1.BackColor = vbGreen
lblPrime2.BackColor = vbGreen
lblNotPrime1.BackColor = vbRed
lblNotPrime2.BackColor = vbRed
End Sub

Private Sub txtNumber1_KeyPress(KeyAscii As Integer)
Dim i As Long
Dim Sum1 As Long
Dim Factors1 As String
Dim NumberOfFactors1 As Long

If KeyAscii = 13 Then
    If Not IsNumeric(txtNumber1) Then
        MsgBox ("Please enter a valid number.")
    Else
        Number1 = Val(txtNumber1)
        
        For i = 1 To Number1
            If Number1 Mod i = 0 Then
                Factors1 = Factors1 + Str(i) + ""
                lstFactors1.AddItem (Str(i))
                Sum1 = Sum1 + i
                NumberOfFactors1 = NumberOfFactors1 + 1
            End If
        Next i
        txtFactors1.Text = Factors1
        txtSum1.Text = Sum1
        txtNumberOfFactors1.Text = NumberOfFactors1
        If Sum1 = Number1 + 1 Then
            lblPrime1.Visible = True
        Else
            lblNotPrime1.Visible = True
        End If
        Number1Check = True
        txtNumber2.SetFocus
        txtNumber1.Enabled = False
        CheckBoth
    End If
End If
End Sub

Private Sub txtNumber2_KeyPress(KeyAscii As Integer)
Dim i As Long
Dim Sum2 As Long
Dim Factors2 As String
Dim NumberOfFactors2 As Long

If KeyAscii = 13 Then
    If Not IsNumeric(txtNumber2) Then
        MsgBox ("Please enter a valid number.")
    Else
        Number2 = Val(txtNumber2)
        
        For i = 1 To Number2
            If Number2 Mod i = 0 Then
                Factors2 = Factors2 + Str(i) + ""
                lstFactors2.AddItem (Str(i))
                Sum2 = Sum2 + i
                NumberOfFactors2 = NumberOfFactors2 + 1
            End If
        Next i
        txtFactors2.Text = Factors2
        txtSum2.Text = Sum2
        txtNumberOfFactors2.Text = NumberOfFactors2
        If Sum2 = Number2 + 1 Then
            lblPrime2.Visible = True
        Else
            lblNotPrime2.Visible = True
        End If
        Number2Check = True
        cmdClear.SetFocus
        txtNumber2.Enabled = False
        CheckBoth
    End If
End If
End Sub

Sub CheckBoth()
Dim i As Integer
Dim GCF As Long
Dim LCM As Long

If Number1Check = True And Number2Check = True Then
    For i = 1 To Number1
        If Number1 Mod i = 0 And Number2 Mod i = 0 Then
            GCF = i
        End If
    Next i
    txtGCF.Text = GCF
    
    LCM = (Number1 / GCF) * Number2
    txtLCM.Text = LCM
End If
End Sub

Sub Clear()
txtNumber1.Enabled = True
txtNumber2.Enabled = True
txtNumber1.Text = ""
txtNumber2.Text = ""
txtFactors1.Text = ""
txtFactors2.Text = ""
txtSum1.Text = ""
txtSum2.Text = ""
txtNumberOfFactors1.Text = ""
txtNumberOfFactors2.Text = ""
txtGCF.Text = ""
txtLCM.Text = ""
lblPrime1.Visible = False
lblPrime2.Visible = False
lblNotPrime1.Visible = False
lblNotPrime2.Visible = False
Number1Check = False
Number2Check = False
lstFactors1.Clear
lstFactors2.Clear

End Sub

