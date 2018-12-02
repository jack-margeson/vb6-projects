VERSION 5.00
Begin VB.Form frmAmortizationTable 
   Caption         =   "Amortization Table"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   7020
      TabIndex        =   3
      Top             =   2100
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   2100
      Width           =   1815
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate and Results"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5580
      TabIndex        =   1
      Top             =   1500
      Width           =   2775
   End
   Begin VB.Frame fraEnterValues 
      Caption         =   "Enter Values:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtExtraPayment 
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         Text            =   "txtExtraPayment"
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtYears 
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Text            =   "txtYears"
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtYearlyInterest 
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Text            =   "txtYearlyInterest"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtLoanAmount 
         Height          =   315
         Left            =   600
         TabIndex        =   8
         Text            =   "txtLoanAmount"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblExtraMonthlyPayment 
         Caption         =   "Extra Monthly Payment:"
         Height          =   495
         Left            =   2580
         TabIndex        =   7
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label lblYears 
         Caption         =   "Years :"
         Height          =   435
         Left            =   420
         TabIndex        =   6
         Top             =   1500
         Width           =   1935
      End
      Begin VB.Label lblYearlyIntrest 
         Caption         =   "Yearly Interest:"
         Height          =   375
         Left            =   2580
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Loan Amount:"
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1875
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Amortization Table"
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   5400
      TabIndex        =   12
      Top             =   240
      Width           =   3195
   End
End
Attribute VB_Name = "frmAmortizationTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Input form variables.
Public loanamount As Currency
Public yearlyinterest As Single
Public years As Integer
Public extrapayment As Currency

'Calculation variables.
Dim payments As Integer
Dim monthlyinterest As Currency
Dim monthlyrate As Currency
Dim currentamount As Currency
Dim paymentnumber As Integer
Dim totalinterest As Currency
Dim monthlyprincipal As Currency
Dim yearnumber As Integer

'Result form variables.
Dim monthlypayment As Currency
Dim hsblength As Single
Dim tabletext As String
Dim finaltable(360) As String
Private Sub cmdCalculate_Click()
If Not txtLoanAmount = "" And Not txtYearlyInterest = "" And Not txtYears = "" And Not txtExtraPayment = "" Then
    If Not Val(txtLoanAmount) = 0 And Not Val(txtYearlyInterest) = 0 And Not Val(txtYears) = 0 Then
        loanamount = Val(txtLoanAmount)
        yearlyinterest = Val(txtYearlyInterest)
        years = Val(txtYears)
        extrapayment = Val(txtExtraPayment)

        frmResults.loanamount = loanamount
        frmResults.yearlyinterest = yearlyinterest
        frmResults.years = years
        frmResults.extrapayment = extrapayment

        frmAmortizationTable.Hide
        frmResults.Show
    Else
        MsgBox ("Only use 0 when filling out the extra monthly payment box if necessary.")
    End If
Else
    MsgBox ("Please fill in all data entries!")
End If
End Sub

Private Sub cmdClear_Click()
Clear
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Activate()
Clear
End Sub

Sub Clear()
txtExtraPayment = ""
txtYears = ""
txtYearlyInterest = ""
txtLoanAmount = ""
End Sub

Private Sub lblTitle_Click()
MsgBox ("Made by Jack Margeson. Computer Programming 1 Credix Flex, Summer 2018.")
End Sub
