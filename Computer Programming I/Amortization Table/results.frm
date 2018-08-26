VERSION 5.00
Begin VB.Form frmResults 
   Caption         =   "Results"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraResults 
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10275
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear and Return"
         BeginProperty Font 
            Name            =   "Source Sans Pro"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5700
         TabIndex        =   11
         Top             =   1560
         Width           =   2715
      End
      Begin VB.TextBox txtPayment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3300
         TabIndex        =   10
         Text            =   "txtPayment"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtTable 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Text            =   "txtTable"
         Top             =   780
         Width           =   9855
      End
      Begin VB.HScrollBar hsbTable 
         Height          =   195
         Left            =   240
         Min             =   1
         TabIndex        =   1
         Top             =   1200
         Value           =   1
         Width           =   9855
      End
      Begin VB.Label lblPayment 
         Caption         =   "Monthly Payment:"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label lblPaymentNumber 
         Alignment       =   2  'Center
         Caption         =   "Payment Number"
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label lblYearNumber 
         Alignment       =   2  'Center
         Caption         =   "Year Number"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label lblLeftToPay 
         Alignment       =   2  'Center
         Caption         =   "Left To Pay"
         Height          =   375
         Left            =   3300
         TabIndex        =   6
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label lblTotalInterest 
         Alignment       =   2  'Center
         Caption         =   "Total Interest"
         Height          =   315
         Left            =   5160
         TabIndex        =   5
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label lblMonthlyInterest 
         Alignment       =   2  'Center
         Caption         =   "Monthly Interest"
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label lblMonthlyPrincipal 
         Alignment       =   2  'Center
         Caption         =   "Monthly Principal"
         Height          =   375
         Left            =   8580
         TabIndex        =   3
         Top             =   420
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmResults"
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
Dim monthlyrate As Single
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

Private Sub cmdClear_Click()
Clear
frmResults.Hide
frmAmortizationTable.Show
End Sub

Private Sub Form_Activate()
Clear

'Basic calculations.
monthlyrate = yearlyinterest / 1200
payments = years * 12

'Monthly payment.
extrapayment = extrapayment * payments
monthlypayment = (loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ (-payments))) + extrapayment
txtPayment = Format(monthlypayment, "Currency")

'Table information.
currentamount = loanamount

For paymentnumber = 1 To payments
    hsblength = hsblength + 1
    
    If currentamount < monthlypayment Then
        monthlypayment = currentamount
        'Payment Number, Year Number, Left to Pay (balance), Total Interest, Monthly Interest, Monthly Principal.
        yearnumber = Int(paymentnumber / 12) + 1
        monthlyinterest = currentamount * monthlyrate
        totalinterest = totalinterest + monthlyinterest
        monthlyprincipal = monthlypayment - monthlyinterest
        currentamount = currentamount + monthlyinterest - monthlypayment
        
        If paymentnumber Mod 12 = 0 Then
            yearnumber = yearnumber - 1
        End If
        
        tabletext = "              " + Format(paymentnumber, "####")
        tabletext = tabletext + "                                " + Format(yearnumber, "#0")
        tabletext = tabletext + "                          " + "$0.00"
        tabletext = tabletext + "                       " + Format(totalinterest, "Currency")
        tabletext = tabletext + "                             " + Format(monthlyinterest, "Currency")
        tabletext = tabletext + "                             " + Format(monthlyprincipal, "Currency")
        
        finaltable(paymentnumber) = tabletext
    Else
        yearnumber = Int(paymentnumber / 12) + 1
        monthlyinterest = currentamount * monthlyrate
        totalinterest = totalinterest + monthlyinterest
        monthlyprincipal = monthlypayment - monthlyinterest
        currentamount = currentamount + monthlyinterest - monthlypayment
        
        If paymentnumber Mod 12 = 0 Then
            yearnumber = yearnumber - 1
        End If
        
        tabletext = "              " + Format(paymentnumber, "####")
        tabletext = tabletext + "                                " + Format(yearnumber, "#0")
        tabletext = tabletext + "                          " + Format(currentamount, "Currency")
        tabletext = tabletext + "                       " + Format(totalinterest, "Currency")
        tabletext = tabletext + "                             " + Format(monthlyinterest, "Currency")
        tabletext = tabletext + "                             " + Format(monthlyprincipal, "Currency")
        
        finaltable(paymentnumber) = tabletext
    End If
Next paymentnumber

hsbTable.Max = hsblength
txtTable = finaltable(1)
End Sub

Sub Clear()
txtPayment = ""
txtTable = ""
hsbTable.Value = 1
hsbTable.Max = 1
hsbTable.Min = 1
End Sub

Private Sub hsbTable_Change()
txtTable = finaltable(hsbTable.Value)
End Sub
