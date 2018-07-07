VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Configure..."
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbB 
      Height          =   255
      Left            =   180
      Max             =   255
      TabIndex        =   7
      Top             =   1680
      Value           =   1
      Width           =   3975
   End
   Begin VB.HScrollBar hsbG 
      Height          =   255
      Left            =   180
      Max             =   255
      TabIndex        =   6
      Top             =   1080
      Value           =   1
      Width           =   3975
   End
   Begin VB.HScrollBar hsbR 
      Height          =   255
      Left            =   180
      Max             =   255
      TabIndex        =   5
      Top             =   420
      Value           =   1
      Width           =   3975
   End
   Begin VB.HScrollBar hsbS 
      Height          =   255
      Left            =   180
      Max             =   50
      Min             =   1
      TabIndex        =   0
      Top             =   2400
      Value           =   1
      Width           =   3975
   End
   Begin VB.Label lblColorDisplay 
      BackColor       =   &H80000007&
      Height          =   1455
      Left            =   4800
      TabIndex        =   9
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label lblSizeDisplay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   2220
      Width           =   1755
   End
   Begin VB.Label lblSize 
      Caption         =   "Size (1-50):"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblBlue 
      Caption         =   "Blue:"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label lblGreen 
      Caption         =   "Green:"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   735
   End
   Begin VB.Label lblRed 
      Caption         =   "Red:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sync variables from frmPaint.
Public ps As Single
Public pr As Single
Public pg As Single
Public pb As Single

Private Sub Form_Load()
lblSizeDisplay = ps
lblColorDisplay.BackColor = RGB(pr, pg, pb)
lblSizeDisplay.ForeColor = RGB(pr, pg, pb)
hsbS.Value = ps
hsbR.Value = pr
hsbG.Value = pg
hsbB.Value = pb
End Sub

Private Sub hsbB_Change()
pb = hsbB.Value
lblColorDisplay.BackColor = RGB(pr, pg, pb)
lblSizeDisplay.ForeColor = RGB(pr, pg, pb)
frmPaint.pb = pb
End Sub

Private Sub hsbG_Change()
pg = hsbG.Value
lblColorDisplay.BackColor = RGB(pr, pg, pb)
lblSizeDisplay.ForeColor = RGB(pr, pg, pb)
frmPaint.pg = pg
End Sub

Private Sub hsbR_Change()
pr = hsbR.Value
lblColorDisplay.BackColor = RGB(pr, pg, pb)
lblSizeDisplay.ForeColor = RGB(pr, pg, pb)
frmPaint.pr = pr
End Sub

Private Sub hsbS_Change()
ps = hsbS.Value
frmPaint.DrawWidth = ps
lblSizeDisplay = hsbS.Value
frmPaint.ps = ps
End Sub

