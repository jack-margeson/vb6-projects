VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCities 
   Caption         =   "Cities"
   ClientHeight    =   5040
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cndOpen 
      Left            =   1800
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   1800
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   10560
      TabIndex        =   17
      Top             =   4620
      Width           =   1395
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   315
      Left            =   9120
      TabIndex        =   16
      Top             =   4620
      Width           =   1395
   End
   Begin VB.ListBox lstMerge 
      Height          =   1230
      Left            =   9600
      TabIndex        =   13
      Top             =   2760
      Width           =   1875
   End
   Begin VB.ListBox lstCommon 
      Height          =   1230
      Left            =   9600
      TabIndex        =   12
      Top             =   420
      Width           =   1875
   End
   Begin VB.ListBox lstCity3 
      Height          =   3765
      Left            =   6900
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox lstCity2 
      Height          =   3765
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox lstCity1 
      Height          =   3765
      Left            =   2460
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtNewList 
      Height          =   4875
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "cities.frx":0000
      Top             =   60
      Width           =   2055
   End
   Begin VB.Label lblMergeNumber 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblMergeNumber"
      Height          =   315
      Left            =   9840
      TabIndex        =   15
      Top             =   4140
      Width           =   1455
   End
   Begin VB.Label lblCommonNumber 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblCommonNumber"
      Height          =   315
      Left            =   9780
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblMerge 
      Caption         =   "Merge:"
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label lblCommon 
      Caption         =   "Common:"
      Height          =   255
      Left            =   9600
      TabIndex        =   10
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label lblCity3Number 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblCity3Number"
      Height          =   315
      Left            =   7140
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblCity2Number 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblCity2Number"
      Height          =   315
      Left            =   4980
      TabIndex        =   8
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblCity1Number 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblCity1Number"
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblTitle3"
      Height          =   315
      Left            =   6900
      TabIndex        =   2
      Top             =   60
      Width           =   2055
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblTitle2"
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   60
      Width           =   2055
   End
   Begin VB.Label lblTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "lblTitle1"
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   60
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As...."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "Commands"
      Begin VB.Menu mnuOpenList1 
         Caption         =   "Open to List 1"
      End
      Begin VB.Menu mnuOpenList2 
         Caption         =   "Open to List 2"
      End
      Begin VB.Menu mnuOpenList3 
         Caption         =   "Open to List 3"
      End
      Begin VB.Menu mnuMerge 
         Caption         =   "Merge"
      End
      Begin VB.Menu mnuCommon 
         Caption         =   "Common"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear All"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Path As String
Dim Complete1 As Boolean
Dim Complete2 As Boolean
Dim Complete3 As Boolean

Private Sub cmdClear_Click()
ClearAll
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
ClearAll
End Sub

Private Sub mnuAbout_Click()
MsgBox ("Created by Jack Margeson. Computer Programming Credit Flex, Summer 2018.")
End Sub

Private Sub mnuClear_Click()
ClearAll
End Sub

Private Sub mnuCommon_Click()
Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer

For i = 0 To lstCity1.ListCount - 1
    For i2 = 0 To lstCity2.ListCount - 1
        For i3 = 0 To lstCity3.ListCount - 1
            If StrComp(lstCity1.List(i), lstCity2.List(i2)) = 0 And StrComp(lstCity2.List(i2), lstCity3.List(i3)) = 0 And StrComp(lstCity1.List(i), lstCity3.List(i3)) = 0 Then
                lstCommon.AddItem lstCity1.List(i)
            End If
        Next i3
    Next i2
Next i
lblCommonNumber = lstCommon.ListCount
End Sub

Private Sub mnuExit_Click()
End
End Sub

Sub ClearAll()
txtNewList = ""
lblTitle1 = ""
lblTitle2 = ""
lblTitle3 = ""
lblCity1Number = ""
lblCity2Number = ""
lblCity3Number = ""
lblCommonNumber = ""
lblMergeNumber = ""
Path = ""

lstCity1.Clear
lstCity2.Clear
lstCity3.Clear
lstCommon.Clear
lstMerge.Clear

Complete1 = False
Complete2 = False
Complete3 = False
mnuCommon.Enabled = False
mnuMerge.Enabled = False
End Sub

Sub SaveFile()
If Path = "" Then
    SaveAsFile
Else
    Open Path For Output As #1
        Print #1, txtNewList
    Close #1
End If
End Sub

Sub SaveAsFile()
Dim Done As Boolean
Dim FileName As String
Dim Answer As String

Done = False
Do While Done = False
    cndSaveAs.Filter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cndSaveAs.ShowSave
    Path = cndSaveAs.FileName
    cndSaveAs.FileName = ""
    If Path = "" Then
        txtNewList.SetFocus
        Exit Do
    Else
        Answer = MsgBox(Path, vbYesNo, "Is this the correct path?")
        If Answer = vbYes Then
            Open Path For Output As #1
                Print #1, txtNewList
            Close #1
            Done = True
         End If
    End If
Loop
End Sub

Sub OpenFile()
Dim Done As Boolean
Dim Answer As String
Dim FileName As String

Done = False
Do While Done = False
    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    If Path = "" Then
        Exit Do
        txtNewList.SetFocus
    Else
        Answer = MsgBox(Path, vbYesNo, "Do you want to open this file?")
        If Answer = vbYes Then
            Open Path For Input As #1
                Dim FileSize As Long
                FileSize = LOF(1)
                txtNewList = Input(FileSize, 1)
            Close #1
            Done = True

        End If
    End If
Loop
End Sub

Private Sub mnuMerge_Click()
Dim i As Integer
Dim i2 As Integer

For i = 0 To lstCity1.ListCount - 1
    lstMerge.AddItem lstCity1.List(i)
Next i

For i = 0 To lstCity2.ListCount - 1
    lstMerge.AddItem lstCity2.List(i)
Next i

For i = 0 To lstCity3.ListCount - 1
    lstMerge.AddItem lstCity3.List(i)
Next i

'Removes duplicates in the Merge list.
For i = 0 To lstMerge.ListCount - 1
    For i2 = i + 1 To lstMerge.ListCount - 1
        If StrComp(lstMerge.List(i), lstMerge.List(i2)) = 0 Then
            lstMerge.RemoveItem i2
        End If
    Next i2
Next i
lblMergeNumber = lstMerge.ListCount
End Sub

Private Sub mnuNew_Click()
Dim Answer As String
Answer = MsgBox("Do you want to start a new list?", vbYesNo, "Confirm:")
        If Answer = vbYes Then
            txtNewList = ""
        End If
End Sub

Private Sub mnuOpen_Click()
OpenFile
End Sub

Private Sub mnuOpenList1_Click()
Dim Done As Boolean
Dim Answer As String
Dim FileName As String
Dim FixedName As String

Done = False
Do While Done = False
    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    If Path = "" Then
        Exit Do
        txtNewList.SetFocus
    Else
        Answer = MsgBox(Path, vbYesNo, "Do you want to open this file to List 1?")
        If Answer = vbYes Then
            lstCity1.Clear
            lblTitle1 = ""
            lblCity1Number = ""
        
            Dim Line As String
            Dim i As Integer
            Open Path For Input As #1
                For i = 1 To 30
                    If EOF(1) Then
                        Exit For
                    End If
                    Line Input #1, Line
                    Line = Trim(Line)
                    If Len(Line) <> 0 Then
                        lstCity1.AddItem Line
                    End If
                Next i
            Close #1
            lblCity1Number = lstCity1.ListCount
            FixedName = cndOpen.FileTitle
            FixedName = Replace(FixedName, ".txt", "")
            lblTitle1 = FixedName
            Done = True
            Complete1 = True
            CheckComplete
        End If
    End If
Loop
End Sub

Private Sub mnuOpenList2_Click()
Dim Done As Boolean
Dim Answer As String
Dim FileName As String
Dim FixedName As String

Done = False
Do While Done = False
    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    If Path = "" Then
        Exit Do
        txtNewList.SetFocus
    Else
        Answer = MsgBox(Path, vbYesNo, "Do you want to open this file to List 2?")
        If Answer = vbYes Then
            lstCity2.Clear
            lblTitle2 = ""
            lblCity2Number = ""
            
            Dim Line As String
            Dim i As Integer
            Open Path For Input As #1
                For i = 1 To 30
                    If EOF(1) Then
                        Exit For
                    End If
                    Line Input #1, Line
                    Line = Trim(Line)
                    If Len(Line) <> 0 Then
                        lstCity2.AddItem Line
                    End If
                Next i
            Close #1
            lblCity2Number = lstCity2.ListCount
            FixedName = cndOpen.FileTitle
            FixedName = Replace(FixedName, ".txt", "")
            lblTitle2 = FixedName
            Done = True
            Complete2 = True
            CheckComplete
        End If
    End If
Loop
End Sub

Private Sub mnuOpenList3_Click()
Dim Done As Boolean
Dim Answer As String
Dim FileName As String
Dim FixedName As String

Done = False
Do While Done = False
    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    If Path = "" Then
        Exit Do
        txtNewList.SetFocus
    Else
        Answer = MsgBox(Path, vbYesNo, "Do you want to open this file to List 3?")
        If Answer = vbYes Then
            lstCity3.Clear
            lblTitle3 = ""
            lblCity3Number = ""
            
            Dim Line As String
            Dim i As Integer
            Open Path For Input As #1
                For i = 1 To 30
                    If EOF(1) Then
                        Exit For
                    End If
                    Line Input #1, Line
                    Line = Trim(Line)
                    If Len(Line) <> 0 Then
                        lstCity3.AddItem Line
                    End If
                Next i
            Close #1
            lblCity3Number = lstCity3.ListCount
            FixedName = cndOpen.FileTitle
            FixedName = Replace(FixedName, ".txt", "")
            lblTitle3 = FixedName
            Done = True
            Complete3 = True
            CheckComplete
        End If
    End If
Loop
End Sub

Sub CheckComplete()
If Complete1 = True And Complete2 = True And Complete3 = True Then
        mnuCommon.Enabled = True
End If
If Complete1 = True And Complete2 = True Or Complete1 = True And Complete3 = True Or Complete2 = True And Complete3 = True Then
    mnuMerge.Enabled = True
End If
End Sub

Private Sub mnuSave_Click()
SaveFile
End Sub

Private Sub mnuSaveAs_Click()
SaveAsFile
End Sub
