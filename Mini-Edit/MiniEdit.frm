VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MiniEdit 
   Caption         =   "Mini-Edit"
   ClientHeight    =   6540
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cndOpenFile 
      Left            =   180
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   180
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6555
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10155
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo Text"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
         Begin VB.Menu mnuMSSansSerif 
            Caption         =   "MS Sans Serif (default)"
         End
         Begin VB.Menu mnuTimesNewRoman 
            Caption         =   "Times New Roman"
         End
         Begin VB.Menu mnuArial 
            Caption         =   "Arial"
         End
         Begin VB.Menu mnuComicSans 
            Caption         =   "Comic Sans"
         End
      End
      Begin VB.Menu mnuFontColor 
         Caption         =   "Font Color"
         Begin VB.Menu mnuFontColorBlack 
            Caption         =   "Black (default)"
         End
         Begin VB.Menu mnuFontColorWhite 
            Caption         =   "White"
         End
         Begin VB.Menu mnuFontColorBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuFontColorRed 
            Caption         =   "Red"
         End
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "Font Size"
         Begin VB.Menu mnuFontSizeTiny8 
            Caption         =   "Tiny"
         End
         Begin VB.Menu mnuTextSizeSmall12 
            Caption         =   "Small (default)"
         End
         Begin VB.Menu mnuTextSizeMedium14 
            Caption         =   "Medium"
         End
         Begin VB.Menu mnuTextSizeLarge18 
            Caption         =   "Large"
         End
      End
      Begin VB.Menu mnuBackgroundColor 
         Caption         =   "Background Color"
         Begin VB.Menu mnuBackgroundColorWhite 
            Caption         =   "White (default)"
         End
         Begin VB.Menu mnuBackgroundColorBlack 
            Caption         =   "Black"
         End
         Begin VB.Menu mnuBackgroundColorBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuBackgroundColorRed 
            Caption         =   "Red"
         End
      End
   End
End
Attribute VB_Name = "MiniEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BodyText As String
Dim OldBodyText As String
Dim UserFont As String
Dim FontColor As String
Dim BackgroundColor As String
Dim UserFontSize As String
Dim FilePath As String
Dim FormatPath As String

Private Sub Form_Load()
MiniEdit.Visible = True
txtText.Visible = True
OldBodyText = ""
FilePath = ""
UserFont = "MS Sans Serif"
FontColor = "2"
UserFontSize = "12"
BackgroundColor = "1"
txtText.Font = "MS Sans Serif"
txtText.FontSize = "12"
End Sub

Private Sub Form_Resize()
txtText.Width = MiniEdit.ScaleWidth
txtText.Height = MiniEdit.ScaleHeight
End Sub

Sub SaveAsFile()
Dim Done As Boolean
Dim FileName As String
Dim Answer As String
Done = False

Do While Done = False
    cndSaveAs.Filter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cndSaveAs.ShowSave
    FilePath = cndSaveAs.FileName
    cndSaveAs.FileName = ""
 
    If FilePath = "" Then
        txtText.SetFocus
        Exit Do
    Else
        Answer = MsgBox(FilePath, vbOKCancel, "The file will be saved to:")
        If Answer = vbOK Then
            Open FilePath For Output As #1
                Print #1, txtText
            Close #1
            Done = True
            MiniEdit.Caption = "Mini-Edit - " + cndSaveAs.FileTitle
            FormatSave
        End If
    End If
Loop
End Sub

Sub SaveFile()
If FilePath = "" Then
    SaveAsFile
Else
    Open FilePath For Output As #1
        Print #1, txtText
    Close #1
    FormatSave
    OldBodyText = BodyText
End If
End Sub

Sub FormatSave()
Dim FormatFileSave As String

FormatFileSave = UserFont + vbCrLf _
+ UserFontSize + vbCrLf _
+ FontColor + vbCrLf _
+ BackgroundColor

FormatPath = FilePath + ".jfmtg"
Open FormatPath For Output As #1
    Print #1, FormatFileSave
Close #1

End Sub

Sub OpenFile()
Dim Done As Boolean
Dim Answer As String
Dim FileName As String
Done = False

Do While Done = False
    cndOpenFile.Filter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cndOpenFile.ShowOpen
    FilePath = cndOpenFile.FileName
    cndOpenFile.FileName = ""
    If FilePath = "" Then
        Exit Do
        txtText.SetFocus
    Else
        Answer = MsgBox(FilePath, vbYesNo, "Would you like to open this file?")
        If Answer = vbYes Then
            Open FilePath For Input As #1
                Dim FileSize As Long
                FileSize = LOF(1)
                txtText = Input(FileSize, 1)
            Close #1
            OldBodyText = txtText
            Done = True
            MiniEdit.Caption = "Mini-Edit - " + cndOpenFile.FileTitle
            FormatOpen
        End If
    End If
Loop
End Sub

Sub FormatOpen()

Dim CurrentLine As String
Dim i As Integer

i = 1
FormatPath = FilePath + ".jfmtg"

Open FormatPath For Input As #1
    Do While Not EOF(1)
        Line Input #1, CurrentLine
        
        If i = 1 Then
            txtText.Font = CurrentLine
            UserFont = CurrentLine
            i = i + 1
            
        ElseIf i = 2 Then
            txtText.FontSize = CurrentLine
            UserFontSize = CurrentLine
            i = i + 1
            
        ElseIf i = 3 Then
            If CurrentLine = 1 Then
                txtText.ForeColor = vbWhite
                FontColor = 1
            ElseIf CurrentLine = 2 Then
                txtText.ForeColor = vbBlack
                FontColor = 2
            ElseIf CurrentLine = 3 Then
                txtText.ForeColor = vbRed
                FontColor = 3
            ElseIf CurrentLine = 4 Then
                txtText.ForeColor = vbBlue
                FontColor = 4
            End If
            i = i + 1
            
        ElseIf i = 4 Then
            If CurrentLine = 1 Then
                txtText.BackColor = vbWhite
                BackgroundColor = 1
            ElseIf CurrentLine = 2 Then
                txtText.BackColor = vbBlack
                BackgroundColor = 2
            ElseIf CurrentLine = 3 Then
                txtText.BackColor = vbRed
                BackgroundColor = 3
            ElseIf CurrentLine = 4 Then
                txtText.BackColor = vbBlue
                BackgroundColor = 4
            End If
            i = 1 + 1
        End If
    Loop
Close #1
End Sub

Private Sub mnuArial_Click()
txtText.Font = "Arial"
UserFont = "Arial"
End Sub

Private Sub mnuBackgroundColorBlack_Click()
txtText.BackColor = vbBlack
BackgroundColor = "2"
End Sub

Private Sub mnuBackgroundColorBlue_Click()
txtText.BackColor = vbBlue
BackgroundColor = "4"
End Sub

Private Sub mnuBackgroundColorRed_Click()
txtText.BackColor = vbRed
BackgroundColor = "3"
End Sub

Private Sub mnuBackgroundColorWhite_Click()
txtText.BackColor = vbWhite
BackgroundColor = "1"
End Sub

Private Sub mnuComicSans_Click()
txtText.Font = "Comic Sans MS"
UserFont = "Comic Sans MS"
End Sub

Private Sub mnuCopy_Click()
Clipboard.SetText txtText.SelText
mnuPaste.Enabled = True
End Sub

Private Sub mnuCut_Click()
Clipboard.SetText txtText.SelText
txtText.SelText = ""
mnuPaste.Enabled = True
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuFontColorBlack_Click()
txtText.ForeColor = vbBlack
FontColor = "2"
End Sub

Private Sub mnuFontColorBlue_Click()
txtText.ForeColor = vbBlue
FontColor = "4"
End Sub

Private Sub mnuFontColorRed_Click()
txtText.ForeColor = vbRed
FontColor = "3"
End Sub

Private Sub mnuFontColorWhite_Click()
txtText.ForeColor = vbWhite
FontColor = "1"
End Sub

Private Sub mnuFontSizeTiny8_Click()
txtText.FontSize = "8"
UserFontSize = "8"
End Sub

Private Sub mnuMSSansSerif_Click()
txtText.Font = "MS Sans Serif"
UserFont = "MS Sans Serif"
End Sub

Private Sub mnuOpen_Click()
OpenFile
End Sub

Private Sub mnuPaste_Click()
txtText.SelText = Clipboard.GetText()
End Sub

Private Sub mnuSave_Click()
Dim Answer As String
Answer = MsgBox("Are you sure you want to save?", vbYesNo)
If Answer = vbYes Then
    SaveFile
End If
End Sub

Private Sub mnuSaveAs_Click()
SaveAsFile
End Sub

Private Sub mnuSelectAll_Click()
txtText.SelStart = 0
txtText.SelLength = Len(txtText.Text)
End Sub

Private Sub mnuTextSizeLarge18_Click()
txtText.FontSize = "18"
UserFontSize = "18"
End Sub

Private Sub mnuTextSizeMedium14_Click()
txtText.FontSize = "14"
UserFontSize = "14"
End Sub

Private Sub mnuTextSizeSmall12_Click()
txtText.FontSize = "12"
UserFontSize = "12"
End Sub

Private Sub mnuTimesNewRoman_Click()
txtText.Font = "Times New Roman"
UserFont = "Times New Roman"
End Sub

Private Sub mnuUndo_Click()
SendKeys ("^Z")
End Sub

Private Sub txtText_Change()
BodyText = txtText.Text
End Sub

Private Sub mnuNew_Click()
Dim Answer As String
    If BodyText <> OldBodyText Then
    Answer = MsgBox("Are you sure you want to create a new document? All changes will be lost.", vbOKCancel)
        If Answer = vbOK Then
            txtText.Text = ""
            BodyText = txtText.Text
            OldBodyText = ""
            FilePath = ""
            UserFont = "MS Sans Serif"
            FontColor = "2"
            UserFontSize = "12"
            BackgroundColor = "1"
            txtText.Font = "MS Sans Serif"
            txtText.FontSize = "12"
            txtText.ForeColor = vbBlack
            txtText.BackColor = vbWhite
            MiniEdit.Caption = "Mini-Edit"
        End If
    Else
    txtText.Text = ""
    BodyText = txtText.Text
    OldBodyText = ""
    FilePath = ""
    UserFont = "MS Sans Serif"
    FontColor = "2"
    UserFontSize = "12"
    BackgroundColor = "1"
    txtText.Font = "MS Sans Serif"
    txtText.FontSize = "12"
    txtText.ForeColor = vbBlack
    txtText.BackColor = vbWhite
    MiniEdit.Caption = "Mini-Edit"
    End If
End Sub
