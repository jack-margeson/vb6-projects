VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPaint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Paint"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8715
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   60
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Data (*.dat)|*.dat|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cndOpen 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Data (*.dat)|*.dat|All Files (*.*)|*.*"
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
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
   End
   Begin VB.Menu mnuConfigure 
      Caption         =   "Configure"
      Begin VB.Menu mnuSizeandColor 
         Caption         =   "Size and Color"
      End
      Begin VB.Menu mnuEraser 
         Caption         =   "Toggle Eraser"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gfDrawing As Boolean
Dim pointcounter As Long
Dim Path As String

'Color and size information - sent from frmOptions on user edit.
Public ps As Single
Public pr As Single
Public pg As Single
Public pb As Single

'ToggleEraser stuff.
Dim EraserStep As Integer
Dim lr As Single
Dim lg As Single
Dim lb As Single

'Array stuff for saving.
Dim gx(1 To 10000) As Single
Dim gy(1 To 10000) As Single
Dim gr(1 To 10000) As Single
Dim gg(1 To 10000) As Single
Dim gb(1 To 10000) As Single
Dim gs(1 To 10000) As Single

Private Sub Form_Load()
ClearAll
frmPaint.BackColor = vbWhite
End Sub

Sub OpenFile()
Dim i As Integer
Dim done As Boolean
Dim answer As String
Dim filename As String

done = False
Do While done = False
    cndOpen.ShowOpen
    Path = cndOpen.filename
    cndOpen.filename = ""
    
    If Path = "" Then
        Exit Do
    Else
        answer = MsgBox(Path, vbYesNo, "Open this file?")
        If answer = vbYes Then
            Open Path For Binary Access Read As #1
                Get #1, , pointcounter
                    For i = 1 To pointcounter
                        Get #1, , gx(i)
                        Get #1, , gy(i)
                        Get #1, , gr(i)
                        Get #1, , gg(i)
                        Get #1, , gb(i)
                        Get #1, , gs(i)
                    Next i
                Close #1
            
            frmPaint.Cls
            DrawLines
            done = True
        End If
    End If
Loop
End Sub

Sub SaveAsFile()
Dim i As Integer
Dim done As Boolean
Dim filename As String
Dim answer As String

done = False
Do While done = False
    cndSaveAs.ShowSave
    Path = cndSaveAs.filename
    cndSaveAs.filename = ""
    If Path = "" Then
        Exit Do
    Else
        answer = MsgBox(Path, vbYesNo, "Save to this location?")
        If answer = vbYes Then
            Open Path For Binary Access Write As #1
                Put #1, , pointcounter
                For i = 1 To pointcounter
                    Put #1, , gx(i)
                    Put #1, , gy(i)
                    Put #1, , gr(i)
                    Put #1, , gg(i)
                    Put #1, , gb(i)
                    Put #1, , gs(i)
                Next i
            Close #1
        done = True
        End If
    End If
Loop
End Sub

Sub SaveFile()
Dim i As Integer

If Path = "" Then
    SaveAsFile
Else
    Open Path For Binary Access Write As #1
        Put #1, , pointcounter
        For i = 1 To pointcounter
            Put #1, , gx(i)
            Put #1, , gy(i)
            Put #1, , gr(i)
            Put #1, , gg(i)
            Put #1, , gb(i)
            Put #1, , gs(i)
        Next i
    Close #1
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
gfDrawing = True
DrawCircle X, Y, Val(ps), Val(pr), Val(pg), Val(pb)

CurrentX = X
CurrentY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gfDrawing = True Then
    Line -(X, Y), RGB(pr, pg, pb)
    DrawCircle X, Y, Val(ps), Val(pr), Val(pg), Val(pb)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
gfDrawing = False

If pointcounter < 10000 Then
    pointcounter = pointcounter + 1
    gx(pointcounter) = -1
    gy(pointcounter) = -1
    gr(pointcounter) = -1
    gg(pointcounter) = -1
    gb(pointcounter) = -1
    gs(pointcounter) = -1
End If
End Sub

Private Sub mnuAbout_Click()
MsgBox ("Created by Jack Margeson. Computer Programming 1 Credit Flex, Summer 2018.")
End Sub

Private Sub mnuClear_Click()
Dim answer As Integer
answer = MsgBox("Are you sure you want to clear?", vbYesNo, "Clear")
If answer = vbYes Then
    ClearAll
End If
End Sub

Private Sub mnuEraser_Click()
ToggleEraser
End Sub

Private Sub mnuExit_Click()
Dim answer As Integer
answer = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")
If answer = vbYes Then
    End
End If
End Sub

Sub ClearAll()
frmPaint.Cls
gfDrawing = False
frmPaint.DrawWidth = 5
pointcounter = 0

pr = 0
pg = 0
pb = 0
ps = 5

frmOptions.ps = ps
frmOptions.pr = pr
frmOptions.pg = pg
frmOptions.pb = pb

Path = ""
Erase gx
Erase gy
Erase gr
Erase gg
Erase gb
Erase gs

EraserStep = 0
End Sub

Sub DrawCircle(X As Single, Y As Single, S As Integer, R As Integer, G As Integer, B As Integer)
Circle (X, Y), S, RGB(R, G, B)

If pointcounter < 10000 Then
    pointcounter = pointcounter + 1
    
    gx(pointcounter) = X
    gy(pointcounter) = Y
    gr(pointcounter) = R
    gg(pointcounter) = G
    gb(pointcounter) = B
    gs(pointcounter) = S
End If
End Sub

Sub DrawLines()
Dim i As Integer
Dim step As Integer

Circle (gx(1), gy(1)), 1, RGB(gr(1), gg(1), gb(1))

For i = 2 To pointcounter
    If gx(i) = -1 Or gy(i) = -1 Or gr(i) = -1 Or gg(i) = -1 Or gb(i) = -1 Or gs(i) = -1 Then
        step = 1
        Circle (0, 0), 1, RGB(255, 255, 255)
    ElseIf step = 1 Then
        step = 2
        Circle (0, 0), 1, RGB(255, 255, 255)
    Else
        If step = 2 Then
            frmPaint.DrawWidth = gs(i)
            Circle (gx(i), gy(i)), 1, RGB(gr(i), gg(i), gb(i))
            step = 0
        Else
            frmPaint.DrawWidth = gs(i)
            Line -(gx(i), gy(i)), RGB(gr(i), gg(i), gb(i))
            Circle (gx(i), gy(i)), 1, RGB(gr(i), gg(i), gb(i))
        End If
    End If
Next i

frmPaint.DrawWidth = ps
End Sub

Sub ToggleEraser()
If EraserStep = 0 Then
    EraserStep = 1
    lr = pr
    lg = pg
    lb = pb
    
    pr = 255
    pg = 255
    pb = 255
ElseIf EraserStep = 1 Then
    EraserStep = 0
'This was done to preserve color rather than set it back to black.
    pr = lr
    pg = lg
    pb = lb
End If
End Sub

Private Sub mnuNew_Click()
Dim answer As Integer
answer = MsgBox("Are you sure you want create a new canvas? All unsaved work will be lost.", vbYesNo, "New")
If answer = vbYes Then
    ClearAll
End If
End Sub

Private Sub mnuOpen_Click()
OpenFile
End Sub

Private Sub mnuSave_Click()
Dim answer As Integer

If Path = "" Then
    SaveFile
Else
    answer = MsgBox("Are you sure you want to save? This will overwrite your previous drawing.", vbYesNo, "New")
    If answer = vbYes Then
        SaveFile
    End If
End If
End Sub

Private Sub mnuSaveAs_Click()
SaveAsFile
End Sub

Private Sub mnuSizeandColor_Click()
If EraserStep = 0 Then
    frmOptions.Show
ElseIf EraserStep = 1 Then
    MsgBox ("Please toggle the eraser before changing colors and size!")
End If
End Sub
