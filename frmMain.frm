VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic-Tac-Toe v3"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vsbLevel 
      Height          =   1695
      Left            =   4800
      Max             =   0
      Min             =   2
      TabIndex        =   2
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.PictureBox TitleBar 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   2400
      Y1              =   4560
      Y2              =   840
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1320
      Y1              =   4560
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   3480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   9
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   8
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   7
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   6
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   5
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   4
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   3
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   840
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   2
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   840
      Width           =   885
   End
   Begin VB.Image Box 
      Height          =   1035
      Index           =   1
      Left            =   360
      Stretch         =   -1  'True
      Top             =   840
      Width           =   885
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BoxInfo(1 To 9) As Long
Private Player As Integer
Private Computer As Integer
Private Level As Integer

Private Sub UpdateBox(ByVal Index As Integer)
 If BoxInfo(Index) = 1 Then
  Box(Index).Picture = LoadPicture("X.BMP")
 ElseIf BoxInfo(Index) = 2 Then
  Box(Index).Picture = LoadPicture("O.BMP")
 Else
  Box(Index).Picture = LoadPicture("")
 End If
End Sub

Private Sub Box_Click(Index As Integer)
 If BoxInfo(Index) = 0 Then
  BoxInfo(Index) = Player
  Call UpdateBox(Index)
  Call MakeMove
 Else
  MsgBox "This space is already occupied"
 End If
End Sub

Private Sub Command1_Click()
 Dim i As Integer
 For i = 1 To 9
  BoxInfo(i) = 0
  Call UpdateBox(i)
 Next i
End Sub

Private Sub Form_Load()
 Dim i As Integer
 For i = 1 To 9
  BoxInfo(i) = 0
 Call UpdateBox(i)
 Next i
 Player = 1
 Computer = 2
 Level = 0
 vsbLevel = Level
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 If Button = vbLeftButton Then
  Call ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
 End If
End Sub

Private Sub vsbLevel_Change()
 Level = vsbLevel.Value
 Caption = Level
End Sub

Private Function CheckForWin() As Boolean
 
 ' Horizontal
 ' 123
 ' 456
 ' 789
 ' Vertical
 ' 147
 ' 258
 ' 369
 ' Diagonal
 ' 159
 ' 753
 
 If BoxInfo(1) = Player And BoxInfo(2) = Player And BoxInfo(3) = Player Or _
    BoxInfo(4) = Player And BoxInfo(5) = Player And BoxInfo(6) = Player Or _
    BoxInfo(7) = Player And BoxInfo(8) = Player And BoxInfo(9) = Player Or _
    BoxInfo(1) = Player And BoxInfo(4) = Player And BoxInfo(7) = Player Or _
    BoxInfo(2) = Player And BoxInfo(5) = Player And BoxInfo(8) = Player Or _
    BoxInfo(3) = Player And BoxInfo(6) = Player And BoxInfo(9) = Player Or _
    BoxInfo(1) = Player And BoxInfo(5) = Player And BoxInfo(9) = Player Or _
    BoxInfo(3) = Player And BoxInfo(5) = Player And BoxInfo(7) = Player Then
  
  CheckForWin = True
  MsgBox "you win"
 
 ElseIf BoxInfo(1) = Computer And BoxInfo(2) = Computer And BoxInfo(3) = Computer Or _
        BoxInfo(4) = Computer And BoxInfo(5) = Computer And BoxInfo(6) = Computer Or _
        BoxInfo(7) = Computer And BoxInfo(8) = Computer And BoxInfo(9) = Computer Or _
        BoxInfo(1) = Computer And BoxInfo(4) = Computer And BoxInfo(7) = Computer Or _
        BoxInfo(2) = Computer And BoxInfo(5) = Computer And BoxInfo(8) = Computer Or _
        BoxInfo(3) = Computer And BoxInfo(6) = Computer And BoxInfo(9) = Computer Or _
        BoxInfo(1) = Computer And BoxInfo(5) = Computer And BoxInfo(9) = Computer Or _
        BoxInfo(3) = Computer And BoxInfo(5) = Computer And BoxInfo(7) = Computer Then

  CheckForWin = True
  MsgBox "you lost"
  
 Else
 
  CheckForWin = False
 
 End If
 
End Function

Private Function RandomMove()
 Dim Move As Integer
 Dim Generated As Integer
 
 Do
  Generated = Generated + 1
  If Generated < 9 Then
   Randomize
   Move = Rnd * 7 + 1
  ElseIf Generated = 9 Then
   Move = 1
  ElseIf Generated > 9 Then
   Move = Move + 1
  End If
 Loop Until Box(Move) = 0
 
 RandomMove = Move
End Function

Private Sub MakeMove()
 On Error GoTo errhandler
 Dim Move As Integer
 Dim Generated As Integer
 
 If Not CheckForWin Then
    
  If Level < 1 Then
   Move = RandomMove()
  ' 2,3 are occupied and 1 is open
  ElseIf Level >= 1 And BoxInfo(2) = Computer And BoxInfo(3) = Computer And BoxInfo(1) = 0 Then
   Move = 1
  ' 1,3 are occupied and 2 is open
  ElseIf Level >= 1 And BoxInfo(1) = Computer And BoxInfo(3) = Computer And BoxInfo(2) = 0 Then
   Move = 2
  ' 1,2 are occupied and 3 is open
  ElseIf Level >= 1 And BoxInfo(1) = Computer And BoxInfo(2) = Computer And BoxInfo(3) = 0 Then
   Move = 3
  ' 5,6 are occupied and 4 is open
  ElseIf Level >= 1 And BoxInfo(5) = Computer And BoxInfo(6) = Computer And BoxInfo(4) = 0 Then
   Move = 4
  ' 4,6 are occupied and 5 is open
  ElseIf Level >= 1 And BoxInfo(4) = Computer And BoxInfo(6) = Computer And BoxInfo(5) = 0 Then
   Move = 5
  ' 4,5 are occupied and 6 is open
  ElseIf Level >= 1 And BoxInfo(4) = Computer And BoxInfo(5) = Computer And BoxInfo(6) = 0 Then
   Move = 6
  ' 8,9 are occupied and 7 is open
  ElseIf Level >= 1 And BoxInfo(1) = Computer And BoxInfo(2) = Computer And BoxInfo(3) = 0 Then
   Move = 7
  ' 7,9 are occupied and 8 is open
  ElseIf Level >= 1 And BoxInfo(7) = Computer And BoxInfo(9) = Computer And BoxInfo(8) = 0 Then
   Move = 8
  ' 7,8 are occupied and 9 is open
  ElseIf Level >= 1 And BoxInfo(7) = Computer And BoxInfo(8) = Computer And BoxInfo(9) = 0 Then
   Move = 9
  ' 1,5 are occupied and 9 is open
  ElseIf Level >= 1 And BoxInfo(1) = Computer And BoxInfo(5) = Computer And BoxInfo(9) = 0 Then
   Move = 9
  ' 1,9 are occupied and 5 is open
  ElseIf Level >= 1 And BoxInfo(1) = Computer And BoxInfo(9) = Computer And BoxInfo(5) = 0 Then
   Move = 5
   ' 5,9 are occupied and 1 is open
  ElseIf Level >= 1 And BoxInfo(5) = Computer And BoxInfo(9) = Computer And BoxInfo(1) = 0 Then
   Move = 1
  ' 3,5 are occupied and 7 is open
  ElseIf Level >= 1 And BoxInfo(3) = Computer And BoxInfo(5) = Computer And BoxInfo(7) = 0 Then
   Move = 7
  ' 5,7 are occupied and 3 is open
  ElseIf Level >= 1 And BoxInfo(5) = Computer And BoxInfo(7) = Computer And BoxInfo(3) = 0 Then
   Move = 3
   ' 3,7 are occupied and 5 is open
  ElseIf Level >= 1 And BoxInfo(3) = Computer And BoxInfo(7) = Computer And BoxInfo(5) = 0 Then
   Move = 5
  Else
   Move = RandomMove
  End If
  
  BoxInfo(Move) = Computer
  Call UpdateBox(Move)
  Call CheckForWin
 
 End If
Exit Sub
errhandler:
   MsgBox Move

End Sub

