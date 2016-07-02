Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public CompScore, PlayScore, TiedScore, Level As Long
Public CompTurn, AllowTurn As Boolean

Public Function ClearGrid()
 With frmMain
  For i = 0 To 8
   .Box(i) = .NullPic
  Next i
  If CompTurn Then
   CompTurn = False
   .Status = "Player goes first."
  Else
   CompTurn = True
   .Status = "Computer goes first."
   MakeMove
  End If
 End With
 AllowTurn = True
End Function

Public Function ClearScores()
 CompScore = 0
 PlayScore = 0
 TiedScore = 0
 With frmMain
  .Computer = "0"
  .Player = "0"
  .Ties = "0"
 End With
End Function

Private Function DrawLine(pos As Integer)
 With frmMain
  Select Case pos
   Case 0
    .Box(0).Line (0, 810)-(1620, 810)
    .Box(1).Line (0, 810)-(1620, 810)
    .Box(2).Line (0, 810)-(1620, 810)
   Case 1
    .Box(3).Line (0, 810)-(1620, 810)
    .Box(4).Line (0, 810)-(1620, 810)
    .Box(5).Line (0, 810)-(1620, 810)
   Case 2
    .Box(6).Line (0, 810)-(1620, 810)
    .Box(7).Line (0, 810)-(1620, 810)
    .Box(8).Line (0, 810)-(1620, 810)
   Case 3
    .Box(0).Line (810, 0)-(810, 1620)
    .Box(3).Line (810, 0)-(810, 1620)
    .Box(6).Line (810, 0)-(810, 1620)
   Case 4
    .Box(1).Line (810, 0)-(810, 1620)
    .Box(4).Line (810, 0)-(810, 1620)
    .Box(7).Line (810, 0)-(810, 1620)
   Case 5
    .Box(2).Line (810, 0)-(810, 1620)
    .Box(5).Line (810, 0)-(810, 1620)
    .Box(8).Line (810, 0)-(810, 1620)
   Case 6
    .Box(0).Line (0, 0)-(1620, 1620)
    .Box(4).Line (0, 0)-(1620, 1620)
    .Box(8).Line (0, 0)-(1620, 1620)
   Case 7
    .Box(2).Line (0, 1620)-(1620, 0)
    .Box(4).Line (0, 1620)-(1620, 0)
    .Box(6).Line (0, 1620)-(1620, 0)
   Case 8
    .Box(3).Line (200, 200)-(1420, 200)
    .Box(3).Line (810, 200)-(810, 1420)
    .Box(4).Line (200, 200)-(1420, 200)
    .Box(4).Line (810, 200)-(810, 1420)
    .Box(4).Line (200, 1420)-(1420, 1420)
    .Box(5).Line (200, 200)-(200, 1420)
    .Box(5).Line (200, 200)-(1420, 200)
    .Box(5).Line (200, 810)-(810, 810)
    .Box(5).Line (200, 1420)-(1420, 1420)
  End Select
 End With
End Function

Public Function CheckForWin() As Boolean
 With frmMain
  If .Box(0) = .OPic And .Box(1) = .OPic And .Box(2) = .OPic Then
   DrawLine 0
   GoTo owins
  ElseIf .Box(3) = .OPic And .Box(4) = .OPic And .Box(5) = .OPic Then
   DrawLine 1
   GoTo owins
  ElseIf .Box(6) = .OPic And .Box(7) = .OPic And .Box(8) = .OPic Then
   DrawLine 2
   GoTo owins
  ElseIf .Box(0) = .OPic And .Box(3) = .OPic And .Box(6) = .OPic Then
   DrawLine 3
   GoTo owins
  ElseIf .Box(1) = .OPic And .Box(4) = .OPic And .Box(7) = .OPic Then
   DrawLine 4
   GoTo owins
  ElseIf .Box(2) = .OPic And .Box(5) = .OPic And .Box(8) = .OPic Then
   DrawLine 5
   GoTo owins
  ElseIf .Box(0) = .OPic And .Box(4) = .OPic And .Box(8) = .OPic Then
   DrawLine 6
   GoTo owins
  ElseIf .Box(2) = .OPic And .Box(4) = .OPic And .Box(6) = .OPic Then
   DrawLine 7
   GoTo owins
  
  ElseIf .Box(0) = .XPic And .Box(1) = .XPic And .Box(2) = .XPic Then
   DrawLine 0
   GoTo xwins
  ElseIf .Box(3) = .XPic And .Box(4) = .XPic And .Box(5) = .XPic Then
   DrawLine 1
   GoTo xwins
  ElseIf .Box(6) = .XPic And .Box(7) = .XPic And .Box(8) = .XPic Then
   DrawLine 2
   GoTo xwins
  ElseIf .Box(0) = .XPic And .Box(3) = .XPic And .Box(6) = .XPic Then
   DrawLine 3
   GoTo xwins
  ElseIf .Box(1) = .XPic And .Box(4) = .XPic And .Box(7) = .XPic Then
   DrawLine 4
   GoTo xwins
  ElseIf .Box(2) = .XPic And .Box(5) = .XPic And .Box(8) = .XPic Then
   DrawLine 5
   GoTo xwins
  ElseIf .Box(0) = .XPic And .Box(4) = .XPic And .Box(8) = .XPic Then
   DrawLine 6
   GoTo xwins
  ElseIf .Box(2) = .XPic And .Box(4) = .XPic And .Box(6) = .XPic Then
   DrawLine 7
   GoTo xwins
  Else
   For a = 0 To 8
    If .Box(a) = .NullPic Then
     GoTo nowin
    ElseIf a = 8 Then
     DrawLine 8
     GoTo tiedwin
    End If
   Next a
  End If
  Exit Function
xwins:
  CheckForWin = True
  PlayScore = PlayScore + 1
  AllowTurn = False
  .Player = PlayScore
  .Status = "You win."
  Exit Function
owins:
  CheckForWin = True
  CompScore = CompScore + 1
  AllowTurn = False
  .Computer = CompScore
  .Status = "Computer Wins."
  Exit Function
tiedwin:
  CheckForWin = True
  TiedScore = TiedScore + 1
  AllowTurn = False
  .Ties = TiedScore
  .Status = "It's a tie."
  Exit Function
nowin:
  CheckForWin = False
 End With
End Function



