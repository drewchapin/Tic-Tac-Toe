Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public ClearAfterOK As Boolean
Public PlayerScore As Integer
Public ComputerScore As Integer
Public TiedScore As Integer
Public CompTurn As Boolean

Public Function CheckForWin() As Boolean
 With frmMain
  If .Box(0) = .XPic And .Box(1) = .XPic And .Box(2) = .XPic Or _
     .Box(3) = .XPic And .Box(4) = .XPic And .Box(5) = .XPic Or _
     .Box(6) = .XPic And .Box(7) = .XPic And .Box(8) = .XPic Or _
     .Box(0) = .XPic And .Box(3) = .XPic And .Box(6) = .XPic Or _
     .Box(1) = .XPic And .Box(4) = .XPic And .Box(7) = .XPic Or _
     .Box(2) = .XPic And .Box(5) = .XPic And .Box(8) = .XPic Or _
     .Box(0) = .XPic And .Box(4) = .XPic And .Box(8) = .XPic Or _
     .Box(2) = .XPic And .Box(4) = .XPic And .Box(6) = .XPic Then
   XWins
   CheckForWin = True
  ElseIf .Box(0) = .OPic And .Box(1) = .OPic And .Box(2) = .OPic Or _
     .Box(3) = .OPic And .Box(4) = .OPic And .Box(5) = .OPic Or _
     .Box(6) = .OPic And .Box(7) = .OPic And .Box(8) = .OPic Or _
     .Box(0) = .OPic And .Box(3) = .OPic And .Box(6) = .OPic Or _
     .Box(1) = .OPic And .Box(4) = .OPic And .Box(7) = .OPic Or _
     .Box(2) = .OPic And .Box(5) = .OPic And .Box(8) = .OPic Or _
     .Box(0) = .OPic And .Box(4) = .OPic And .Box(8) = .OPic Or _
     .Box(2) = .OPic And .Box(4) = .OPic And .Box(6) = .OPic Then
   OWins
   CheckForWin = True
  ElseIf .Box(0) <> .NullPic And .Box(1) <> .NullPic And .Box(2) <> .NullPic And _
         .Box(3) <> .NullPic And .Box(4) <> .NullPic And .Box(5) <> .NullPic And _
         .Box(6) <> .NullPic And .Box(7) <> .NullPic And .Box(8) <> .NullPic Then
   TiedWin
   CheckForWin = True
  Else
   CheckForWin = False
  End If
 End With
End Function

Public Function XWins()
 frmMsgBox.Show , frmMain
 frmMsgBox.Msg = "You win."
 ClearAfterOK = True
 PlayerScore = PlayerScore + 1
 frmMain.Player = PlayerScore
End Function

Public Function OWins()
 frmMsgBox.Show , frmMain
 frmMsgBox.Msg = "You lost."
 ClearAfterOK = True
 ComputerScore = ComputerScore + 1
 frmMain.Computer = ComputerScore
End Function

Public Function TiedWin()
 frmMsgBox.Show , frmMain
 frmMsgBox.Msg = "It's a tie."
 ClearAfterOK = True
 TiedScore = TiedScore + 1
 frmMain.Ties = TiedScore
End Function

Public Function ClearBoard()
 Dim i As Integer
 i = 0
 With frmMain
  For Each hnd In .Box
   .Box(i) = .NullPic
   i = i + 1
  Next
  If CompTurn Then
   CompTurn = False
   MakeMove
  Else
   CompTurn = True
  End If
 End With
End Function

Public Function MakeMove()
 Dim RandomBox As Integer
 If Not CheckForWin Then
  With frmMain
   Sleep 300
   ' Ofensive Top Row
   If .Box(0) = .OPic And .Box(1) = .OPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(1) = .OPic And .Box(2) = .OPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(0) = .OPic And .Box(2) = .OPic And .Box(1) = .NullPic Then
    .Box(1) = .OPic
   ' Ofensive Middle Row
   ElseIf .Box(3) = .OPic And .Box(4) = .OPic And .Box(5) = .NullPic Then
    .Box(5) = .OPic
   ElseIf .Box(4) = .OPic And .Box(5) = .OPic And .Box(3) = .NullPic Then
    .Box(3) = .OPic
   ElseIf .Box(3) = .OPic And .Box(5) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ' Ofensive Bottom Row
   ElseIf .Box(6) = .OPic And .Box(7) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(7) = .OPic And .Box(8) = .OPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(6) = .OPic And .Box(8) = .OPic And .Box(7) = .NullPic Then
    .Box(7) = .OPic
   ' Ofensive Left Column
   ElseIf .Box(0) = .OPic And .Box(3) = .OPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(3) = .OPic And .Box(6) = .OPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(6) = .OPic And .Box(0) = .OPic And .Box(3) = .NullPic Then
    .Box(3) = .OPic
   ' Ofensive Middle Column
   ElseIf .Box(1) = .OPic And .Box(4) = .OPic And .Box(7) = .NullPic Then
    .Box(7) = .OPic
   ElseIf .Box(4) = .OPic And .Box(7) = .OPic And .Box(1) = .NullPic Then
    .Box(1) = .OPic
   ElseIf .Box(7) = .OPic And .Box(1) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ' Ofensive Right Column
   ElseIf .Box(2) = .OPic And .Box(5) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(5) = .OPic And .Box(8) = .OPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(8) = .OPic And .Box(2) = .OPic And .Box(5) = .NullPic Then
    .Box(5) = .OPic
   ' Ofensive Diagonal Top Left To Bottom Right
   ElseIf .Box(0) = .OPic And .Box(4) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(4) = .OPic And .Box(8) = .OPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(0) = .OPic And .Box(8) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ' Ofensive Diagonal Bottom Left To Top Right
   ElseIf .Box(2) = .OPic And .Box(4) = .OPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(4) = .OPic And .Box(6) = .OPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(2) = .OPic And .Box(6) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
    
    
    
   ' Defensive Top Row
   ElseIf .Box(0) = .XPic And .Box(1) = .XPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(1) = .XPic And .Box(2) = .XPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(0) = .XPic And .Box(2) = .XPic And .Box(1) = .NullPic Then
    .Box(1) = .OPic
   ' Defensive Middle Row
   ElseIf .Box(3) = .XPic And .Box(4) = .XPic And .Box(5) = .NullPic Then
    .Box(5) = .OPic
   ElseIf .Box(4) = .XPic And .Box(5) = .XPic And .Box(3) = .NullPic Then
    .Box(3) = .OPic
   ElseIf .Box(3) = .XPic And .Box(5) = .XPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ' Defensive Bottom Row
   ElseIf .Box(6) = .XPic And .Box(7) = .XPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(7) = .XPic And .Box(8) = .XPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(6) = .XPic And .Box(8) = .XPic And .Box(7) = .NullPic Then
    .Box(7) = .OPic
   ' Defensive Left Column
   ElseIf .Box(0) = .XPic And .Box(3) = .XPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(3) = .XPic And .Box(6) = .XPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(6) = .XPic And .Box(0) = .XPic And .Box(3) = .NullPic Then
    .Box(3) = .OPic
   ' Defensive Middle Column
   ElseIf .Box(1) = .XPic And .Box(4) = .XPic And .Box(7) = .NullPic Then
    .Box(7) = .OPic
   ElseIf .Box(4) = .XPic And .Box(7) = .XPic And .Box(1) = .NullPic Then
    .Box(1) = .OPic
   ElseIf .Box(7) = .XPic And .Box(1) = .XPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ' Defensive Right Column
   ElseIf .Box(2) = .XPic And .Box(5) = .XPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(5) = .XPic And .Box(8) = .XPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(8) = .XPic And .Box(2) = .XPic And .Box(5) = .NullPic Then
    .Box(5) = .OPic
   ' Defensive Diagonal Top Left To Bottom Right
   ElseIf .Box(0) = .XPic And .Box(4) = .XPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(4) = .XPic And .Box(8) = .XPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(0) = .XPic And .Box(8) = .XPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ' Defensive Diagonal Bottom Left To Top Right
   ElseIf .Box(2) = .XPic And .Box(4) = .XPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(4) = .XPic And .Box(6) = .XPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(2) = .XPic And .Box(6) = .XPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
    
   ' Ofesnive Smart Moves
   ElseIf .Box(0) = .OPic And .Box(2) = .OPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(0) = .OPic And .Box(2) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(0) = .OPic And .Box(6) = .OPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(2) = .OPic And .Box(8) = .OPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(2) = .OPic And .Box(0) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(6) = .OPic And .Box(8) = .OPic And .Box(0) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(6) = .OPic And .Box(8) = .OPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
   
   ElseIf .Box(0) = .OPic And .Box(2) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ElseIf .Box(0) = .OPic And .Box(6) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ElseIf .Box(2) = .OPic And .Box(8) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   ElseIf .Box(6) = .OPic And .Box(8) = .OPic And .Box(4) = .NullPic Then
    .Box(4) = .OPic
   
   ElseIf .Box(7) = .OPic And .Box(0) = .OPic And .Box(6) = .NullPic And .Box(3) = .NullPic Then
    .Box(6) = .OPic
   ElseIf .Box(7) = .OPic And .Box(2) = .OPic And .Box(8) = .NullPic And .Box(5) = .NullPic Then
    .Box(8) = .OPic
   ElseIf .Box(1) = .OPic And .Box(6) = .OPic And .Box(0) = .NullPic And .Box(3) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(1) = .OPic And .Box(8) = .OPic And .Box(2) = .NullPic And .Box(5) = .NullPic Then
    .Box(2) = .OPic
   ElseIf .Box(3) = .OPic And .Box(2) = .OPic And .Box(0) = .NullPic And .Box(1) = .NullPic Then
    .Box(0) = .OPic
   ElseIf .Box(3) = .OPic And .Box(7) = .OPic And .Box(6) = .NullPic And .Box(0) = .NullPic Then
    .Box(6) = .OPic
   
   ' Random Move if there are no defensive/ofensive moves
   Else
    RandomBox = Int(Rnd(1) * 8)
    While .Box(RandomBox) <> .NullPic And RandomBox <= 8
     RandomBox = RandomBox + 1
     If RandomBox = 9 Then: RandomBox = 0
    Wend
    .Box(RandomBox) = .OPic
   End If
  End With
  CheckForWin
 End If
End Function

