Attribute VB_Name = "Module3"
Public Function MakeMove()
'   ElseIf .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
'          .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
'    .Box() = .OPic
   
 Dim randombox As Integer
 If Not CheckForWin Then
  Sleep 200
  With frmMain
   ' Winning Moves
   '  Check for Instant win
   If .Box(0) = .OPic And .Box(2) = .XPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If

   If .Box(0) = .OPic And .Box(2) = .NullPic And .Box(6) = .XPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .XPic And .Box(2) = .OPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .OPic And .Box(6) = .NullPic And .Box(8) = .XPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .XPic And .Box(2) = .NullPic And .Box(6) = .OPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .OPic And .Box(8) = .XPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .XPic And .Box(6) = .NullPic And .Box(8) = .OPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .XPic And .Box(8) = .OPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(0) = .OPic
    GoTo done
   End If
 
   '  Check for instant loss
   If .Box(0) = .XPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(5) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .XPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(3) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .XPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(5) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .XPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(3) = .OPic
    GoTo done
   End If
   
   ' Check for Triagle set up
   If .Box(0) = .OPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .XPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .OPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .XPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .OPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .XPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .OPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .XPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .OPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .XPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .OPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .XPic And .Box(4) = .NullPic And .Box(5) = .NullPic And .Box(7) = .NullPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .XPic And .Box(7) = .OPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .OPic And .Box(2) = .NullPic And .Box(6) = .NullPic And .Box(8) = .NullPic And _
      .Box(1) = .NullPic And .Box(3) = .NullPic And .Box(4) = .NullPic And .Box(5) = .XPic And .Box(7) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   
   '  Top Row
   If .Box(0) = .OPic And .Box(1) = .OPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .OPic And .Box(1) = .NullPic And .Box(2) = .OPic Then
    .Box(1) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(1) = .OPic And .Box(2) = .OPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   '  Middle Row
   If .Box(3) = .OPic And .Box(4) = .OPic And .Box(5) = .NullPic Then
    .Box(5) = .OPic
    GoTo done
   End If
   
   If .Box(3) = .OPic And .Box(4) = .NullPic And .Box(5) = .OPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(3) = .NullPic And .Box(4) = .OPic And .Box(5) = .OPic Then
    .Box(3) = .OPic
    GoTo done
   End If
   
   '  Bottom Row
   If .Box(6) = .OPic And .Box(7) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(6) = .OPic And .Box(7) = .NullPic And .Box(8) = .OPic Then
    .Box(7) = .OPic
    GoTo done
   End If
   
   If .Box(6) = .NullPic And .Box(7) = .OPic And .Box(8) = .OPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   '  Left Column
   If .Box(0) = .OPic And .Box(3) = .OPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .OPic And .Box(3) = .NullPic And .Box(6) = .OPic Then
    .Box(3) = .OPic
    GoTo done
   End If
   If .Box(0) = .NullPic And .Box(3) = .OPic And .Box(6) = .OPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   '  Middle Column
   If .Box(1) = .OPic And .Box(4) = .OPic And .Box(7) = .NullPic Then
    .Box(7) = .OPic
    GoTo done
   End If
   
   If .Box(1) = .OPic And .Box(4) = .NullPic And .Box(7) = .OPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(1) = .NullPic And .Box(4) = .OPic And .Box(7) = .OPic Then
    .Box(1) = .OPic
    GoTo done
   End If
   
   '  Right Column
   If .Box(2) = .OPic And .Box(5) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .OPic And .Box(5) = .NullPic And .Box(8) = .OPic Then
    .Box(5) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .NullPic And .Box(5) = .OPic And .Box(8) = .OPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   '  Right Diagonal
   If .Box(0) = .OPic And .Box(4) = .OPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .OPic And .Box(4) = .NullPic And .Box(8) = .OPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(4) = .OPic And .Box(8) = .OPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   '  Left Diagonal
   If .Box(2) = .OPic And .Box(4) = .OPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .OPic And .Box(4) = .NullPic And .Box(6) = .OPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .NullPic And .Box(4) = .OPic And .Box(6) = .OPic Then
    .Box(2) = .OPic
    GoTo done
   End If
      
   ' Preventive Winning Moves
   '  Top Row
   If .Box(0) = .XPic And .Box(1) = .XPic And .Box(2) = .NullPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .XPic And .Box(1) = .NullPic And .Box(2) = .XPic Then
    .Box(1) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(1) = .XPic And .Box(2) = .XPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   '  Middle Row
   If .Box(3) = .XPic And .Box(4) = .XPic And .Box(5) = .NullPic Then
    .Box(5) = .OPic
    GoTo done
   End If
   
   If .Box(3) = .XPic And .Box(4) = .NullPic And .Box(5) = .XPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(3) = .NullPic And .Box(4) = .XPic And .Box(5) = .XPic Then
    .Box(3) = .OPic
    GoTo done
   End If
   
   '  Bottom Row
   If .Box(6) = .XPic And .Box(7) = .XPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(6) = .XPic And .Box(7) = .NullPic And .Box(8) = .XPic Then
    .Box(7) = .OPic
    GoTo done
   End If
   
   If .Box(6) = .NullPic And .Box(7) = .XPic And .Box(8) = .XPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   '  Left Column
   If .Box(0) = .XPic And .Box(3) = .XPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .XPic And .Box(3) = .NullPic And .Box(6) = .XPic Then
    .Box(3) = .OPic
    GoTo done
   End If
   If .Box(0) = .NullPic And .Box(3) = .XPic And .Box(6) = .XPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   '  Middle Column
   If .Box(1) = .XPic And .Box(4) = .XPic And .Box(7) = .NullPic Then
    .Box(7) = .OPic
    GoTo done
   End If
   
   If .Box(1) = .XPic And .Box(4) = .NullPic And .Box(7) = .XPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(1) = .NullPic And .Box(4) = .XPic And .Box(7) = .XPic Then
    .Box(1) = .OPic
    GoTo done
   End If
   
   '  Right Column
   If .Box(2) = .XPic And .Box(5) = .XPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .XPic And .Box(5) = .NullPic And .Box(8) = .XPic Then
    .Box(5) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .NullPic And .Box(5) = .XPic And .Box(8) = .XPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   '  Right Diagonal
   If .Box(0) = .XPic And .Box(4) = .XPic And .Box(8) = .NullPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .XPic And .Box(4) = .NullPic And .Box(8) = .XPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(0) = .NullPic And .Box(4) = .XPic And .Box(8) = .XPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   '  Left Diagonal
   If .Box(2) = .XPic And .Box(4) = .XPic And .Box(6) = .NullPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .XPic And .Box(4) = .NullPic And .Box(6) = .XPic Then
    .Box(4) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .NullPic And .Box(4) = .XPic And .Box(6) = .XPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   ' Tricks
   '   Triangle
   If .Box(0) = .NullPic And .Box(4) = .OPic And .Box(6) = .OPic Or _
      .Box(0) = .NullPic And .Box(4) = .OPic And .Box(2) = .OPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .NullPic And .Box(4) = .OPic And .Box(0) = .OPic Or _
      .Box(2) = .NullPic And .Box(4) = .OPic And .Box(8) = .OPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(6) = .NullPic And .Box(4) = .OPic And .Box(0) = .OPic Or _
      .Box(6) = .NullPic And .Box(4) = .OPic And .Box(8) = .OPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(8) = .NullPic And .Box(4) = .OPic And .Box(2) = .OPic Or _
      .Box(8) = .NullPic And .Box(4) = .OPic And .Box(6) = .OPic Then
    .Box(8) = .OPic
    GoTo done
   End If
   
   If .Box(4) = .NullPic And .Box(0) = .OPic And .Box(2) = .OPic Or _
      .Box(4) = .NullPic And .Box(0) = .OPic And .Box(6) = .OPic Or _
      .Box(4) = .NullPic And .Box(2) = .OPic And .Box(8) = .OPic Or _
      .Box(4) = .NullPic And .Box(6) = .OPic And .Box(8) = .OPic Then
    .Box(4) = .OPic
    GoTo done
   End If

   '  Three corners
   If .Box(0) = .NullPic And .Box(2) = .OPic And .Box(6) = .OPic Or _
      .Box(0) = .NullPic And .Box(2) = .OPic And .Box(8) = .OPic Or _
      .Box(0) = .NullPic And .Box(6) = .OPic And .Box(8) = .OPic Then
    .Box(0) = .OPic
    GoTo done
   End If
   
   If .Box(2) = .NullPic And .Box(0) = .OPic And .Box(6) = .OPic Or _
      .Box(2) = .NullPic And .Box(0) = .OPic And .Box(8) = .OPic Or _
      .Box(2) = .NullPic And .Box(6) = .OPic And .Box(8) = .OPic Then
    .Box(2) = .OPic
    GoTo done
   End If
   
   If .Box(6) = .NullPic And .Box(0) = .OPic And .Box(2) = .OPic Or _
      .Box(6) = .NullPic And .Box(0) = .OPic And .Box(8) = .OPic Or _
      .Box(6) = .NullPic And .Box(2) = .OPic And .Box(8) = .OPic Then
    .Box(6) = .OPic
    GoTo done
   End If
   
   If .Box(8) = .NullPic And .Box(0) = .OPic And .Box(2) = .OPic Or _
      .Box(8) = .NullPic And .Box(0) = .OPic And .Box(6) = .OPic Or _
      .Box(8) = .NullPic And .Box(6) = .OPic And .Box(2) = .OPic Then
    .Box(8) = .OPic
    GoTo done
   End If
     
   ' Random Move
   randombox = Int(Rnd(1) * 8)
   While .Box(randombox) <> .NullPic And randombox <= 8
    randombox = randombox + 1
    If randombox = 9 Then: randombox = 0
   Wend
   .Box(randombox) = .OPic
  End With
done:
  DoEvents
  CheckForWin
 End If
End Function




