VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00414141&
   BorderStyle     =   0  'None
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox NullPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   5760
      ScaleHeight     =   1620
      ScaleWidth      =   1665
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   7545
      TabIndex        =   19
      Top             =   0
      Width           =   7575
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tic-Tac-Toe"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.PictureBox OPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   2880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1620
      ScaleWidth      =   1620
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox XPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   1200
      Picture         =   "Form1.frx":5784
      ScaleHeight     =   1620
      ScaleWidth      =   1620
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   8
      Left            =   3840
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   10
      Top             =   4080
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   7
      Left            =   2040
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   9
      Top             =   4080
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   6
      Left            =   240
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   8
      Top             =   4080
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   5
      Left            =   3840
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   7
      Top             =   2280
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   4
      Left            =   2040
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   6
      Top             =   2280
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   3
      Left            =   240
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   5
      Top             =   2280
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   2
      Left            =   3840
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   4
      Top             =   480
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   1
      Left            =   2040
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   3
      Top             =   480
      Width           =   1680
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   0
      Left            =   240
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   2
      Top             =   480
      Width           =   1680
   End
   Begin VB.Shape frmBorder 
      Height          =   1530
      Left            =   5760
      Top             =   2160
      Width           =   1650
   End
   Begin VB.Label ExitButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Exit Game"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "Exit Game"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label NewGame 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&New Game"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      ToolTipText     =   "Clear the bord for a new game"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Ties 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ties:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Computer 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Player 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Computer:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Shape Shape2 
      Height          =   1695
      Left            =   5760
      Top             =   360
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   5520
      Left            =   120
      Top             =   360
      Width           =   5520
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Box_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  Select Case Box(Index).Picture
   Case NullPic
    Box(Index) = XPic
    DoEvents
    MakeMove
   Case XPic
    ClearAfterOK = False
    frmMsgBox.Show , Me
    frmMsgBox.Msg = "This space is already occupied by an X"
   Case OPic
    ClearAfterOK = False
    frmMsgBox.Show , Me
    frmMsgBox.Msg = "This space is already occupied by an O"
  End Select
 End If
End Sub

Private Sub ExitButton_Click()
 Unload Me
End Sub

Private Sub ExitButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ExitButton.BackColor = RGB(128, 128, 128)
End Sub

Private Sub ExitButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ExitButton.BackColor = RGB(65, 65, 65)
 ReleaseCapture
End Sub

Private Sub Form_Load()
 Dim X As Double
 With frmBorder
  .Left = 0
  .Top = 0
  .Width = Width - 1
  .Height = Height - 1
 End With
 If Second(Time) Mod 2 = 0 Then: CompTurn = False
 ClearBoard
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 1
 End If
End Sub

Private Sub NewGame_Click()
 ClearBoard
 Player.Caption = 0
 Computer.Caption = 0
 Ties.Caption = 0
 PlayerScore = 0
 ComputerScore = 0
 TiedScore = 0
End Sub

Private Sub NewGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 NewGame.BackColor = RGB(128, 128, 128)
End Sub

Private Sub NewGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 NewGame.BackColor = RGB(65, 65, 65)
 ReleaseCapture
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 1
 End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 1
 End If
End Sub
