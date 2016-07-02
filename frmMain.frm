VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Toolbar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   10
      ScaleHeight     =   390
      ScaleWidth      =   7290
      TabIndex        =   4
      Top             =   10
      Width           =   7290
   End
   Begin VB.PictureBox NullPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   2160
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox OPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   2880
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox XPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   1080
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   7
      Left            =   1920
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   11
      Top             =   3960
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   6
      Left            =   240
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   10
      Top             =   3960
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   5
      Left            =   3600
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   9
      Top             =   2280
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   4
      Left            =   1920
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   8
      Top             =   2280
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   3
      Left            =   240
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   7
      Top             =   2280
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   2
      Left            =   3600
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   6
      Top             =   600
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   1
      Left            =   1920
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   5
      Top             =   600
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   0
      Left            =   240
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   3
      Top             =   600
      Width           =   1620
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   1620
      Index           =   8
      Left            =   3600
      ScaleHeight     =   1590
      ScaleWidth      =   1590
      TabIndex        =   12
      Top             =   3960
      Width           =   1620
   End
   Begin VB.Shape ShBorder 
      Height          =   615
      Left            =   5520
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Clear Scores"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   23
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Line Line1 
      X1              =   5460
      X2              =   7200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ties:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Player:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Computer:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lbButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   19
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Status 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5520
      TabIndex        =   18
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lbButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Exit"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   17
      Top             =   2520
      Width           =   1635
   End
   Begin VB.Label lbButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&New Game"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   16
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Ties 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   15
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Player 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Computer 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.Shape Shape2 
      Height          =   3540
      Left            =   5460
      Top             =   480
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      Height          =   5220
      Left            =   120
      Top             =   480
      Width           =   5220
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Box_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  If AllowTurn Then
   Select Case Box(Index).Picture
    Case XPic
     Status = "Space is already occupied by an X"
     DoEvents
    Case OPic
     Status = "Space is already occupied by an O"
     DoEvents
    Case NullPic
     Box(Index) = XPic
     Status = ""
     DoEvents
     MakeMove
   End Select
  Else
   Status = "The game is over. You must start a new game."
  End If
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveSetting App.Title, "Form", "Left", frmMain.Left
 SaveSetting App.Title, "Form", "Top", frmMain.Top
End Sub

Private Sub lbButton_Click(Index As Integer)
 Select Case Index
  Case 0: ClearGrid
  Case 1: Unload Me
  Case 2: frmOptions.Show 1, Me
  Case 3: ClearScores
 End Select
End Sub

Private Sub lbButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbButton(Index).BackColor = RGB(192, 192, 192)
 lbButton(Index).ForeColor = RGB(0, 0, 0)
End Sub

Private Sub lbButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbButton(Index).BackColor = RGB(64, 64, 64)
 lbButton(Index).ForeColor = RGB(255, 255, 255)
 ReleaseCapture
End Sub

Private Sub Form_Load()
 GetBitmapSettings
 With Shborder
  .Left = 0
  .Top = 0
  .Width = Width
  .Height = Height
 End With
 Computer = "0"
 Player = "0"
 Ties = "0"
 Status = ""
 If Second(Time) Mod 2 = 0 Then
  CompTurn = False
 Else
  CompTurn = True
 End If
 ClearGrid
 frmMain.Left = GetSetting(App.Title, "Form", "Left", frmMain.Left)
 frmMain.Top = GetSetting(App.Title, "Form", "Top", frmMain.Top)
End Sub

Private Sub Toolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 1
 End If
End Sub
