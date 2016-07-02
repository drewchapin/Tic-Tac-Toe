VERSION 5.00
Begin VB.Form frmMsgBox 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4650
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tic-Tac-Toe"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.Shape frmBorder 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label OKButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00414141&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&OK"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "OK"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   120
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Msg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 frmMain.Enabled = False
  With frmBorder
  .Left = 0
  .Top = 0
  .Width = Width - 1
  .Height = Height - 1
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMain.Enabled = True
End Sub

Private Sub OKButton_Click()
 If ClearAfterOK Then: ClearBoard
 Unload Me
End Sub

Private Sub OKButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 OKButton.BackColor = RGB(128, 128, 128)
End Sub

Private Sub OKButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 OKButton.BackColor = RGB(65, 65, 65)
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
