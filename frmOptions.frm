VERSION 5.00
Begin VB.Form frmOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   1725
   ClientLeft      =   2520
   ClientTop       =   1170
   ClientWidth     =   5670
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ComboO 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox ComboToolbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox ComboX 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "O Style"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   510
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   2880
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Toolbar Style"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   930
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   120
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "X Style"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shborder 
      Height          =   255
      Left            =   2880
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lbButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Apply"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 With ShBorder
  .Left = 0
  .Top = 0
  .Width = Width - 1
  .Height = Height - 1
 End With
 With ComboX
  .AddItem "Red"
  .AddItem "Blue"
  .AddItem "Green"
  .AddItem "Classic"
  Select Case LCase(szX)
   Case "red": .ListIndex = 0
   Case "blue": .ListIndex = 1
   Case "classic": .ListIndex = 3
   Case Else: .ListIndex = 2
  End Select
 End With
 With ComboO
  .AddItem "Red"
  .AddItem "Blue"
  .AddItem "Green"
  .AddItem "Classic"
  Select Case LCase(szO)
   Case "red": .ListIndex = 0
   Case "blue": .ListIndex = 1
   Case "classic": .ListIndex = 3
   Case Else: .ListIndex = 2
  End Select
 End With
 With ComboToolbar
  .AddItem "Red"
  .AddItem "Blue"
  .AddItem "Green"
  .AddItem "White"
  Select Case LCase(szToolbar)
   Case "red": .ListIndex = 0
   Case "blue": .ListIndex = 1
   Case "white": .ListIndex = 3
   Case Else: .ListIndex = 2
  End Select
 End With
End Sub

Private Sub lbButton_Click(Index As Integer)
 Select Case Index
  Case 1: Unload Me
  Case 0
   SaveSetting App.Title, "Bitmaps", "X", ComboX.Text
   SaveSetting App.Title, "Bitmaps", "O", ComboO.Text
   SaveSetting App.Title, "Bitmaps", "Toolbar", ComboToolbar.Text
   GetBitmapSettings
   ClearGrid
   Unload Me
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
