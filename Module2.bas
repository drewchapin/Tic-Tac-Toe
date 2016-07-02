Attribute VB_Name = "Module2"
Public szToolbar As String
Public szX As String
Public szO As String

Public Function GetBitmapSettings()
 szToolbar = GetSetting(App.Title, "Bitmaps", "Toolbar", "Green")
 szX = GetSetting(App.Title, "Bitmaps", "X", "Green")
 szO = GetSetting(App.Title, "Bitmaps", "O", "Green")
 Select Case LCase(szToolbar)
  Case "red": frmMain.Toolbar = LoadResPicture("REDTOOLBAR", 0)
  Case "blue": frmMain.Toolbar = LoadResPicture("BLUETOOLBAR", 0)
  Case "white": frmMain.Toolbar = LoadResPicture("WHITETOOLBAR", 0)
  Case Else: frmMain.Toolbar = LoadResPicture("GREENTOOLBAR", 0)
 End Select
 Select Case LCase(szX)
  Case "red": frmMain.XPic = LoadResPicture("REDX", 0)
  Case "blue": frmMain.XPic = LoadResPicture("BLUEX", 0)
  Case "classic": frmMain.XPic = LoadResPicture("CLASSICX", 0)
  Case Else: frmMain.XPic = LoadResPicture("GREENX", 0)
 End Select
 Select Case LCase(szO)
  Case "red": frmMain.OPic = LoadResPicture("REDO", 0)
  Case "blue": frmMain.OPic = LoadResPicture("BLUEO", 0)
  Case "classic": frmMain.OPic = LoadResPicture("CLASSICO", 0)
  Case Else: frmMain.OPic = LoadResPicture("GREENO", 0)
 End Select
End Function

