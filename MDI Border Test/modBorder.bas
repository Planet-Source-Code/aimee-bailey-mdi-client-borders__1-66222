Attribute VB_Name = "modBorder"
Option Explicit

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_BORDER = &H800000
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Sub pSetWinStyle(lhWNd As Long, ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)
Dim lS As Long
   lS = GetWindowLong(lhWNd, lType)
   lS = lS And Not lStyleNot
   lS = lS Or lStyle
   SetWindowLong lhWNd, lType, lS
   SetWindowPos lhWNd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Public Sub removeMDIClientBorder(mainForm As MDIForm)
    Dim lhWNd As Long
    
    lhWNd = modBorder.FindWindowEx(mainForm.hWnd, 0&, "MDIClient", vbNullString)
    
    pSetWinStyle lhWNd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
    pSetWinStyle lhWNd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
    
End Sub


