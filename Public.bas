Attribute VB_Name = "Public"
Option Explicit

Public Const RGN_DGR As Long = 2
Public Const TITLE As String = " - 郑智化歌迷联盟特别版"

Public Const SW_SHOWNORMAL = 1
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

Public Const WM_SYSCOMMAND = &H112
Public Const WM_HOTKEY = &H312
Public Const WM_USER = &H400
Public Const WM_SYSTEM_TRAYICON = WM_USER + 1
Public Const WM_LBUTTONDOWN = &H201

Public Const SC_MOVE = &HF010&
Public Const HTCAPTION = 2

Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Function AbountMe() As MSHTML.HTMLDivElement
    Dim t As MSHTML.HTMLDivElement

    Set AbountMe = t
End Function

Public Sub DrawBk(ByVal Frm As Form)
On Error GoTo cw
    'Dim lpRect As RECT
    Dim hRgn As Long
    'Frm.ScaleMode = 3
    Frm.AutoRedraw = True
    '画标题
'    With lpRect
'        .Left = (Frm.Width \ 15 - 95) \ 2
'        .Top = 0
'        .Bottom = 25
'        .Right = .Left + 95
'    End With
'    DrawIconEx Frm.hdc, 5, 2, Frm.Icon.Handle, 20, 20, 0, 0, &H1 Or &H2
'
'    Frm.ForeColor = RGB(230, 230, 230)
'    DrawText Frm.hdc, Frm.Caption, -1, lpRect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
'    Frm.ForeColor = vbBlack
'    lpRect.Left = lpRect.Left + 1
'    lpRect.Top = lpRect.Top + 1
'    DrawText Frm.hdc, Frm.Caption, -1, lpRect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
    '边框
    hRgn = CreateRoundRectRgn(0, 0, Frm.ScaleWidth, Frm.ScaleHeight, RGN_DGR, RGN_DGR)
    FrameRgn Frm.hdc, hRgn, CreateSolidBrush(vbBlack), 1, 1
    
'    hRgn = CreateRoundRectRgn(2, 2, Frm.ScaleWidth - 2, Frm.ScaleHeight - 2, RGN_DGR,RGN_DGR)
'    FrameRgn Frm.hdc, hRgn, CreateSolidBrush(RGB(50, 50, 50)), 1, 1
    DeleteObject hRgn

    'Frm.ScaleMode = 1
Exit Sub

cw:
    MsgBox "出问题了，找小马哥去。", 0 + 48, "疼迅"
    Exit Sub
End Sub
