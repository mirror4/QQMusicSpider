VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "About Me"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":014A
   ScaleHeight     =   138
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   254
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "["
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1740
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   3795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "- AKI Studio 2011 -"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1740
      Width           =   1815
   End
   Begin VB.Label id 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrior-Pro"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   1740
      Width           =   1155
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Call CloseMe
End Sub

Private Sub CloseMe()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CloseMe
        Case vbKeyEscape
            Call CloseMe
        Case vbKeySpace
            Call CloseMe
    End Select
End Sub

Private Sub Form_Load()
    Dim lpRect As RECT, hRgn As Long
    'me.ScaleMode = 3
    Me.AutoRedraw = True
    'ª≠±ÍÃ‚
    With lpRect
        .Left = (Me.Width \ 15 - 95) \ 2
        .Top = 0
        .Bottom = 25
        .Right = .Left + 95
    End With

    hRgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, 24, RGN_DGR, RGN_DGR)
    FrameRgn Me.hdc, hRgn, CreateSolidBrush(RGB(67, 117, 250)), 1, 12

    DrawIconEx Me.hdc, 5, 2, Me.Icon.Handle, 20, 20, 0, 0, &H1 Or &H2
    Me.FontBold = True
    Me.ForeColor = vbBlack
    DrawText Me.hdc, App.TITLE, -1, lpRect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
    Me.ForeColor = vbWhite
    lpRect.Left = lpRect.Left - 2
    lpRect.Top = lpRect.Top - 3
    DrawText Me.hdc, App.TITLE, -1, lpRect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE

    hRgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, RGN_DGR, RGN_DGR)
    SetWindowRgn Me.hWnd, hRgn, True

    DeleteObject hRgn
    'Me.ScaleMode = 1
    DrawBk Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Me.MousePointer = vbDefault
    With id
        .ForeColor = &H80000002
        '.FontBold = False
    End With
End Sub

'Private Sub id_Click()
'    ShellExecute Me.hWnd, vbNullString, "http://t.qq.com/warrior-pro", vbNullString, "", SW_SHOWNORMAL
'End Sub

Private Sub id_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With id
        .ForeColor = vbRed
        SetCursor 65581
        '.FontBold = True
        '.MousePointer = vbhand
    End With
End Sub

Private Sub id_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ShellExecute Me.hWnd, vbNullString, "http://t.qq.com/warrior-pro", vbNullString, "", SW_SHOWNORMAL
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MOVE Or HTCAPTION, ByVal 0&
    End If
End Sub

