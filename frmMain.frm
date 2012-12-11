VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "就让爱，把我点亮 - QQMusicSpider"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":3BFA
   ScaleHeight     =   653
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6960
      Top             =   1020
   End
   Begin SHDocVwCtl.WebBrowser soso 
      Height          =   9795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      ExtentX         =   15875
      ExtentY         =   17277
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "mnuSysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SOSO_MUSIC_INDEX$ = "http://music.qq.com/midportal/frame/index.html"
Private Const SOSO_MUSIC_SEARCH$ = "http://music.qq.com/midportal/static/search/index.html"
Private Const SOSO_MUSIC_COOKIE$ = "http://soso.music.qq.com/fcgi-bin/music2.fcg"

Private Const SOSO_SPIDER_ABOUT$ = "<iframe src=\'http://follow.v.t.qq.com/index.php?c=follow&a=quick&name=yijijin&style=1&t=1351878783588&f=1\' frameborder=\'0\' scrolling=\'auto\' width=\'227\' height=\'75\' marginwidth=\'0\' marginheight=\'0\' allowtransparency=\'true\'></iframe><br>" _
        & "<iframe src=\'http://follow.v.t.qq.com/index.php?c=follow&a=quick&name=warrior-pro&style=1&t=1351878783588&f=1\' frameborder=\'0\' scrolling=\'auto\' width=\'227\' height=\'75\' marginwidth=\'0\' marginheight=\'0\' allowtransparency=\'true\'></iframe>"

Private Sub Form_Load()
    With soso
        .Height = 0
        .Width = 0
    End With

    Me.AutoRedraw = True
    Print "Loading..."
    Timer1.Enabled = True
    'Timer1_Timer
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Select Case Me.WindowState
        Case vbMinimized
            TrayAddIcon Me, mnuSysTray, App.TITLE & vbCrLf _
                & "美利达(@warrior-pro)" & vbCrLf _
                & "http://t.qq.com/warrior-pro"   '这段的作用是在任务蓝里新建一个图标
            Me.Hide
            'TrayBalloon Me, "气泡内容", "气泡标题", NIIF_INFO
        'Case vbMaximized
        Case vbNormal
            TrayRemoveIcon  '程序关闭时触发，删除任务栏图标
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TrayRemoveIcon  '程序关闭时触发，删除任务栏图标
End Sub

Private Sub About()
    Randomize
    If (Fix(Rnd() * 100) Mod 2) Then
        frmAbout.Show 1
    Else
        Call AboutMe
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuShow_Click()
    If Not Me.Visible Then
        Me.WindowState = vbNormal
        Me.Show
    End If
End Sub

'Private Sub soso_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
'    MsgBox Command
'End Sub

Private Sub soso_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    HookDocument pDisp.Document
End Sub

Private Sub HookDocument(ByVal Document As HTMLDocument)
    On Error Resume Next
    With Document
        If InStr(.URL, "res://ieframe.dll") > 0 Then
            Print "Cannot connect to QQMusic server,please check your internet connection."
            Exit Sub
        End If
    
        If InStr(.parentWindow.Top.Document.URL, SOSO_MUSIC_INDEX) > 0 Then
            If InStr(.URL, "music.qq.com") <> 0 And InStr(.URL, "/ad/") = 0 Then SetCookie Document
        End If

        If StrComp(.URL, SOSO_MUSIC_COOKIE, vbTextCompare) = 0 Then
            If InStr(.cookie, "qqmusic_version") = 0 Then
                Timer1.Enabled = False
                SetCookie Document
                soso.navigate SOSO_MUSIC_INDEX
            End If
        ElseIf soso.Height < 10 Then
            With soso
                .Height = Me.ScaleHeight
                .Width = Me.ScaleWidth
            End With
        End If
    End With

    If StrComp(Document.URL, SOSO_MUSIC_SEARCH, vbTextCompare) = 0 Then
        HookSearch Document
    Else
        HookDownload Document
    End If
End Sub

Private Sub SetCookie(ByVal Document As HTMLDocument)
    On Error Resume Next
    With Document
        Me.Caption = .TITLE & TITLE

        If InStr(.cookie, "qqmusic_version") = 0 Then
            .cookie = "qqmusic_uin=10001"
            .cookie = "qqmusic_version=8"
            .cookie = "qqmusic_miniversion=00"
            '.cookie = "qqmusic_fromtag=17"
            .cookie = "detail=10001,1,%u5C4C%u4E1D%u4EEC%uFF0C%u8054%u5408%u8D77%u6765%uFF01,0,0,5,," & Format(DateAdd("d", 1, Now), "yyyy-MM-dd hh:mm:ss") & ",1,,," & Format(Now, "yyyy-MM-dd hh:mm:ss") & ",,0,0,0,0,0"
            '.cookie = "qqmusic_key=45F069FBD5A7D67816F08390612ED4DA2CB6AAED3904B3066A2510B6AE477EA5"
            '.cookie = "qqmusic_privatekey=4D3ED0EEE242A83F5ECF43A1572CF8BF03FE993A15CE8DCF"
        End If
    End With
End Sub

Private Sub HookDownload(ByVal Document As HTMLDocument)
    On Error Resume Next
    Dim JS As HTMLScriptElement
    Dim Script As String

    If InStr(Document.URL, "music.qq.com") <> 0 Then
        Script = "if (parent.g_download){" & vbCrLf _
            & "//download.js#usertype/downtype" & vbCrLf _
            & "if (!this.g_downloadOne.setDownButton_x)" _
                & "this.g_downloadOne.setDownButton_x = this.g_downloadOne.setDownButton;" & vbCrLf _
            & "this.g_downloadOne.setDownButton = " _
                & "function(song_type){this.usertype='vip';this.downtype='vipcard';this.setDownButton_x(song_type);}" & vbCrLf _
            & "//qmfl-core.js#g_download.start()" & vbCrLf _
            & "if (!this.g_download.start_x){" _
                & "this.g_download.start_x = this.g_download.start;" _
            & "}" & vbCrLf _
            & "this.g_download.start = " _
                & "function(objList, isOne, isVipCard){this.start_x(objList,isOne,isVipCard);}" & vbCrLf _
            & "" _
            & "parent.g_player.setQQMusic = function(xml){" & vbCrLf _
            & "var songs = '';var obj;" & vbCrLf _
            & "var re = /<filename><!\[CDATA\[(.*?)\]\]><\/filename>.*?<url><!\[CDATA\[(.*?)\]\]><\/url>/ig;" & vbCrLf _
            & "while ((obj = re.exec(xml)) != null){songs += obj[1] + '\t' + obj[2] + '\r\n'}" & vbCrLf _
   		&"if (!confirm(songs)) return false;this.initQQMusic();parent.QQMusic.WebPerform(xml);};" & vbCrLf _
            & "}" & vbCrLf
    
        Script = Script & AboutScript
    Else
        Script = "this.document.oncontextmenu = function(){return false;}"
    End If

    With Document
        Set JS = .createElement("SCRIPT")
        JS.Text = Script
        .body.appendChild JS
    End With
End Sub

Private Sub HookSearch(ByVal Document As HTMLDocument)
    On Error Resume Next
    Dim JS As HTMLScriptElement
    Dim Script As String
    With Document
'        With .getElementById("rem_pic")
'            With .Style
'                .TextAlign = "center"
'                .display = "block"
'            End With
'
'            '.innerHTML = "<iframe src='http://follow.v.t.qq.com/index.php?c=follow&a=quick&name=warrior-pro&style=1&t=1351878783588&f=1' frameborder='0' scrolling='auto' width='227' height='75' marginwidth='0' marginheight='0' allowtransparency='true'></iframe>"
'            '.innerHTML = "<a href='http://t.qq.com/warrior-pro' target='_blank'><img src='http://v.t.qq.com/sign/warrior-pro/518b57f6b28a7b419acebc6e8b5e0bfaa3db6e08/1.jpg' width='380' height='100' /></a>"
'            .innerHTML = "<iframe src='http://follow.v.t.qq.com/index.php?c=follow&a=quick&name=yijijin&style=1&t=1351878783588&f=1' frameborder='0' scrolling='auto' width='227' height='75' marginwidth='0' marginheight='0' allowtransparency='true'></iframe><br><iframe src='http://follow.v.t.qq.com/index.php?c=follow&a=quick&name=warrior-pro&style=1&t=1351878783588&f=1' frameborder='0' scrolling='auto' width='227' height='75' marginwidth='0' marginheight='0' allowtransparency='true'></iframe>"
'            '.innerHTML = .innerHTML & "<A href='http://www.52pojie.cn/' target=_blank>" _
'                & "<IMG title='吾爱破解' style=' width: 160px; height: 66px;' " _
'                & "src='http://www.52pojie.cn/static/image/common/logo.png'></A>"
'            '.innerHTML = .innerHTML & "<A href='http://www.meizu.com' target=_blank>" _
'                & "<IMG title='世间喧哗不如孤独的美 - 我是囧王 射死你们' " _
'                & "style='border-spacing: 2px;border: 2px solid rgb(255,255,255); width: 120px; height: 90px;' " _
'                & "src='http://user.meizu.com/data/avatar/000/00/00/02_avatar_middle.jpg'></A>"
'            '.innerHTML = "<div id=""txWB_W1""></div>" _
'                & "<script type=""text/javascript"">" _
'                & "var tencent_wb_name = ""warrior-pro"";" _
'                & "var tencent_wb_sign = ""518b57f6b28a7b419acebc6e8b5e0bfaa3db6e08"";" _
'                & "var tencent_wb_style = ""1"";" _
'                & "</script>" _
'                & "<script type=""text/javascript"" src=""http://v.t.qq.com/follow/widget.js"" charset=""utf-8""/>" _
'                & "</script>"
'        End With

                Script = "var gyad=this.document.getElementById('rem_pic');if(gyad){;gyad.innerHTML = '" & SOSO_SPIDER_ABOUT _
                        & "';gyad.style.textAlign='center';gyad.style.display='block';gyad.style.top = '370px';}" & vbCrLf
        Script = Script & "var tomato = setInterval(function(){clearInterval(tomato);with(document.getElementById('w')){" _
                        & "var songs = [" _
                                & "'一块钱','让世界充满爱','What More Can I Give Michael Jackson','一块钱'," _
                                & "'What\'s Going On Marvin Gaye','We Will Get There','Heal The World','一块钱'," _
                                & "'We Are The World','Imagine - John Lennon','Tell Me Why Declan Galbraith','一块钱'," _
                                & "'Amani BEYOND','The Lost Children','承诺 刘德华','血染的风采 :)','一块钱'" _
                                & "/*value='Arrietty\'s Song'//（「壹基金」主题曲）*/];" & vbCrLf _
                        & "value=songs[Math.floor(Math.random()*songs.length)];select();focus()}},10);"

        Set JS = .createElement("SCRIPT")
        JS.Text = Script & AboutScript
        .body.appendChild JS
    End With
End Sub

Private Function AboutScript() As String
    Dim tScript$

    'document.oncontextmenu
    tScript = tScript & "this.document.oncontextmenu = function(){" & vbCrLf _
            & "var dg = window.g_dialog?window.g_dialog:window.parent.g_dialog;" & vbCrLf _
            & "if(dg){dg.show({" _
                & "mode: 'common'," _
                & "title:'" & App.TITLE & " - VER/ " & App.Major & "." & App.Minor & "." & App.Revision & "'," & vbCrLf _
                & "icon_type: 0," _
                & "sub_title: '" & SOSO_SPIDER_ABOUT & "'," & vbCrLf _
                & "//sub_title:'<img style=\'margin-right: 10px;\' src=\'http://t3.qlogo.cn/mbloghead/a2f00b80457357b99658/100\' width=40 height=40>" _
                    & "向杕献礼！'," & vbCrLf _
                & "desc:'<hr><b><font color=\'black\'>AKI STUDIO &copy;2012-2013</font></b><span style=\'margin-left:20px;background:url(http://imgcache.qq.com/mediastyle/minimusic/img/sprite_download.png) no-repeat 0 -145px;padding:0 0 0 70px\'></span><br>" _
					&"<a target=_blank title=\'尽我所能 人人公益\' alt=\'壹基金\' style=\'border: #4A95DF solid 1px;padding:0 0 4px 48px;margin:2px 10px 0 0;background:url(http://www.onefoundation.cn/images/2011/logo.gif) no-repeat -2px -6px;background-size:50px 29px;\' href=\'http://www.onefoundation.cn/html/cn/beneficence.html\'></a>" _
					& "'" & vbCrLf _
            & "})}" & vbCrLf _
        & "return false;}"

        AboutScript = tScript
End Function

Private Sub Timer1_Timer()
	Timer1.Interval = 10000
    soso.navigate SOSO_MUSIC_COOKIE
End Sub
