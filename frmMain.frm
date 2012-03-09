VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QQMusicSpider"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser soso 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7635
      ExtentX         =   13467
      ExtentY         =   12515
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SOSO_MUSIC_DEFAULT As String = "http://music.qq.com/miniportal/search.html"
Private Const SOSO_MUSIC_CGI As String = "http://soso.music.qq.com/fcgi-bin/music.fcg"
Private Const SOSO_MUSIC_SONGS As String = "http://imgcache.qq.com/music/miniportal_v3/tips/songs_download.html"
Private Const SOSO_MUSIC_SINGLE As String = "http://imgcache.qq.com/music/miniportal_v3/tips/single_download_soso.html?v=1"
Private Const SOSO_MUSIC_STATIC As String = "http://music.qq.com/miniportal/static"

Private Const SOOS_MUSIC_PARSER As String _
    = "function soso_music_parser(objMusic){var d_url = '';" _
        & "switch (parseInt(objMusic.type_index)){" _
            & "case 0:d_url += 'http://download.music.qq.com/track_flac/' + (70000000 + parseInt(objMusic.mid)) + '.flac';break;" _
            & "case 1:d_url += 'http://download.music.qq.com/track_ape/' + (80000000 + parseInt(objMusic.mid)) + '.ape';break;" _
            & "case 2:d_url += 'http://download.music.qq.com/320k/' + objMusic.mid + '.mp3';break;" _
            & "case 3:var downloadID = parseInt(objMusic.mid) + 30000000;d_url += 'http://stream' + (parseInt(objMusic.mstream) + 10) + '.qqmusic.qq.com/' + downloadID + '.mp3';break;" _
            & "case 4:d_url += 'http://stream' + parseInt(objMusic.mstream) + '.qqmusic.qq.com/' + (parseInt(objMusic.mid) + 12000000) + '.wma';break;" _
            & "default: d_url += objMusic.msongurl;" _
        & "}return '<' + objMusic.mid + '> ' + objMusic.msinger + ' - ' + objMusic.msong + ': ' + d_url + '\r\n'" _
    & "} function _confirm(songlist,callback){alert(songlist);return confirm('Download songs immediately?');}"

Private WithEvents m_Doc As MSHTML.HTMLDocument
Attribute m_Doc.VB_VarHelpID = -1
Private bCookied As Boolean

Private Sub Form_Load()
    'soso.Visible = False
    soso.navigate SOSO_MUSIC_CGI
    'soso.Visible = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
End Sub

Private Function m_Doc_oncontextmenu() As Boolean
    frmAbout.Show 1
End Function

Private Sub About()
    Randomize
    If (Fix(Rnd() * 100) Mod 2) Then
        frmAbout.Show 1
    Else
        Call AbountMe
    End If
End Sub

Private Sub soso_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    
    If Not bCookied Then
        soso.navigate SOSO_MUSIC_DEFAULT
        bCookied = True
    End If
    
    With soso.document
        If StrComp(URL, SOSO_MUSIC_DEFAULT, vbTextCompare) = 0 Then
            With .getElementById("rem_pic")
                With .Style
                    .TextAlign = "center"
                    .display = "block"
                End With
                .innerHTML = "<A href='http://www.zhleague.com/' target=_blank>" _
                    & "<IMG title='zapza1517 @ Ö£ÖÇ»¯¸èÃÔÁªÃË' style='background: url(http://www.sohtanaka.com/web-design/examples/drop-shadow/shadow-1000x1000.gif) no-repeat right bottom;height:40' src='http://www.zhleague.com/image_square/title.jpg'></A>" _
                    & "<A href='http://t.qq.com/warrior-pro' target=_blank>" _
                    & "<IMG title='warrior-pro' style='background: url(http://www.sohtanaka.com/web-design/examples/drop-shadow/shadow-1000x1000.gif) no-repeat right bottom;padding: 5px 10px 10px 5px;height:40' src='http://t3.qlogo.cn/mbloghead/69ffc58e06c333ce3474/100'></A>"
            End With
            With .getElementById("w")
                If Len(.getAttribute("value")) = 0 Then
                    .setAttribute "value", "µ­Ë®ºÓ±ßµÄÑÌ»ð"
                End If
                .Select
                .focus
            End With
        End If
    End With
End Sub

Private Sub soso_NavigateComplete2(ByVal pDisp As Object, tURL As Variant)
    On Error Resume Next
    Dim js As MSHTML.HTMLScriptElement

'    If InStr(1, tURL, SOSO_MUSIC_STATIC, vbTextCompare) > 0 _
'        Or InStr(1, tURL, SOSO_MUSIC_CGI, vbTextCompare) > 0 Then
        Set m_Doc = soso.document
        With m_Doc
            Me.Caption = .TITLE & TITLE
            .cookie = "qqmusic_version=7"
            .cookie = "qqmusic_miniversion=96"
            .cookie = "qqmusic_uin=10001"
        End With
'    End If

    If StrComp(tURL, SOSO_MUSIC_SONGS, vbTextCompare) = 0 Then
        Set m_Doc = m_Doc.frames("frame_tips").document
        With m_Doc '.frames("frame_tips").Document
            Set js = .createElement("SCRIPT")
            js.Text = SOOS_MUSIC_PARSER _
                & "g_songsDownload.download_x = g_songsDownload.download;" _
                & "g_songsDownload.download = function(){" _
                & "var d_url = '';" _
                & "var checklist = document.getElementsByName('typecheck');" _
                & "for (var i = 0; i < checklist.length; i++) {" _
                    & "if (checklist[i].checked) {" _
                        & "var no = checklist[i].parentNode.parentNode.getAttribute('no');" _
                        & "var type_index = checklist[i].parentNode.parentNode.getAttribute('type_index');" _
                        & "var objMusic = g_songsDownload.songlist[no];" _
                        & "objMusic.type_index = type_index;" _
                        & "d_url += soso_music_parser(objMusic);" _
                    & "}" _
                & "}" _
                & "if (_confirm(d_url)) g_songsDownload.download_x();" _
            & "};"
            .body.appendChild js
        End With
    End If

    If StrComp(tURL, SOSO_MUSIC_SINGLE, vbTextCompare) = 0 Then
        Set m_Doc = m_Doc.frames("frame_tips").document
        With m_Doc '.frames("frame_tips").Document
            Set js = .createElement("SCRIPT")
            js.Text = SOOS_MUSIC_PARSER _
                & "g_singleDownload.init_x = g_singleDownload.init;" _
                & "g_singleDownload.init = function(){" _
                & "g_singleDownload.init_x();" _
                & "g_singleDownload.downtype='vip';" _
                & "};" _
                & "g_singleDownload.download_x = g_singleDownload.download;" _
                & "g_singleDownload.download = function(){" _
                & "var type = 0;" _
                & "var songList = g_singleDownload._$('song_list').getElementsByTagName('input');" _
                & "for (var i = 0; i < songList.length; i++){if (songList[i].checked){type = songList[i].value;break;}}" _
                & "switch (parseInt(type)){" _
                & "case 1:g_singleDownload.music.type_index=4;break;" _
                & "case 11:g_singleDownload.music.type_index=4;break;" _
                & "case 2:g_singleDownload.music.type_index=3;break;" _
                & "case 12:g_singleDownload.music.type_index=3;break;" _
                & "case 3:g_singleDownload.music.type_index=2;break;" _
                & "case 4:g_singleDownload.music.type_index=1;break;" _
                & "case 5:g_singleDownload.music.type_index=0;break;" _
                & "default:g_singleDownload.music.type_index=5;break;" _
                & "}" _
                & "if (_confirm(soso_music_parser(g_singleDownload.music))) g_singleDownload.download_x();" _
                & "};"
            .body.appendChild js
        End With
    End If
End Sub
