VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3GP 转换精灵 Beta 1.0 "
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10515
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer_Refresh 
      Interval        =   100
      Left            =   6660
      Top             =   45
   End
   Begin VB.CommandButton cmdAD 
      Caption         =   "个性定制"
      Height          =   900
      Left            =   9405
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "想让这个软件显示你的个性内容吗？快快点我吧！"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdViewBrowser 
      Caption         =   "迷你首页"
      Height          =   900
      Left            =   8370
      Picture         =   "frmMain.frx":0D88
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "显示迷你首页"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "关于软件"
      Height          =   900
      Left            =   7335
      Picture         =   "frmMain.frx":13B9
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "关于软件的简单介绍"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "显示列表"
      Height          =   900
      Left            =   5220
      Picture         =   "frmMain.frx":1684
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "点击此处显示待转换的文件列表"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "退出程序"
      Height          =   900
      Left            =   4185
      Picture         =   "frmMain.frx":337E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "退出程序"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "提示设置"
      Height          =   900
      Left            =   3150
      Picture         =   "frmMain.frx":41C0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "点击此处设置转换完成后的声音提醒功能"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "移除文件"
      Height          =   900
      Left            =   2115
      Picture         =   "frmMain.frx":5EBA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "点击此处从文件列表中移出待转换的文件"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "添加文件"
      Height          =   900
      Left            =   45
      Picture         =   "frmMain.frx":6CFC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "点击此处选择待转换的视频文件"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "添加文件夹"
      Height          =   900
      Left            =   1080
      Picture         =   "frmMain.frx":89F6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "点击此处选择待转换的视频文件夹"
      Top             =   0
      Width           =   1050
   End
   Begin VB.Timer timer_Url 
      Interval        =   8000
      Left            =   6795
      Top             =   360
   End
   Begin VB.Frame Frame_Output 
      Height          =   6315
      Left            =   6300
      TabIndex        =   15
      Top             =   900
      Width           =   4155
      Begin VB.CommandButton cmdOpenDir 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   1080
         Picture         =   "frmMain.frx":9838
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "用资源管理器打开输出的文本夹"
         Top             =   4230
         Width           =   555
      End
      Begin VB.CheckBox chkShutdown 
         Caption         =   "视频转换完成后关闭计算机"
         Height          =   285
         Left            =   135
         TabIndex        =   40
         Top             =   5175
         Width           =   3885
      End
      Begin VB.TextBox txtEndTime 
         Height          =   330
         Left            =   2160
         TabIndex        =   38
         Text            =   "0"
         ToolTipText     =   "零表示不启用"
         Top             =   3870
         Width           =   1905
      End
      Begin VB.TextBox txtBeginTime 
         Height          =   330
         Left            =   135
         TabIndex        =   37
         Text            =   "0"
         Top             =   3870
         Width           =   1905
      End
      Begin VB.CommandButton cmdSelectLoc 
         Caption         =   "浏览"
         Height          =   360
         Left            =   3195
         TabIndex        =   25
         ToolTipText     =   "选择视频文件转换后的存放路径"
         Top             =   4725
         Width           =   825
      End
      Begin VB.ComboBox cmbOutFormat 
         Height          =   300
         ItemData        =   "frmMain.frx":9B42
         Left            =   135
         List            =   "frmMain.frx":9B5B
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   540
         Width           =   3930
      End
      Begin VB.ComboBox cmbSoundSD 
         Height          =   300
         ItemData        =   "frmMain.frx":9BE5
         Left            =   2160
         List            =   "frmMain.frx":9BEF
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3060
         Width           =   1905
      End
      Begin VB.ComboBox cmbSoundBT 
         Height          =   300
         ItemData        =   "frmMain.frx":9C03
         Left            =   135
         List            =   "frmMain.frx":9C31
         TabIndex        =   21
         Top             =   3060
         Width           =   1905
      End
      Begin VB.ComboBox cmbSoundBL 
         Height          =   300
         ItemData        =   "frmMain.frx":9C77
         Left            =   2160
         List            =   "frmMain.frx":9C93
         TabIndex        =   20
         Top             =   2250
         Width           =   1905
      End
      Begin VB.ComboBox cmbVideoSize 
         Height          =   300
         ItemData        =   "frmMain.frx":9CCE
         Left            =   135
         List            =   "frmMain.frx":9CF9
         TabIndex        =   19
         Top             =   2250
         Width           =   1905
      End
      Begin VB.ComboBox cmbVideoZL 
         Height          =   300
         ItemData        =   "frmMain.frx":9D72
         Left            =   2160
         List            =   "frmMain.frx":9D7F
         TabIndex        =   18
         Top             =   1350
         Width           =   1905
      End
      Begin VB.ComboBox cmbVideoBT 
         Height          =   300
         ItemData        =   "frmMain.frx":9D8F
         Left            =   135
         List            =   "frmMain.frx":9DBD
         TabIndex        =   17
         Top             =   1350
         Width           =   1905
      End
      Begin VB.TextBox txtLoc 
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   4725
         Width           =   2985
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "点此此处立刻开始转换"
         Height          =   435
         Left            =   135
         TabIndex        =   24
         ToolTipText     =   "我要开始转换喽"
         Top             =   5490
         Width           =   3900
      End
      Begin VB.Label lblInfo 
         Caption         =   "状态显示：系统准备中,请选择待转换的视频"
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   135
         TabIndex        =   41
         Top             =   5985
         Width           =   3795
      End
      Begin VB.Label Label10 
         Caption         =   "视频结束时间（秒）："
         Height          =   240
         Left            =   2160
         TabIndex        =   36
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "视频开始时间（秒）："
         Height          =   240
         Left            =   135
         TabIndex        =   35
         Top             =   3600
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "输出格式："
         Height          =   285
         Left            =   135
         TabIndex        =   33
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "视频比特率："
         Height          =   240
         Left            =   135
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "帧率："
         Height          =   240
         Left            =   2160
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "输出路径："
         Height          =   240
         Left            =   135
         TabIndex        =   30
         Top             =   4410
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "声道："
         Height          =   240
         Left            =   2160
         TabIndex        =   29
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "声音比特率："
         Height          =   240
         Left            =   135
         TabIndex        =   28
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "声音频率："
         Height          =   240
         Left            =   2160
         TabIndex        =   27
         Top             =   1935
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "视频尺寸："
         Height          =   240
         Left            =   135
         TabIndex        =   26
         Top             =   1935
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6330
      Left            =   45
      TabIndex        =   3
      Top             =   900
      Width           =   6330
      Begin VB.TextBox txtAboutMe 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   495
         Width           =   6015
      End
      Begin VB.ListBox lstFiles 
         Height          =   5640
         Left            =   135
         TabIndex        =   13
         Top             =   495
         Width           =   6015
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   5775
         Left            =   135
         TabIndex        =   12
         Top             =   495
         Width           =   6015
         ExtentX         =   10610
         ExtentY         =   10186
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
         Location        =   "http:///"
      End
      Begin VB.Label lblUnlikeURL 
         Caption         =   "我要举报"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5445
         MouseIcon       =   "frmMain.frx":9E07
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Tag             =   "http://www.mama520.cn/IdontLoveYou.php"
         ToolTipText     =   "如果您觉得以下内容是非法信息，请点击此处，告诉我们！"
         Top             =   225
         Width           =   780
      End
      Begin VB.Label lblTitle3 
         Caption         =   "文件列表："
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   180
         Width           =   5145
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   180
      ScaleHeight     =   15
      ScaleWidth      =   4860
      TabIndex        =   2
      Top             =   2295
      Width           =   4860
   End
   Begin VB.Label lblLink 
      Caption         =   "官方网站"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      MouseIcon       =   "frmMain.frx":A111
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Tag             =   "http://www.mama520.cn"
      Top             =   7290
      Width           =   10230
   End
   Begin VB.Menu myMenu 
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu Menu1 
         Caption         =   "菜单一"
      End
      Begin VB.Menu Menu2 
         Caption         =   "菜单一"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sAppPath As String
Private sADs As String '广告链接
Private sRejectLinks As String '被拒绝的网址
Private sRejectSRC As String '被拒绝自动跑掉的网址

Private sRegUser As String '注册用户名或者URL地址
Private sRegCode As String '注册码或者授权码
Private sNickName As String '呢称或者个性名称
Public iSoundOpen As Integer '是否开启声音

Private sSeleFileName '当前所选择的文件
Private sOutPutPath '当前所选择的输出路径

Private iFrmHeight As Integer
Private iFrmHeight_Min As Integer

Private WithEvents webdoc  As Mshtml.HTMLDocument
Attribute webdoc.VB_VarHelpID = -1
Private textbody As HTMLBody
Private Rng As IHTMLTxtRange
Private iADCount As Integer
Private iADIndex As Integer

Private clsMD5 As New clsMD5
Private clsInI As New clsInI
Private IeGetOfflineMode As Boolean

Private WithEvents pThread As MThreadVB.Thread
Attribute pThread.VB_VarHelpID = -1




Private IsRegisterUser As Boolean
Private NeedToChange As Boolean


Private Sub EndThread_Click()
    pThread.TerminateWin32Thread
End Sub


Private Sub EnableControl(ByVal bEnable As Boolean)
    
    Frame_Output.Enabled = bEnable
    
    cmdAddFile.Enabled = bEnable
    cmdAddFolder.Enabled = bEnable
    cmdDel.Enabled = bEnable
    cmdSetting.Enabled = bEnable
    cmdEnd.Enabled = bEnable
    cmdRefresh.Enabled = bEnable
    cmdAbout.Enabled = bEnable
    cmdViewBrowser.Enabled = bEnable
    cmdAD.Enabled = bEnable
    lblLink.Enabled = bEnable
    lblUnlikeURL.Enabled = bEnable
    
     
End Sub


Private Sub cmdOpenDir_Click()
    
    Shell "Explorer.exe " & txtLoc.Text, vbNormalFocus
    
       
    
End Sub

Private Sub pThread_OnThreadCreateFailure()

'    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread could not be Created"
End Sub
     
Private Sub pThread_OnThreadCreateSuccess(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
'    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread Created (Calculations started)"
'    StartThread.Enabled = False
'    EndThread.Enabled = True
NeedToChange = False
cmdConvert.Enabled = False

EnableControl False

Loading
SetStatus "正在转换视频……"


End Sub

Private Sub pThread_OnThreadFinish(ByVal ThreadHandle As Long, ByVal ThreadID As Long)
'    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread has finished running(Calculations Ended)"
'    PText.Text = ""
'    PText.Text = Primes
'    Primes = ""
'    StartThread.Enabled = True
'    EndThread.Enabled = False

Dim sTmp As String

 NeedToChange = True
 cmdConvert.Enabled = True
 
 EnableControl True

 'StartIndexPage
 'BlankPage
 
 SetStatus "视频转换成功！哥们，要继续转换吗？ ^_^"
 
 
   If iSoundOpen = 1 Then
   
      PlaySound sAppPath & "Sound.wav"
   
   End If
   

  If chkShutdown.Value = 1 Then
  
     ShutDownWin
     
  End If
  


 Dim i As Integer
 
 
 
 Dim sOutput As String
 
 
Dim bOk As Boolean

 i = InStr(cmbOutFormat.Text, "(")
 sTmp = left(cmbOutFormat.Text, i - 1)


For i = 0 To lstFiles.ListCount - 1
    
    sOutput = txtLoc.Text & Filename(Trim(lstFiles.List(i))) & "." & sTmp
    
     DeleteFile sOutput & ".TMP.MP4"
    
   
Next

 'MsgBox "转换完成！", vbInformation, sInfo_TiTle_Msg
 


 
End Sub

Private Sub pThread_OnThreadPriorityChange(ByVal ThreadHandle As Long, ThreadID As Long, ByVal OldPriority As MThreadVB.ThreadPriorityConsts, ByVal NewPriority As MThreadVB.ThreadPriorityConsts)
'    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread priority set or changed"
End Sub

Private Sub pThread_OnThreadTerminate(ByVal ThreadHandle As Long, ByVal ThreadID As Long, ByVal ExitCode As Long)
'    ELog.Text = ELog.Text & Chr$(13) & Chr$(10) & "Thread has been forcefully terminated"
'    StartThread.Enabled = True
'    EndThread.Enabled = False

End Sub


'Private Function CheckRegister(ByVal RegUser As String, ByVal RegCode As String) As Boolean
'
'Dim sTrueCode As String
'
'sTrueCode = UCase(clsMD5.CalculateMD5(clsMD5.CalculateMD5(clsMD5.CalculateMD5("Love3GP" & RegUser & "520"))))
'
'If UCase(RegCode) = sTrueCode Then
'   CheckRegister = True
'
'Else
'    CheckRegister = False
'
'End If
'
'
'
'
'End Function



Private Function CheckConvert() As Boolean


    CheckConvert = False
           
    If lstFiles.ListCount = 0 Then
          
          MsgBox sInfo_MustSele, vbInformation, sInfo_TiTle_Msg
          Exit Function
    End If
    
    
    If txtLoc.Text = "" Then
          
          MsgBox sInfo_OutLoc, vbInformation, sInfo_TiTle_Msg
          txtLoc.SetFocus
          Exit Function
    End If
    

    If cmbOutFormat.ListIndex < 0 Then
    
          MsgBox "请选择输出的格式", vbInformation, sInfo_TiTle_Msg
          cmbOutFormat.SetFocus
          Exit Function
    
    End If
            
        If Not IsNumeric(cmbVideoBT.Text) Then
            MsgBox "视频比特率必须为数字！", vbInformation, sInfo_TiTle_Msg
            cmbVideoBT.SetFocus
             Exit Function
        End If

        If Not IsNumeric(cmbVideoZL.Text) Then
            MsgBox "帧率必须为数字！", vbInformation, sInfo_TiTle_Msg
            cmbVideoZL.SetFocus
             Exit Function
        End If
        
        If cmbVideoSize.Text = "" Then
    
          MsgBox "请必须选择或者输入视频尺寸", vbInformation, sInfo_TiTle_Msg
          cmbVideoSize.SetFocus
          Exit Function
    
        End If
        
        
        If Not IsNumeric(cmbSoundBL.Text) Then
            MsgBox "声道频率必须为数字！", vbInformation, sInfo_TiTle_Msg
            cmbSoundBL.SetFocus
             Exit Function
        End If
        
        
    
        If cmbSoundBT.Text = "" Then
        
              MsgBox "请必须选择或者输入声音比特率", vbInformation, sInfo_TiTle_Msg
              cmbSoundBT.SetFocus
              Exit Function
        
        End If
        
        If cmbSoundSD.Text = "" Then
        
              MsgBox "请必须选择或者输入声道", vbInformation, sInfo_TiTle_Msg
              cmbSoundSD.SetFocus
              Exit Function
        
        End If

    
   CheckConvert = True
    
    
    
End Function

Private Sub SetStatus(ByVal Info As String)

    lblInfo.Caption = Info
    
End Sub



Private Sub cmbOutFormat_Change()
' Dim i As Integer
' Dim sTmp As String
'
' i = InStr(cmbOutFormat.Text, "(")
' sTmp = left(cmbOutFormat.Text, i - 1)
'
'
' Select Case sTmp
'  Case "3GP"
'
'  Case "RMVB", "RM"
'
'  Case "MPG"
'  Case Else
' End Select
 
 
 
 
 
End Sub

Private Sub SetCMBValue(cmbVideoBTValue As String, cmbVideoZLValue As String, cmbVideoSizeValue As String, cmbSoundBLValue As String, cmbSoundBTValue As String, cmbSoundSDValue As String)
    cmbVideoBT.Text = cmbVideoBTValue
    cmbVideoZL.Text = cmbVideoZLValue
    cmbVideoSize.Text = cmbVideoSizeValue
    cmbSoundBL.Text = cmbSoundBLValue
    cmbSoundBT.Text = cmbSoundBTValue
    cmbSoundSD.Text = cmbSoundSDValue
End Sub

Private Sub cmbOutFormat_Click()
 Dim i As Integer
 Dim sTmp As String
 
 i = InStr(cmbOutFormat.Text, "(")
 sTmp = left(cmbOutFormat.Text, i - 1)
 

 
 
'3GP (手机3GP格式)
'MP4 (手机MP4格式)
'RMVB (RealPlayerRMVB)
'RM (RealPlayerRM)
'MPG (MPEG2)
'AVI(Windows Media Player格式)
'MP4 (高清MPEG4格式)
'MOV(Quick Time格式)
'MP3 (纯音频格式)
 
 Select Case cmbOutFormat.ListIndex
  Case 0
'   MsgBox ("3GP")
'Before PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\Film\Net\韩国最美的美女裸体.rmvb" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
'After  PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\Film\Net\韩国最美的美女裸体.rmvb" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4


'Before PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\韩国最美的美女裸体-233755.3gp"
'After  PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\韩国最美的美女裸体-233755.3gp"




'Before PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\Media\玉蒲团之玉女心经.rmvb" -ss 0 -endpos 30 -oac lavc -lavcopts acodec=mp2:abitrate=64 -ovc lavc -lavcopts vcodec=mpeg2video:vbitrate=600:vpass=1 -ofps 24 -of lavf -lavfopts format=yuv4mpeg -o "c:\output\玉蒲团之玉女心经-23339.avi"
'After  PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\Media\玉蒲团之玉女心经.rmvb" -ss 0 -endpos 30 -oac lavc -lavcopts acodec=mp2:abitrate=64 -ovc lavc -lavcopts vcodec=mpeg2video:vbitrate=600:vpass=1 -ofps 24 -of lavf -lavfopts format=yuv4mpeg -o "c:\output\玉蒲团之玉女心经-23339.avi"


   SetCMBValue "128", "29", "320x240", "8000", "12.2", "单声道"
  Case 1
'    MsgBox ("MP4")
    SetCMBValue "128", "29", "320x240", "8000", "12.2", "单声道"
  Case 2
'    MsgBox ("RMVB")
    SetCMBValue "1200", "29", "原始大小", "44100", "320", "双声道"
  Case 3
'    MsgBox ("RM")
    SetCMBValue "1200", "29", "原始大小", "44100", "320", "双声道"
'  Case 4
''    MsgBox ("MPG")
'    SetCMBValue "600", "24", "原始大小", "28000", "64", "双声道"
  Case 5 - 1
'    MsgBox ("AVI")
    SetCMBValue "600", "24", "原始大小", "28000", "64", "双声道"
  Case 6 - 1
'    MsgBox ("MP4")
    SetCMBValue "1200", "29", "原始大小", "48000", "320", "双声道"
  Case 7 - 1
'    MsgBox ("MOV")
    SetCMBValue "256", "29", "原始大小", "48000", "64", "双声道"
  Case 8 - 1
'    MsgBox ("MP3")
    SetCMBValue "256", "29", "原始大小", "48000", "320", "双声道"
 End Select
 
 
 
End Sub

Private Sub cmdAD_Click()
frmAD.Show 1
End Sub

Private Sub cmdAddFolder_Click()

     Dim sLoc As String
     Dim sFiles() As String
     Dim i As Long
     Dim j As Long
     
     
     
     
     ViewControl (1)
     sLoc = BrowseForFolderByPath(Me.hWnd, "请选择待转换视频的文件夹", sAppPath)
     
     If sLoc <> "" Then
        i = FileList(sLoc, sFiles)
     End If
     
     For j = 0 To i - 1
       If InStr(sFiles(j), ".") > 0 Then
        lstFiles.AddItem (IIf(right(sLoc, 1) <> "\", sLoc & "\", sLoc) & sFiles(j))
       End If
     Next
     
     If lstFiles.ListCount > 0 Then
        SetStatus "待转换的视频文件已添加至待转换列表！"
     End If
     
        
'    lstFiles.ForeColor = &HFF&
'
'    lstFiles.List(0) = lstFiles.List(0) & " <转换成功>"
    
    
    

     
End Sub

Private Sub cmdConvert_Click()

 Dim i As Integer
 
 
 Dim sTmp As String
 Dim sOutput As String
 

On Error GoTo errCmdConvert








If Not CheckConvert() Then Exit Sub

Dim iStyle As Integer


If right(txtLoc.Text, 1) <> "\" Then

  txtLoc.Text = txtLoc.Text + "\"
  
End If


 i = InStr(cmbOutFormat.Text, "(")
 sTmp = left(cmbOutFormat.Text, i - 1)
 



If Not PathExists(txtLoc.Text) Then

        CreateDirectory (txtLoc.Text)
        
End If


'Dim bOK As Boolean
'
'
'
'Loading
'
'cmdConvert.Enabled = False


'ConvertNow 0
'Set pThread = New MThreadVB.Thread


pThread.CreateWin32Thread Me, "ConvertNow", 0

pThread.ThreadPriority = THREAD_PRIORITY_HIGHEST

Exit Sub

errCmdConvert:

MsgBox " 错误号:  " & Err.Number & vbCrLf & _
       " 错误内容:" & Err.Description & vbCrLf & _
       " 错误位置:errCmdConvert", vbExclamation + vbOKOnly




'Select Case Pr.List(Pr.ListIndex)
'    Case "Lowest"
'        pThread.ThreadPriority = THREAD_PRIORITY_LOWEST
'    Case "Below normal"
'        pThread.ThreadPriority = THREAD_PRIORITY_BELOW_NORMAL
'    Case "Normal"
'pThread.ThreadPriority = THREAD_PRIORITY_NORMAL
'    Case "Above Normal"
'        pThread.ThreadPriority = THREAD_PRIORITY_ABOVE_NORMAL
'    Case "Highest"
       ' pThread.ThreadPriority = THREAD_PRIORITY_HIGHEST
'End Select


'DoEvents

'If 1 = 1 Then
'
'    bOK = StartPDFtoText(txtPDFFile.Text, txtLoc.Text & sSeleFileName + ".txt", iStyle, True, 0, 0, txtOwnerPWD.Text, txtUserPWD.Text)
'Else
'    bOK = StartPDFtoText(txtPDFFile.Text, txtLoc.Text & sSeleFileName + ".txt", iStyle, False, txtPrePage.Text, txtLastPage.Text, txtOwnerPWD.Text, txtUserPWD.Text)
'
'
'End If

'For i = 0 To lstFiles.ListCount - 1
'
'
'
'
'    sOutput = txtLoc.Text & Filename(Trim(lstFiles.List(i))) & "." & sTmp
'    bOK = StartConvert(lstFiles.List(i), sOutput, cmbOutFormat.ListIndex, cmbVideoBT.Text, cmbVideoZL.Text, cmbVideoSize.Text, cmbSoundBL.Text, cmbSoundBT.Text, cmbSoundSD.Text)
'
'  ' MsgBox i
'
'Next





'If bOK Then
'         MsgBox "转换成功！", vbInformation, sInfo_TiTle_Msg
'Else
'         MsgBox "转换失败喽，不好意思，再试一次！", vbCritical, sInfo_TiTle_Msg
'End If

'cmdConvert.Enabled = True
'
'StartIndexPage
'
'ViewControl (2)
  


End Sub

Sub ConvertNow(DummyArgument As Variant)


 Dim i As Integer
 
 
 Dim sTmp As String
 Dim sOutput As String
 
 
Dim bOk As Boolean

 i = InStr(cmbOutFormat.Text, "(")
 sTmp = left(cmbOutFormat.Text, i - 1)


For i = 0 To lstFiles.ListCount - 1


    SetStatus "正在转换：" & Filename(Trim(lstFiles.List(i)))
    
    sOutput = txtLoc.Text & Filename(Trim(lstFiles.List(i))) & "." & sTmp
    bOk = StartConvert(lstFiles.List(i), sOutput, cmbOutFormat.ListIndex, cmbVideoBT.Text, cmbVideoZL.Text, cmbVideoSize.Text, cmbSoundBL.Text, cmbSoundBT.Text, cmbSoundSD.Text)

  ' MsgBox i
   
Next



End Sub



Private Function StartConvert(ByVal InFile As String, ByVal OutFile As String, ByVal Options As Integer, cmbVideoBTValue As String, cmbVideoZLValue As String, cmbVideoSizeValue As String, cmbSoundBLValue As String, cmbSoundBTValue As String, cmbSoundSDValue As String) As Boolean

Dim Params As String
Dim sAll As String
Dim sShellCD As String

sAll = ""




sShellCD = sAppPath + sShellPath
StartConvert = False

    If (OutFile <> Empty) Then
    
        If FileExists(OutFile) Then DeleteFile OutFile
        
        
      
        sAll = "something"
        
    
        
        
        
        
'        Select Case PDFtoTextOptions
'            Case 1
'                'pdftotext1.exe | Options: -raw
'                Params = "-enc GBK -layout " & sAll & " """ & InFile & """ """ & OutFile & """" 'for pdftotext1.exe
'                Execute sAppPath & sShell & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
'            Case 2
'                'pdftotext3.exe | Options: -layout
'                Params = "-enc GBK -raw " & sAll & " """ & InFile & """ """ & OutFile & """"  'for pdftotext3.exe
'                Execute sAppPath & sShell & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
'            Case 3
'                'pdftotext3.exe | Options: -raw -layout
'                Params = "-enc GBK -raw -layout " & sAll & " """ & InFile & """ """ & OutFile & """"
'                Execute sAppPath & sShell & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
'
'        End Select
        
        
            Select Case Options
                Case 0
                '   MsgBox ("3GP")
                
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\Girl.flv" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\Girl-224610.3gp"
                
                
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\透明时装秀迷人的中国时装秀-23948.3gp"
                
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\全球透明时装秀 - 韩国美__.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4

                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 128x96 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\全球透明时装秀 - 韩国美__-15850.3gp"

                  '步骤一，输出到临时mp4文件中去。
                  
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '步骤二，正式输出到对应的输出文件中去。
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y " & IIf(cmbVideoSizeValue = "原始尺寸", "", " -s " & cmbVideoSizeValue) & " -vcodec mpeg4 -acodec amr_nb -ac 1 -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k """ & OutFile & """"
                  Execute sAppPath & sShell_Coder2 & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  
'                      cmbVideoBT.Text = cmbVideoBTValue
'    cmbVideoZL.Text = cmbVideoZLValue
'    cmbVideoSize.Text = cmbVideoSizeValue
'    cmbSoundBL.Text = cmbSoundBLValue
'    cmbSoundBT.Text = cmbSoundBTValue
'    cmbSoundSD.Text = cmbSoundSDValue

                
                 'SetCMBValue "128", "29", "320x240", "8000", "12.2", "单声道"
                Case 1
                
                
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\透明时装秀迷人的中国时装秀-23948.3gp"
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 480x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\透明时装秀迷人的中国时装秀-1231.mp4"
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '步骤二，正式输出到对应的输出文件中去。
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "原始尺寸", "", " -s " & cmbVideoSizeValue) & " -vcodec mpeg4 -acodec amr_nb -ac 1 -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k """ & OutFile & """"
                  Execute sAppPath & sShell_Coder2 & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD


                '    MsgBox ("MP4")
                  'SetCMBValue "128", "29", "320x240", "8000", "12.2", "单声道"
                Case 2
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                 'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec rv10 -b 1000k -f rm -ar 24000 -ab 320k -ac 双声道 "c:\output\透明时装秀迷人的中国时装秀-1212.rmvb"

                '    MsgBox ("RMVB")
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '步骤二，正式输出到对应的输出文件中去。
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "原始尺寸", "", " -s " & cmbVideoSizeValue) & " -vcodec rv10 -b " & cmbVideoBTValue & "k -f rm -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  'SetCMBValue "1200", "29", "原始大小", "44100", "320", "双声道"
                Case 3
                
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec rv10 -b 1200k -f rm -ar 44100 -ab 320k -ac 双声道 "c:\output\透明时装秀迷人的中国时装秀-115533.rm"
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '步骤二，正式输出到对应的输出文件中去。
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "原始尺寸", "", " -s " & cmbVideoSizeValue) & " -vcodec rv10 -b " & cmbVideoBTValue & "k -f rm -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD


                '    MsgBox ("RM")
                  'SetCMBValue "1200", "29", "原始大小", "44100", "320", "双声道"
                '  Case 4
                ''    MsgBox ("MPG")
                '    SetCMBValue "600", "24", "原始大小", "28000", "64", "双声道"
'                Case 5
'                '    MsgBox ("AVI")
'                  SetCMBValue "600", "24", "原始大小", "28000", "64", "双声道"
                Case 6 - 2
                '    MsgBox ("MP4")
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o "c:\output\透明时装秀迷人的中国时装秀-122232.mp4"
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & """"
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD

                  'SetCMBValue "1200", "29", "原始大小", "48000", "320", "双声道"
                Case 7 - 2
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\conv.exe "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -ac 2 -ar 48000 -ab 128k "c:\output\透明时装秀迷人的中国时装秀-123048.mov"
                
                                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '步骤二，正式输出到对应的输出文件中去。
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "原始尺寸", "", " -s " & cmbVideoSizeValue) & "  -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD

                '    MsgBox ("MOV")
                  'SetCMBValue "256", "29", "原始大小", "48000", "64", "双声道"
                Case 8 - 2
                '    MsgBox ("MP3")
                'PAnsiChar  lpCommandLine: C:\3GP格式转换器\data\coder.exe -i "E:\MTV\透明时装秀迷人的中国时装秀.mp4" -y -ss 0 -t 30 -y -vn -ar 48000 -ab 320k -ac 2 "c:\output\透明时装秀迷人的中国时装秀-124145.mp3"
                
                  Params = " -i """ & InFile & """ -y -y -vn -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD


                  'SetCMBValue "256", "29", "原始大小", "48000", "320", "双声道"
            End Select
        
        
        
        
        If FileExists(OutFile) Then StartConvert = True
        
    End If
End Function


Private Sub ViewControl(bViewSele As Integer)
On Error Resume Next

lblTitle3.ForeColor = &H80000012

    Select Case bViewSele
       Case 1
           lblTitle3.Caption = "文件名称"
           lstFiles.Visible = True
           WebBrowser.Visible = False
           txtAboutMe.Visible = False

       Case 2
         
'           If left(sIAgreeUrl, 10) <> "http://www.mama520.cn/" Then
         If sRegUser = sADUrl Then
                lblTitle3.Caption = "迷你首页"
                lblUnlikeURL.Visible = False
           Else
                If CheckRegister(sRegUser, sRegCode, False) And Not CheckContactURL(sRegUser, sRegCode) Then
                    lblTitle3.ForeColor = &HC00000
                    lblTitle3.Caption = "★ 以下内容非官方迷你首页提供，显示内容为个性定制页面 ★"
                End If
                
                If CheckRegister(sRegUser, sRegCode, False) And CheckContactURL(sRegUser, sRegCode) Then
                    lblTitle3.ForeColor = 49152
                    lblTitle3.Caption = "★ 以下内容非官方迷你首页提供，但已通过合法性验证 ★"
                End If
                
                
                If Not CheckRegister(sRegUser, sRegCode, False) Then
                    lblTitle3.ForeColor = &HFF&
                    lblTitle3.Caption = "★ 个性网址授权码错误，如内容非法请点击“我要举报” ★"
                End If
                
                
           End If
           
           
           
           
           
           lstFiles.Visible = False
           WebBrowser.Visible = True
           txtAboutMe.Visible = False
       Case Else
           lblTitle3.Caption = "关于软件"
           lstFiles.Visible = False
           WebBrowser.Visible = False
           txtAboutMe.Visible = True
       
    End Select
    
    
    

    
End Sub


Private Sub cmdDel_Click()

Dim iTmp As Integer

    iTmp = lstFiles.ListIndex

    If lstFiles.ListCount > 0 And iTmp >= 0 Then
    
        
        lstFiles.RemoveItem (iTmp)
        lstFiles.ListIndex = iTmp - 1
        
    End If
    
End Sub

Private Sub cmdEnd_Click()
    
    Unload Me
    End
    
    
    
End Sub

Private Sub cmdRefresh_Click()

    ViewControl (1)

End Sub

Private Sub cmdSelectLoc_Click()

     'txtLoc.Text = BrowseForFolder(Me, "请选择输出路径")
     txtLoc.Text = BrowseForFolderByPath(Me.hWnd, "请选择输出路径", sAppPath)
    
End Sub

Private Sub cmdAddFile_Click()
    'txtPDFFile.Text = BrowseForFolderByPath(Me.Hwnd, "请选择PDF文件", sAppPath)
    
    
     Dim iRet As Long
     Dim sFileName As String
     
     
     ViewControl (1)
     
     iRet = DLG_Open(Me.hWnd, "*.*|*.*", vbNullString, "*", 0, "请选择", True, sFileName, 0)
     
     If sFileName <> "" Then
     
        lstFiles.AddItem (sFileName)
     
     End If
     
     
     
     
       
     
     sSeleFileName = Filename(sFileName)
     
     If lstFiles.ListCount > 0 Then
        SetStatus "待转换的视频文件已添加至待转换列表！"
     End If
     
     
  
     
   '  sSeleFileName = ShortPath("C:\Documents and Settings\Administrator\桌面\网络安全")
     
   '  sSeleFileName = ShortPath("C:\Documents and Settings\Administrator\桌面\网络安全")
     
End Sub



Private Sub cmdSetting_Click()

    frmMsg.Show 1
    
End Sub



Private Sub Form_Initialize()

SkinH_Attach
'    Dim ctl As Control
'    Dim oFont As StdFont
'
'    On Error Resume Next
'
'    For Each ctl In Controls
'        Set oFont = ctl.Font
'        With oFont
'            .Name = "宋体"
'            .Charset = 134
'        End With
'    Next

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sTmp As String

'
'SkinH_Attach
NeedToChange = False

Set pThread = New MThreadVB.Thread
    
    If right(App.Path, 1) <> "\" Then
        sAppPath = App.Path + "\"
    Else
        sAppPath = App.Path
    End If
    
  '/////////----------------------读取是授权配置文档--------------------------
  
  clsInI.Filename = sAppPath & cRegister_File
  
  If clsInI.CheckIniFile(clsInI.Filename) Then
  
   sNickName = clsInI.GetKeyFromSection(cNickName_Key, cRegister_Section)
  
   sRegUser = clsInI.GetKeyFromSection(cRegUser_Key, cRegister_Section)
  
   sRegCode = clsInI.GetKeyFromSection(cRegCode_Key, cRegister_Section)
   
   sTmp = clsInI.GetKeyFromSection(cSoundOpen_Key, cRegister_Section)
  
  End If
  
  
  iSoundOpen = CInt(sTmp)
  
  If iSoundOpen = 1 Then
  
      frmMsg.OptOpen = True
      
  Else
      
      frmMsg.OptClose = True
      
  End If
  
  
    
  If sRegUser <> "" Then
      frmAD.txtADUser.Text = sNickName
      frmAD.txtADUrl.Text = sRegUser
      frmAD.txtADRegisterCode.Text = sRegCode
  End If
  
  If CheckRegister(sRegUser, sRegCode, False) Then
    IsRegisterUser = True
  Else
    IsRegisterUser = False
  End If
  
  RefreshVer
  
  
  
  
  
  'sIAgreeUrl = "http://www.mama520.cn/software_AD/3GP/IAgree.html"
  
  DownloadFromWeb "http://www.mama520.cn/software_AD/3GP/WebConfig.txt", sAppPath & "LocalMiniWeb\" & cWebConfig_File
  
  '/////////----------------------读取WEB控制台配置文档--------------------------
  
  If FileExists(sAppPath & "LocalMiniWeb\" & cWebConfig_File) Then
    IeGetOfflineMode = False
    MoveFile sAppPath & "LocalMiniWeb\" & cWebConfig_File, sAppPath & cWebConfig_File, True
    
 Else
    IeGetOfflineMode = True
  End If
  
    clsInI.Filename = sAppPath & cRegister_File
  
  If clsInI.CheckIniFile(clsInI.Filename) Then
  
   sNickName = clsInI.GetKeyFromSection(cNickName_Key, cRegister_Section)
  
   sRegUser = clsInI.GetKeyFromSection(cRegUser_Key, cRegister_Section)
  
   sRegCode = clsInI.GetKeyFromSection(cRegCode_Key, cRegister_Section)
   
  
  End If
  
  
  


    clsInI.Filename = sAppPath & cWebConfig_File
  
  If clsInI.CheckIniFile(clsInI.Filename) Then
  
   sADs = clsInI.GetKeyFromSection(cADLinks_Key, cWebConfig_Section)
  
   sRejectLinks = clsInI.GetKeyFromSection(cRejectLinks_Key, cWebConfig_Section)
  
   sRejectSRC = clsInI.GetKeyFromSection(cRejectSRC_Key, cWebConfig_Section)
  
  End If
  
  
  
  'sADs = ReadFromFile(sAppPath & "WebConfig.dat")
  
 

    
  WebBrowser.left = txtAboutMe.left
  WebBrowser.top = txtAboutMe.top
  WebBrowser.Width = txtAboutMe.Width
  WebBrowser.Height = txtAboutMe.Height
  
  
  iFrmHeight = Me.Height
  
  iFrmHeight_Min = txtAboutMe.top + cmdConvert.Height - 80
  

  
  If InStr(sRejectLinks, sRegCode) <> 0 Then
  
     RejectPage sRejectSRC
  Else
     StartIndexPage
     
  End If
  
     
  
  cmbOutFormat.ListIndex = 0

  
  
  ViewControl (2)

  
  
  
        
End Sub


Private Sub RejectPage(ByVal sRejectURL As String)

  Dim sTmp As String
  
  ViewControl (2)
  
  
  sTmp = Replace(sAppPath, "\", "/")
  If IeGetOfflineMode Then
    sTmp = "file:///" + sTmp + "localMiniWeb/index.html"
  Else
    sTmp = sRejectURL
  End If
  
  
  WebBrowser.Navigate sTmp

End Sub


Private Sub StartIndexPage()

  Dim sTmp As String
  
  ViewControl (2)
  
  sTmp = Replace(sAppPath, "\", "/")
  If IeGetOfflineMode Then
    sTmp = "file:///" + sTmp + "localMiniWeb/index.html"
  Else
    sTmp = sADUrl
  End If
  

  
  WebBrowser.Navigate sTmp

End Sub


Private Sub Loading()

  Dim sTmp As String
  
  ViewControl (2)
  
  sTmp = Replace(sAppPath, "\", "/")
  
  sTmp = "file:///" + sTmp + "localMiniWeb/loading.html"
  
  WebBrowser.Navigate sTmp

End Sub



Private Sub BlankPage()

  Dim sTmp As String
  
  ViewControl (2)
  
  
  sTmp = "about:blank"
  
  WebBrowser.Navigate sTmp

End Sub

Private Sub RefreshVer()

If IsRegisterUser Then
    Me.Caption = sCaption_REG
Else
    Me.Caption = sCaption_Free
End If
Me.Refresh


End Sub


Private Sub ViewBrowser(ByVal bVisible As Boolean)

    On Error Resume Next
    WebBrowser.Visible = bVisible
    txtAboutMe.Visible = Not bVisible

End Sub

Private Sub cmdViewBrowser_Click()
On Error Resume Next
ViewControl (2)

End Sub

Private Sub cmdAbout_Click()
On Error Resume Next
    Dim i As Integer
    i = 5
    
    txtAboutMe.Text = Space(i) & "软件名称：" & frmMain.Caption & vbCrLf & _
                    Space(i) & "开发作者：黄启清" & vbCrLf & _
                    Space(i) & "腾讯微博：http://t.qq.com/huangqiqing   (记得成为我的粉丝)" & vbCrLf & _
                    Space(i) & "新浪微博：http://t.sina.com.cn/happyqq  (记得成为我的粉丝)" & vbCrLf & _
                    Space(i) & "开发日期：2011-2-28" & vbCrLf & _
                    Space(i) & "官方网站：http://www.mama520.cn" & vbCrLf & _
                    Space(i) & "邮件地址：mama520.cn@gmail.com" & vbCrLf & _
                    Space(i) & "特别感谢开源项目FFMEPG,MEncoder等视频/音频转码项目的开发者们！"

    
ViewControl (3)
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    If pThread Is Nothing Then
        Unload Me
        End
    End If
    
    If pThread.IsThreadRunning = True Then
        MsgBox "正在执行操作，请稍等片刻，等我转换完吧！^_^"
        Cancel = True
        Exit Sub
    End If
    
    Set pThread = Nothing
    
    Unload Me
    End
    
End Sub

Private Sub Form_Terminate()
 Unload Me
 End
End Sub

'Private Sub Form_Unload(Cancel As Integer)
' Unload Me
' End
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
' Unload Me
'End Sub

Private Sub lblLink_Click()
OpenLink (lblLink.Tag)
End Sub

Private Sub lblUnlikeURL_Click()
OpenLink (lblUnlikeURL.Tag & "?Cat=3GP&HS=" & frmAD.txtADRegisterCode & "&URL=" & frmAD.txtADUrl)
End Sub

Private Sub timer_Load_Timer()
Loading
End Sub

Private Sub Timer_Refresh_Timer()

If NeedToChange = True Then

    StartIndexPage

End If

NeedToChange = False

End Sub

Private Sub timer_Url_Timer()

On Error Resume Next

Dim sTitleAD() As String
Dim sTitleURL() As String
'Dim sADs As String
'Dim i As Integer

'sADs = "========================================开心网，点点我就可以进入喽。=================================|||http://www.kaixin.com$$$人人网，我们都一起玩！|||http://www.renren.com$$$妈妈微博，最唠叨的就是你！|||http://www.mama520.cn"

sTitleAD = Split(sADs, "$$$")
iADCount = UBound(sTitleAD)

'For i = 0 To UBound(sTitleAD)

sTitleURL = Split(sTitleAD(iADIndex), "|||")

 lblLink.Caption = sTitleURL(0)
 lblLink.Tag = sTitleURL(1)
 lblLink.ToolTipText = sTitleURL(1)
 
'
'Next i
If iADIndex < iADCount Then
   iADIndex = iADIndex + 1
Else
   iADIndex = 0
End If

   
'If NeedToChange = True Then
'
'    StartIndexPage
'
'End If
'
'NeedToChange = False





End Sub

Private Sub WebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)

On Error Resume Next

   Dim i As Integer

    '将webdoc设置到被WebBrowser的Document属性返回的文挡对象中
    Set webdoc = WebBrowser.Document



End Sub

Private Function webdoc_oncontextmenu() As Boolean
webdoc_oncontextmenu = False
'Me.PopupMenu myMenu
End Function

