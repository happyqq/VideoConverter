VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3GP ת������ Beta 1.0 "
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer_Refresh 
      Interval        =   100
      Left            =   6660
      Top             =   45
   End
   Begin VB.CommandButton cmdAD 
      Caption         =   "���Զ���"
      Height          =   900
      Left            =   9405
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "������������ʾ��ĸ��������𣿿����Ұɣ�"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdViewBrowser 
      Caption         =   "������ҳ"
      Height          =   900
      Left            =   8370
      Picture         =   "frmMain.frx":0D88
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "��ʾ������ҳ"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "�������"
      Height          =   900
      Left            =   7335
      Picture         =   "frmMain.frx":13B9
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "��������ļ򵥽���"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "��ʾ�б�"
      Height          =   900
      Left            =   5220
      Picture         =   "frmMain.frx":1684
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "����˴���ʾ��ת�����ļ��б�"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "�˳�����"
      Height          =   900
      Left            =   4185
      Picture         =   "frmMain.frx":337E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "�˳�����"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "��ʾ����"
      Height          =   900
      Left            =   3150
      Picture         =   "frmMain.frx":41C0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "����˴�����ת����ɺ���������ѹ���"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "�Ƴ��ļ�"
      Height          =   900
      Left            =   2115
      Picture         =   "frmMain.frx":5EBA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "����˴����ļ��б����Ƴ���ת�����ļ�"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "����ļ�"
      Height          =   900
      Left            =   45
      Picture         =   "frmMain.frx":6CFC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "����˴�ѡ���ת������Ƶ�ļ�"
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "����ļ���"
      Height          =   900
      Left            =   1080
      Picture         =   "frmMain.frx":89F6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "����˴�ѡ���ת������Ƶ�ļ���"
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
         ToolTipText     =   "����Դ��������������ı���"
         Top             =   4230
         Width           =   555
      End
      Begin VB.CheckBox chkShutdown 
         Caption         =   "��Ƶת����ɺ�رռ����"
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
         ToolTipText     =   "���ʾ������"
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
         Caption         =   "���"
         Height          =   360
         Left            =   3195
         TabIndex        =   25
         ToolTipText     =   "ѡ����Ƶ�ļ�ת����Ĵ��·��"
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
         Caption         =   "��˴˴����̿�ʼת��"
         Height          =   435
         Left            =   135
         TabIndex        =   24
         ToolTipText     =   "��Ҫ��ʼת���"
         Top             =   5490
         Width           =   3900
      End
      Begin VB.Label lblInfo 
         Caption         =   "״̬��ʾ��ϵͳ׼����,��ѡ���ת������Ƶ"
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   135
         TabIndex        =   41
         Top             =   5985
         Width           =   3795
      End
      Begin VB.Label Label10 
         Caption         =   "��Ƶ����ʱ�䣨�룩��"
         Height          =   240
         Left            =   2160
         TabIndex        =   36
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "��Ƶ��ʼʱ�䣨�룩��"
         Height          =   240
         Left            =   135
         TabIndex        =   35
         Top             =   3600
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "�����ʽ��"
         Height          =   285
         Left            =   135
         TabIndex        =   33
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "��Ƶ�����ʣ�"
         Height          =   240
         Left            =   135
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "֡�ʣ�"
         Height          =   240
         Left            =   2160
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "���·����"
         Height          =   240
         Left            =   135
         TabIndex        =   30
         Top             =   4410
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "������"
         Height          =   240
         Left            =   2160
         TabIndex        =   29
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "���������ʣ�"
         Height          =   240
         Left            =   135
         TabIndex        =   28
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "����Ƶ�ʣ�"
         Height          =   240
         Left            =   2160
         TabIndex        =   27
         Top             =   1935
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "��Ƶ�ߴ磺"
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
            Name            =   "����_GB2312"
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
         Caption         =   "��Ҫ�ٱ�"
         BeginProperty Font 
            Name            =   "����"
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
         ToolTipText     =   "������������������ǷǷ���Ϣ�������˴����������ǣ�"
         Top             =   225
         Width           =   780
      End
      Begin VB.Label lblTitle3 
         Caption         =   "�ļ��б�"
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
      Caption         =   "�ٷ���վ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu Menu1 
         Caption         =   "�˵�һ"
      End
      Begin VB.Menu Menu2 
         Caption         =   "�˵�һ"
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
Private sADs As String '�������
Private sRejectLinks As String '���ܾ�����ַ
Private sRejectSRC As String '���ܾ��Զ��ܵ�����ַ

Private sRegUser As String 'ע���û�������URL��ַ
Private sRegCode As String 'ע���������Ȩ��
Private sNickName As String '�سƻ��߸�������
Public iSoundOpen As Integer '�Ƿ�������

Private sSeleFileName '��ǰ��ѡ����ļ�
Private sOutPutPath '��ǰ��ѡ������·��

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
SetStatus "����ת����Ƶ����"


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
 
 SetStatus "��Ƶת���ɹ������ǣ�Ҫ����ת���� ^_^"
 
 
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

 'MsgBox "ת����ɣ�", vbInformation, sInfo_TiTle_Msg
 


 
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
    
          MsgBox "��ѡ������ĸ�ʽ", vbInformation, sInfo_TiTle_Msg
          cmbOutFormat.SetFocus
          Exit Function
    
    End If
            
        If Not IsNumeric(cmbVideoBT.Text) Then
            MsgBox "��Ƶ�����ʱ���Ϊ���֣�", vbInformation, sInfo_TiTle_Msg
            cmbVideoBT.SetFocus
             Exit Function
        End If

        If Not IsNumeric(cmbVideoZL.Text) Then
            MsgBox "֡�ʱ���Ϊ���֣�", vbInformation, sInfo_TiTle_Msg
            cmbVideoZL.SetFocus
             Exit Function
        End If
        
        If cmbVideoSize.Text = "" Then
    
          MsgBox "�����ѡ�����������Ƶ�ߴ�", vbInformation, sInfo_TiTle_Msg
          cmbVideoSize.SetFocus
          Exit Function
    
        End If
        
        
        If Not IsNumeric(cmbSoundBL.Text) Then
            MsgBox "����Ƶ�ʱ���Ϊ���֣�", vbInformation, sInfo_TiTle_Msg
            cmbSoundBL.SetFocus
             Exit Function
        End If
        
        
    
        If cmbSoundBT.Text = "" Then
        
              MsgBox "�����ѡ�������������������", vbInformation, sInfo_TiTle_Msg
              cmbSoundBT.SetFocus
              Exit Function
        
        End If
        
        If cmbSoundSD.Text = "" Then
        
              MsgBox "�����ѡ�������������", vbInformation, sInfo_TiTle_Msg
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
 

 
 
'3GP (�ֻ�3GP��ʽ)
'MP4 (�ֻ�MP4��ʽ)
'RMVB (RealPlayerRMVB)
'RM (RealPlayerRM)
'MPG (MPEG2)
'AVI(Windows Media Player��ʽ)
'MP4 (����MPEG4��ʽ)
'MOV(Quick Time��ʽ)
'MP3 (����Ƶ��ʽ)
 
 Select Case cmbOutFormat.ListIndex
  Case 0
'   MsgBox ("3GP")
'Before PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\Film\Net\������������Ů����.rmvb" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
'After  PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\Film\Net\������������Ů����.rmvb" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4


'Before PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\������������Ů����-233755.3gp"
'After  PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\������������Ů����-233755.3gp"




'Before PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\Media\������֮��Ů�ľ�.rmvb" -ss 0 -endpos 30 -oac lavc -lavcopts acodec=mp2:abitrate=64 -ovc lavc -lavcopts vcodec=mpeg2video:vbitrate=600:vpass=1 -ofps 24 -of lavf -lavfopts format=yuv4mpeg -o "c:\output\������֮��Ů�ľ�-23339.avi"
'After  PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\Media\������֮��Ů�ľ�.rmvb" -ss 0 -endpos 30 -oac lavc -lavcopts acodec=mp2:abitrate=64 -ovc lavc -lavcopts vcodec=mpeg2video:vbitrate=600:vpass=1 -ofps 24 -of lavf -lavfopts format=yuv4mpeg -o "c:\output\������֮��Ů�ľ�-23339.avi"


   SetCMBValue "128", "29", "320x240", "8000", "12.2", "������"
  Case 1
'    MsgBox ("MP4")
    SetCMBValue "128", "29", "320x240", "8000", "12.2", "������"
  Case 2
'    MsgBox ("RMVB")
    SetCMBValue "1200", "29", "ԭʼ��С", "44100", "320", "˫����"
  Case 3
'    MsgBox ("RM")
    SetCMBValue "1200", "29", "ԭʼ��С", "44100", "320", "˫����"
'  Case 4
''    MsgBox ("MPG")
'    SetCMBValue "600", "24", "ԭʼ��С", "28000", "64", "˫����"
  Case 5 - 1
'    MsgBox ("AVI")
    SetCMBValue "600", "24", "ԭʼ��С", "28000", "64", "˫����"
  Case 6 - 1
'    MsgBox ("MP4")
    SetCMBValue "1200", "29", "ԭʼ��С", "48000", "320", "˫����"
  Case 7 - 1
'    MsgBox ("MOV")
    SetCMBValue "256", "29", "ԭʼ��С", "48000", "64", "˫����"
  Case 8 - 1
'    MsgBox ("MP3")
    SetCMBValue "256", "29", "ԭʼ��С", "48000", "320", "˫����"
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
     sLoc = BrowseForFolderByPath(Me.hWnd, "��ѡ���ת����Ƶ���ļ���", sAppPath)
     
     If sLoc <> "" Then
        i = FileList(sLoc, sFiles)
     End If
     
     For j = 0 To i - 1
       If InStr(sFiles(j), ".") > 0 Then
        lstFiles.AddItem (IIf(right(sLoc, 1) <> "\", sLoc & "\", sLoc) & sFiles(j))
       End If
     Next
     
     If lstFiles.ListCount > 0 Then
        SetStatus "��ת������Ƶ�ļ����������ת���б�"
     End If
     
        
'    lstFiles.ForeColor = &HFF&
'
'    lstFiles.List(0) = lstFiles.List(0) & " <ת���ɹ�>"
    
    
    

     
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

MsgBox " �����:  " & Err.Number & vbCrLf & _
       " ��������:" & Err.Description & vbCrLf & _
       " ����λ��:errCmdConvert", vbExclamation + vbOKOnly




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
'         MsgBox "ת���ɹ���", vbInformation, sInfo_TiTle_Msg
'Else
'         MsgBox "ת��ʧ��ඣ�������˼������һ�Σ�", vbCritical, sInfo_TiTle_Msg
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


    SetStatus "����ת����" & Filename(Trim(lstFiles.List(i)))
    
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
                
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\Girl.flv" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\Girl-224610.3gp"
                
                
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\͸��ʱװ�����˵��й�ʱװ��-23948.3gp"
                
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\ȫ��͸��ʱװ�� - ������__.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4

                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 128x96 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\ȫ��͸��ʱװ�� - ������__-15850.3gp"

                  '����һ���������ʱmp4�ļ���ȥ��
                  
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '���������ʽ�������Ӧ������ļ���ȥ��
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y " & IIf(cmbVideoSizeValue = "ԭʼ�ߴ�", "", " -s " & cmbVideoSizeValue) & " -vcodec mpeg4 -acodec amr_nb -ac 1 -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k """ & OutFile & """"
                  Execute sAppPath & sShell_Coder2 & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  
'                      cmbVideoBT.Text = cmbVideoBTValue
'    cmbVideoZL.Text = cmbVideoZLValue
'    cmbVideoSize.Text = cmbVideoSizeValue
'    cmbSoundBL.Text = cmbSoundBLValue
'    cmbSoundBT.Text = cmbSoundBTValue
'    cmbSoundSD.Text = cmbSoundSDValue

                
                 'SetCMBValue "128", "29", "320x240", "8000", "12.2", "������"
                Case 1
                
                
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\͸��ʱװ�����˵��й�ʱװ��-23948.3gp"
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder2.exe -i c:\tmp\tmp.mp4 -y -s 480x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k "c:\output\͸��ʱװ�����˵��й�ʱװ��-1231.mp4"
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '���������ʽ�������Ӧ������ļ���ȥ��
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "ԭʼ�ߴ�", "", " -s " & cmbVideoSizeValue) & " -vcodec mpeg4 -acodec amr_nb -ac 1 -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k """ & OutFile & """"
                  Execute sAppPath & sShell_Coder2 & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD


                '    MsgBox ("MP4")
                  'SetCMBValue "128", "29", "320x240", "8000", "12.2", "������"
                Case 2
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                 'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec rv10 -b 1000k -f rm -ar 24000 -ab 320k -ac ˫���� "c:\output\͸��ʱװ�����˵��й�ʱװ��-1212.rmvb"

                '    MsgBox ("RMVB")
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '���������ʽ�������Ӧ������ļ���ȥ��
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "ԭʼ�ߴ�", "", " -s " & cmbVideoSizeValue) & " -vcodec rv10 -b " & cmbVideoBTValue & "k -f rm -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  'SetCMBValue "1200", "29", "ԭʼ��С", "44100", "320", "˫����"
                Case 3
                
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -vcodec rv10 -b 1200k -f rm -ar 44100 -ab 320k -ac ˫���� "c:\output\͸��ʱװ�����˵��й�ʱװ��-115533.rm"
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '���������ʽ�������Ӧ������ļ���ȥ��
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "ԭʼ�ߴ�", "", " -s " & cmbVideoSizeValue) & " -vcodec rv10 -b " & cmbVideoBTValue & "k -f rm -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD


                '    MsgBox ("RM")
                  'SetCMBValue "1200", "29", "ԭʼ��С", "44100", "320", "˫����"
                '  Case 4
                ''    MsgBox ("MPG")
                '    SetCMBValue "600", "24", "ԭʼ��С", "28000", "64", "˫����"
'                Case 5
'                '    MsgBox ("AVI")
'                  SetCMBValue "600", "24", "ԭʼ��С", "28000", "64", "˫����"
                Case 6 - 2
                '    MsgBox ("MP4")
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o "c:\output\͸��ʱװ�����˵��й�ʱװ��-122232.mp4"
                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & """"
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD

                  'SetCMBValue "1200", "29", "ԭʼ��С", "48000", "320", "˫����"
                Case 7 - 2
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\conv.exe "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -ss 0 -endpos 30 -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o c:\tmp\tmp.mp4
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder.exe -i c:\tmp\tmp.mp4 -y -s 640x480 -ac 2 -ar 48000 -ab 128k "c:\output\͸��ʱװ�����˵��й�ʱװ��-123048.mov"
                
                                
                  Params = " """ & InFile & """ -oac mp3lame -lameopts preset=320 -ovc lavc -lavcopts vcodec=mpeg4:vbitrate=1200 -of avi -o  """ & OutFile & ".TMP.MP4"""
                  Execute sAppPath & sShell_Conv & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD
                  
                  '���������ʽ�������Ӧ������ļ���ȥ��
                  'Params = " -i """ & OutFile & ".TMP.MP4"" -y -s 640x480 -vcodec mpeg4 -acodec amr_nb -ac 1 -ar 8000 -ab 12.2k """ & OutFile & """"
                  
                  Params = " -i """ & OutFile & ".TMP.MP4"" -y  " & IIf(cmbVideoSizeValue = "ԭʼ�ߴ�", "", " -s " & cmbVideoSizeValue) & "  -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD

                '    MsgBox ("MOV")
                  'SetCMBValue "256", "29", "ԭʼ��С", "48000", "64", "˫����"
                Case 8 - 2
                '    MsgBox ("MP3")
                'PAnsiChar  lpCommandLine: C:\3GP��ʽת����\data\coder.exe -i "E:\MTV\͸��ʱװ�����˵��й�ʱװ��.mp4" -y -ss 0 -t 30 -y -vn -ar 48000 -ab 320k -ac 2 "c:\output\͸��ʱװ�����˵��й�ʱװ��-124145.mp3"
                
                  Params = " -i """ & InFile & """ -y -y -vn -ar " & cmbSoundBLValue & " -ab " & cmbSoundBTValue & "k -ac 1 """ & OutFile & """"
                  Execute sAppPath & sShell_Coder & " " & Params, , NORMAL_PRIORITY_CLASS, , sShellCD


                  'SetCMBValue "256", "29", "ԭʼ��С", "48000", "320", "˫����"
            End Select
        
        
        
        
        If FileExists(OutFile) Then StartConvert = True
        
    End If
End Function


Private Sub ViewControl(bViewSele As Integer)
On Error Resume Next

lblTitle3.ForeColor = &H80000012

    Select Case bViewSele
       Case 1
           lblTitle3.Caption = "�ļ�����"
           lstFiles.Visible = True
           WebBrowser.Visible = False
           txtAboutMe.Visible = False

       Case 2
         
'           If left(sIAgreeUrl, 10) <> "http://www.mama520.cn/" Then
         If sRegUser = sADUrl Then
                lblTitle3.Caption = "������ҳ"
                lblUnlikeURL.Visible = False
           Else
                If CheckRegister(sRegUser, sRegCode, False) And Not CheckContactURL(sRegUser, sRegCode) Then
                    lblTitle3.ForeColor = &HC00000
                    lblTitle3.Caption = "�� �������ݷǹٷ�������ҳ�ṩ����ʾ����Ϊ���Զ���ҳ�� ��"
                End If
                
                If CheckRegister(sRegUser, sRegCode, False) And CheckContactURL(sRegUser, sRegCode) Then
                    lblTitle3.ForeColor = 49152
                    lblTitle3.Caption = "�� �������ݷǹٷ�������ҳ�ṩ������ͨ���Ϸ�����֤ ��"
                End If
                
                
                If Not CheckRegister(sRegUser, sRegCode, False) Then
                    lblTitle3.ForeColor = &HFF&
                    lblTitle3.Caption = "�� ������ַ��Ȩ����������ݷǷ���������Ҫ�ٱ��� ��"
                End If
                
                
           End If
           
           
           
           
           
           lstFiles.Visible = False
           WebBrowser.Visible = True
           txtAboutMe.Visible = False
       Case Else
           lblTitle3.Caption = "�������"
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

     'txtLoc.Text = BrowseForFolder(Me, "��ѡ�����·��")
     txtLoc.Text = BrowseForFolderByPath(Me.hWnd, "��ѡ�����·��", sAppPath)
    
End Sub

Private Sub cmdAddFile_Click()
    'txtPDFFile.Text = BrowseForFolderByPath(Me.Hwnd, "��ѡ��PDF�ļ�", sAppPath)
    
    
     Dim iRet As Long
     Dim sFileName As String
     
     
     ViewControl (1)
     
     iRet = DLG_Open(Me.hWnd, "*.*|*.*", vbNullString, "*", 0, "��ѡ��", True, sFileName, 0)
     
     If sFileName <> "" Then
     
        lstFiles.AddItem (sFileName)
     
     End If
     
     
     
     
       
     
     sSeleFileName = Filename(sFileName)
     
     If lstFiles.ListCount > 0 Then
        SetStatus "��ת������Ƶ�ļ����������ת���б�"
     End If
     
     
  
     
   '  sSeleFileName = ShortPath("C:\Documents and Settings\Administrator\����\���簲ȫ")
     
   '  sSeleFileName = ShortPath("C:\Documents and Settings\Administrator\����\���簲ȫ")
     
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
'            .Name = "����"
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
    
  '/////////----------------------��ȡ����Ȩ�����ĵ�--------------------------
  
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
  
  '/////////----------------------��ȡWEB����̨�����ĵ�--------------------------
  
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
    
    txtAboutMe.Text = Space(i) & "������ƣ�" & frmMain.Caption & vbCrLf & _
                    Space(i) & "�������ߣ�������" & vbCrLf & _
                    Space(i) & "��Ѷ΢����http://t.qq.com/huangqiqing   (�ǵó�Ϊ�ҵķ�˿)" & vbCrLf & _
                    Space(i) & "����΢����http://t.sina.com.cn/happyqq  (�ǵó�Ϊ�ҵķ�˿)" & vbCrLf & _
                    Space(i) & "�������ڣ�2011-2-28" & vbCrLf & _
                    Space(i) & "�ٷ���վ��http://www.mama520.cn" & vbCrLf & _
                    Space(i) & "�ʼ���ַ��mama520.cn@gmail.com" & vbCrLf & _
                    Space(i) & "�ر��л��Դ��ĿFFMEPG,MEncoder����Ƶ/��Ƶת����Ŀ�Ŀ������ǣ�"

    
ViewControl (3)
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    If pThread Is Nothing Then
        Unload Me
        End
    End If
    
    If pThread.IsThreadRunning = True Then
        MsgBox "����ִ�в��������Ե�Ƭ�̣�����ת����ɣ�^_^"
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

'sADs = "========================================������������ҾͿ��Խ���ඡ�=================================|||http://www.kaixin.com$$$�����������Ƕ�һ���棡|||http://www.renren.com$$$����΢��������߶�ľ����㣡|||http://www.mama520.cn"

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

    '��webdoc���õ���WebBrowser��Document���Է��ص��ĵ�������
    Set webdoc = WebBrowser.Document



End Sub

Private Function webdoc_oncontextmenu() As Boolean
webdoc_oncontextmenu = False
'Me.PopupMenu myMenu
End Function

