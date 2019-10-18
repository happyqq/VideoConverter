VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "提醒设置"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton OptClose 
      Caption         =   "关闭声音"
      Height          =   240
      Left            =   1350
      TabIndex        =   3
      Top             =   900
      Width           =   1140
   End
   Begin VB.OptionButton OptOpen 
      Caption         =   "开启声音"
      Height          =   240
      Left            =   1350
      TabIndex        =   2
      Top             =   495
      Value           =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回"
      Height          =   390
      Left            =   2385
      TabIndex        =   1
      Top             =   1260
      Width           =   885
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "保存设置"
      Height          =   390
      Left            =   1305
      TabIndex        =   0
      Top             =   1260
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "转换完毕后声音提示"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   135
      Width           =   2040
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private clsInI As New clsInI
Private sAppPath As String

Private Sub cmdReg_Click()
'        clsInI.SaveKey txtRegUser.Text, cRegUser_Key, cRegister_Section
'        clsInI.SaveKey txtRegCode.Text, cRegCode_Key, cRegister_Section
        
    '    PlaySound "D:\Product\3GP格式转换精灵\sound.wav"
    
    
    If OptOpen.Value = True Then
    
        clsInI.SaveKey "1", cSoundOpen_Key, cRegister_Section
        frmMain.iSoundOpen = 1
        
    Else
    
        clsInI.SaveKey "0", cSoundOpen_Key, cRegister_Section
        frmMain.iSoundOpen = 0
        
    End If
    
    
    MsgBox "保存设置成功！", vbInformation + vbOKOnly, "提醒"
    
    
    
        
End Sub

Private Sub cmdReturn_Click()
Me.Hide
End Sub

Private Sub Command1_Click()
frmMain.Show
End Sub

Private Sub Form_Load()
'  SkinH_Attach
    If right(App.Path, 1) <> "\" Then
        sAppPath = App.Path + "\"
    Else
        sAppPath = App.Path
    End If
    
    clsInI.Filename = sAppPath & cRegister_File
    
End Sub
