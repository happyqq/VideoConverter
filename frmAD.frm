VERSION 5.00
Begin VB.Form frmAD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定制个性迷你首页"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5310
   Icon            =   "frmAD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdADInfo 
      Appearance      =   0  'Flat
      Caption         =   "点击此处查看合作协作"
      Height          =   330
      Left            =   1755
      TabIndex        =   9
      Top             =   1665
      Width           =   3060
   End
   Begin VB.TextBox txtADUrl 
      Height          =   375
      Left            =   1755
      TabIndex        =   8
      Text            =   "http://YourName_URL.com/Ad.html"
      ToolTipText     =   "请输入迷你首页的地址"
      Top             =   1260
      Width           =   3030
   End
   Begin VB.CheckBox chkAgree 
      Caption         =   "同意此合作协议"
      Height          =   285
      Left            =   1755
      TabIndex        =   6
      Top             =   2070
      Width           =   1860
   End
   Begin VB.TextBox txtADRegisterCode 
      Height          =   375
      Left            =   1755
      TabIndex        =   5
      Text            =   "FREECODE"
      ToolTipText     =   "请输入官方授权码"
      Top             =   720
      Width           =   3030
   End
   Begin VB.TextBox txtADUser 
      Height          =   375
      Left            =   1755
      TabIndex        =   3
      Text            =   "YourName_URL.com"
      ToolTipText     =   "个性名称随便你取"
      Top             =   180
      Width           =   3030
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "返回"
      Height          =   375
      Left            =   3555
      TabIndex        =   1
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "定制生成"
      Height          =   375
      Left            =   1710
      TabIndex        =   0
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "我的个性网址"
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   1305
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "授权码"
      Height          =   195
      Left            =   450
      TabIndex        =   4
      Top             =   810
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "个性名称"
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   270
      Width           =   960
   End
End
Attribute VB_Name = "frmAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private clsInI As New clsInI
Private sAppPath As String
 
Private Sub cmdADInfo_Click()

OpenLink sIAgreeUrl

End Sub

Private Sub cmdCancle_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()

If chkAgree.Value = 0 Then
    MsgBox "合作协议，你必须先同意才行，勾选一下”同意此合作协议“吧！", vbInformation + vbOKOnly, "提醒"
    Exit Sub
End If


If Not CheckRegister(txtADUrl.Text, txtADRegisterCode.Text, True) Then Exit Sub


        clsInI.SaveKey txtADUser.Text, cNickName_Key, cRegister_Section
        clsInI.SaveKey txtADUrl.Text, cRegUser_Key, cRegister_Section
        clsInI.SaveKey txtADRegisterCode.Text, cRegCode_Key, cRegister_Section



 
End Sub

Private Sub Form_Load()

    If right(App.Path, 1) <> "\" Then
        sAppPath = App.Path + "\"
    Else
        sAppPath = App.Path
    End If
    
    clsInI.Filename = sAppPath & cRegister_File
    
'    Dim isRegister As Boolean
'
'    isRegister = False
    
    
'    If LCase(Trim(txtRegUser.Text)) = "freeuser" And LCase(Trim(txtRegCode.Text)) = "freeregistercode" Then
'
'        isRegister = False
'
'    Else
'        '验证注册用户名以及密码
'          If txtRegUser.Text = "HappyQQ" And txtRegCode.Text = "HQQDESN" Then
'          '我的保留用户
'            isRegister = True
'
'          End If
'
'          If IsRegisterUser Then
'            isRegister = True
'          End If
'
'
'
'    End If
    
    
'    If Not isRegister Then
'
'          chkAll.Value = 0
'
'          If Val(txtLastPage.Text) - Val(txtPrePage.Text) > 50 Then
'            txtLastPage.Text = Val(txtPrePage.Text) + 50
'          End If
'
'    End If
    
End Sub
