VERSION 5.00
Begin VB.Form frmAD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ƹ���������ҳ"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdADInfo 
      Appearance      =   0  'Flat
      Caption         =   "����˴��鿴����Э��"
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
      ToolTipText     =   "������������ҳ�ĵ�ַ"
      Top             =   1260
      Width           =   3030
   End
   Begin VB.CheckBox chkAgree 
      Caption         =   "ͬ��˺���Э��"
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
      ToolTipText     =   "������ٷ���Ȩ��"
      Top             =   720
      Width           =   3030
   End
   Begin VB.TextBox txtADUser 
      Height          =   375
      Left            =   1755
      TabIndex        =   3
      Text            =   "YourName_URL.com"
      ToolTipText     =   "�������������ȡ"
      Top             =   180
      Width           =   3030
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "����"
      Height          =   375
      Left            =   3555
      TabIndex        =   1
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��������"
      Height          =   375
      Left            =   1710
      TabIndex        =   0
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "�ҵĸ�����ַ"
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   1305
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "��Ȩ��"
      Height          =   195
      Left            =   450
      TabIndex        =   4
      Top             =   810
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
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
    MsgBox "����Э�飬�������ͬ����У���ѡһ�¡�ͬ��˺���Э�顰�ɣ�", vbInformation + vbOKOnly, "����"
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
'        '��֤ע���û����Լ�����
'          If txtRegUser.Text = "HappyQQ" And txtRegCode.Text = "HQQDESN" Then
'          '�ҵı����û�
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
