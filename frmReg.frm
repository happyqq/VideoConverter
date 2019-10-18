VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Èí¼þ×¢²á"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5310
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdReturn 
      Caption         =   "·µ»Ø"
      Height          =   435
      Left            =   4095
      TabIndex        =   6
      Top             =   1530
      Width           =   1020
   End
   Begin VB.TextBox txtRegCode 
      BackColor       =   &H8000000E&
      Height          =   315
      Left            =   1305
      TabIndex        =   2
      Text            =   "FreeRegisterCode"
      Top             =   855
      Width           =   3840
   End
   Begin VB.TextBox txtRegUser 
      BackColor       =   &H8000000E&
      Height          =   315
      Left            =   1305
      TabIndex        =   1
      Text            =   "FreeUser"
      Top             =   225
      Width           =   3840
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Èí¼þ×¢²á"
      Height          =   435
      Left            =   2835
      TabIndex        =   0
      Top             =   1530
      Width           =   1020
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   225
      OleObjectBlob   =   "frmReg.frx":0000
      TabIndex        =   3
      Top             =   270
      Width           =   990
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   225
      OleObjectBlob   =   "frmReg.frx":005D
      TabIndex        =   4
      Top             =   945
      Width           =   945
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   615
      Left            =   315
      OleObjectBlob   =   "frmReg.frx":00B8
      TabIndex        =   5
      Top             =   2205
      Width           =   4995
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
