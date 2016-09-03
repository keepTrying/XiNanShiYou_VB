VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDataServerConfig 
   Caption         =   "数据服务器配置"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10260
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退出"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton ccmdSave 
      Caption         =   "保存"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTDept 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8281
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "X光影像科"
      TabPicture(0)   =   "frmDataServerConfig.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraAccept"
      Tab(0).Control(1)=   "fraSend"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "B超影像科"
      TabPicture(1)   =   "frmDataServerConfig.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "血常规等~（其它待定）"
      TabPicture(2)   =   "frmDataServerConfig.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.Frame fraAccept 
         Caption         =   "接收端"
         Height          =   2295
         Left            =   -70200
         TabIndex        =   5
         Top             =   1920
         Width           =   4335
      End
      Begin VB.Frame fraSend 
         Caption         =   "发送端"
         Height          =   2295
         Left            =   -74640
         TabIndex        =   4
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "P.S.本来应该做个 存储图片服务器的设置界面，但当前测试阶段，只是用来设置图片的默认文件夹的。"
         Height          =   255
         Left            =   -74640
         TabIndex        =   3
         Top             =   600
         Width           =   8295
      End
   End
End
Attribute VB_Name = "frmDataServerConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
