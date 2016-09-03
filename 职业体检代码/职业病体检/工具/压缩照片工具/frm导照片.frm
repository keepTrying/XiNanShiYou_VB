VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm导照片 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "压缩体检健康证照片"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退出(&X)"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton ccmdExport 
      Caption         =   "开始压缩(&S)"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin ComctlLib.StatusBar Cstau状态栏 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2925
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7409
            MinWidth        =   7409
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm导照片"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdExport_Click()
    Dim lobj数据服务 As Object
    On Error GoTo errHandler
    MousePointer = 11
    ccmdExit.Enabled = False
    ccmdExport.Enabled = False
    Cstau状态栏.Panels.Item(1).Text = "正在压缩体检健康证图片..."
    
    Set lobj数据服务 = New Cls数据服务
    lobj数据服务.sub服务进程

    Cstau状态栏.Panels.Item(1).Text = "照片压缩完毕！"
    MousePointer = 0
    ccmdExit.Enabled = True
    ccmdExport.Enabled = True
    Exit Sub
errHandler:
    MsgBox Error, vbOKOnly + vbExclamation, "系统提示"
End Sub

