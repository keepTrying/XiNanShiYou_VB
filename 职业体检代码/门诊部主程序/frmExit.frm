VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00E4E8C6&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "退出系统"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4200
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4200
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton copt选项 
      BackColor       =   &H00E0E4BE&
      Caption         =   "&2注销后，以其它身份登录"
      Height          =   375
      Index           =   1
      Left            =   975
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.OptionButton copt选项 
      BackColor       =   &H00E0E4BE&
      Caption         =   "&1退出本次操作并关闭程序"
      Height          =   375
      Index           =   0
      Left            =   975
      TabIndex        =   0
      Top             =   225
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.Image ccmdOk 
      Height          =   300
      Left            =   585
      Picture         =   "frmExit.frx":27A2
      Top             =   1305
      Width           =   945
   End
   Begin VB.Image ccmdCancel 
      Height          =   300
      Left            =   2460
      Picture         =   "frmExit.frx":4FB0
      Top             =   1305
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmExit.frx":7769
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ccmdCancel_Click()
    pblnCancel = True
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    On Error GoTo errHandle
    pblnCancel = False
    If copt选项(0).Value = True Then
        pblnExit = True
    Else
        pblnExit = False
    End If
    Unload Me
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frmExit", "ccmdOk_Click", Err.Number, Err.Description, False)
End Sub

