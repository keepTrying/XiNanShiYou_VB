VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据接口"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdImport 
      Caption         =   "导入结果(&I)"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton ccmdExport 
      Caption         =   "导出名单(&O)"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ccmdExport_Click()
    frmOutputData.Show
End Sub

Private Sub ccmdImport_Click()
    frmInputData.Show
End Sub
