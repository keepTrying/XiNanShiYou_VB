VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim objLocalConfigure As ClsLocalConfigure
    
    Set objLocalConfigure = New ClsLocalConfigure
    
    objLocalConfigure.工作模式 = 1
    objLocalConfigure.Excel文件 = "c:\"
    objLocalConfigure.内部导入文件 = "c:\体检管理"
    objLocalConfigure.内部导出文件 = "c:\hee"
End Sub
