VERSION 5.00
Begin VB.Form frm业务设置 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "业务设置"
   ClientHeight    =   6420
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   8880
   Icon            =   "frm业务设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu cmnuItemOther 
      Caption         =   "收费项目(&I)"
      Index           =   1
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "收费标准(&T)"
      Index           =   2
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "票据格式(&F)"
      Index           =   3
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "打折(&D)"
      Index           =   4
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "开户银行(&B)"
      Index           =   5
   End
   Begin VB.Menu cmnuBase 
      Caption         =   "退出系统"
   End
End
Attribute VB_Name = "frm业务设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Dim mlngID As Long      '当前修改的号段的ID

Private Sub ccmdSet_Click(Index As Integer)
    On Error GoTo errhandler
    Select Case Index
    Case 0, 1, 2
        cmnuItemOther_Click Index + 1
    Case 3
        cmnuItemOther_Click 5
    Case 4
        frm设置科室比例.Show 1, Me
        
    End Select
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm业务设置", "ccmdSet_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub cmnuBase_Click()
    Unload Me
End Sub

Private Sub cmnuItemOther_Click(Index As Integer)
    On Error GoTo errhandler
    Select Case Index
    Case 1
        frm设置收费项目.Move Me.Left, Me.Top
        frm设置收费项目.Show 1
    Case 2 '收费标准
        frm设置收费标准.Move Me.Left, Me.Top
        frm设置收费标准.Show 1, Me
    Case 3 '票据格式
        frm设置票据格式.Move Me.Left, Me.Top
        frm设置票据格式.Show 1, Me
    Case 4 '打折
        frm设置打折.Move Me.Left, Me.Top
        frm设置打折.Show 1, Me
    Case 5
        frm开户行设置.Move Me.Left, Me.Top
        frm开户行设置.Show 1, Me
        
    End Select
    
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm业务设置", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub Form_Load()
        
    If pblnInUse Then Exit Sub
    pblnInUse = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub
