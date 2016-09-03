VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "主要完成以下功能:\r\n对卫生许可证界面的管理;"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B315DFF0367"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

' 功能：    卸载指定操作界面。
Public Function funcClose(ByVal para操作名 As String) As Boolean
    On Error GoTo errHandler
    Dim lfrmForm As Form '待关闭的窗体
    funcClose = False
    '根据操作名设置不同的待关闭窗体。
    Select Case para操作名
        Case "健康证管理_健康证管理"
            Set lfrmForm = frm健康证管理
        Case "健康证管理_业务设置"
            Set lfrmForm = frm业务设置
        Case Else
            funcClose = True
    End Select
    '如果找到窗体则关闭。
    If funcClose = False Then
        Unload lfrmForm
        '如果窗体未确认退出则返回False。
        funcClose = Not lfrmForm.pblnInUse
    End If
    Exit Function
errHandler:
    sfsub错误处理 "健康证管理部件", "clsManageForm", "funcClose", Err.Number, Err.Description, True
End Function

'功能：启动本部检所包括的一级操作界面，
Public Function funcStart(ByVal para操作名 As String) As Long
    On Error GoTo errHandler
    Dim lfrmForm As Form '待打开的窗体
    funcStart = 0
    
    Select Case para操作名
        Case "健康证管理_健康证管理"
            Set lfrmForm = frm健康证管理
        Case "健康证管理_业务设置"
            Set lfrmForm = frm业务设置
        Case Else
            '如果未找到相应界面则设置返回值为-1。
            funcStart = -1
    End Select
    '如果找到窗体则打开窗体并调整窗体位置，返回窗体句柄。
    If funcStart <> -1 Then
        If lfrmForm.pblnInUse Then
            funcStart = -2
        Else
            funcStart = lfrmForm.hWnd
        End If
        lfrmForm.Move 700, 350 '1560, 0
        'lfrmForm.ZOrder 0
        lfrmForm.Show
    End If
    Exit Function
errHandler:
    sfsub错误处理 "健康证管理部件", "clsManageForm", "funcStart", Err.Number, Err.Description, True
End Function

Public Function funcCheck(paraName As String) As Variant
    funcCheck = "ab!&d3290"
End Function

