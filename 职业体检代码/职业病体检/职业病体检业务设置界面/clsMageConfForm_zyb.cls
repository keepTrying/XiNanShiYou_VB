VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMageConfForm_zyb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'功能：启动本部检所包括的一级操作界面，
'      按照总体结构的要求，必须无模式启动界面，并指定其parent指向主导航界面，并把界面移动到指定位置。
'输入：para操作名称      要启动的界面对应的操作名称。
'返回：操作对应的界面句柄。
Public Function funcStart(ByVal para操作名称 As String) As Long
    Dim lfrmForm As Form
    On Error GoTo errHandler
    Select Case para操作名称
    Case "体检表设置", "职业病体检_体检表设置"
        Set lfrmForm = frmSetMedicalExamTemplate
        
    Case "业务设置", "职业病体检_业务设置"
        Set lfrmForm = frmConfigureMedicalExam
    
    Case Else
        funcStart = -1
        Err.Raise 6666, , "当前体检管理版本不支持“" & para操作名称 & "”操作。"
    End Select
    '如果找到窗体则打开窗体并调整窗体位置，返回窗体句柄。
    If funcStart <> -1 Then
        If lfrmForm.pblnInUse Then
            funcStart = -2
        Else
            funcStart = lfrmForm.hWnd
        End If
        lfrmForm.Move 700, 350 '1560, 0
'        lfrmForm.ZOrder 0
        lfrmForm.Show
        
        Set lfrmForm = Nothing
    End If
    Exit Function
errHandler:
    Set lfrmForm = Nothing
    sfsub错误处理 "职业病设置界面", "clsmageconfform_zyb", "funcStart", Err.numer, Err.Description, True
    Exit Function
    Resume
End Function

' 功能：    将操作名所对应的窗体关闭，所有界面模板必须提供此方法。
' 输入：    para操作名
' 返回：    是否成功关闭窗体。
Public Function funcClose(para操作名 As String) As Boolean
    On Error GoTo errHandler
    Dim lfrmForm As Form '待关闭的窗体
    
    funcClose = False
    '根据操作名设置不同的待关闭窗体。
    Select Case para操作名
        Case "体检表设置", "职业病体检_体检表设置"
            Set lfrmForm = frmSetMedicalExamTemplate
            
        Case "业务设置", "职业病体检_业务设置"
            Set lfrmForm = frmConfigureMedicalExam
        Case Else
            funcClose = True
    End Select
    '如果找到窗体则关闭。
    If funcClose = False Then
        Unload lfrmForm
        '如果窗体未确认退出则返回False。
        funcClose = Not lfrmForm.pblnInUse
    End If
    Set lfrmForm = Nothing
    Exit Function
    
errHandler:
    Set lfrmForm = Nothing
    sfsub错误处理 "职业病设置界面", "clsManageManageConfigureForm", "funcClose", Err.Number, Err.Description, True
End Function

'功能：为主导航界面获取主推信息。
Public Function Func获取主推信息(ByVal para操作名称 As String) As String
    Func获取主推信息 = ""
End Function
