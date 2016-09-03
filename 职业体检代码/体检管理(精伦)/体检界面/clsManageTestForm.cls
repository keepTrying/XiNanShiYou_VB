VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManageTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'功能：启动本部件所包括的一级操作界面，
'      按照总体结构的要求，必须无模式启动界面，并指定其parent指向主导航界面，并把界面移动到指定位置。
'输入：para操作名称      要启动的界面对应的操作名称。
'返回：界面的句柄（若指定操作不存在，返回-1）。
'修改：2001-12-27（初检登记、年检登记合并为“体检登记”。
Public Function funcStart(ByVal para操作名称 As String) As Long
    Dim lfrmForm As Form '待打开的窗体
    
    On Error GoTo errHandler
    Select Case para操作名称
    Case "体检登记", "体检管理_体检登记"
        Set lfrmForm = frmRegisterManage
        
    Case "体检管理_体检结果录入"
        Set lfrmForm = FrmInputTestResult
    
    Case "体检管理_体检结论录入"
        Set lfrmForm = frmGiveConclusion
        
    Case "文书打印", "体检管理_文书打印"
        Set lfrmForm = frmPrintPaper
        
'    Case "体检公众查询", "体检管理_体检公众查询"
'        Set lfrmForm = frmPublicQueryResult
'        funcStart = lfrmForm.hWnd
'        lfrmForm.Show 1
'        Exit Function
    
    Case Else
        funcStart = -1
        Err.Raise 6666, , "当前体检管理版本不支持操作“" & para操作名称 & "”。"
    End Select
    '如果找到窗体则打开窗体并调整窗体位置，返回窗体句柄。
    If funcStart <> -1 Then
        If lfrmForm.pblnInUse Then
            funcStart = -2
        Else
            funcStart = lfrmForm.hWnd
        End If
'        lfrmForm.ZOrder 0
        lfrmForm.Show
        lfrmForm.Move 700, 350 '1560, 0
        
        Set lfrmForm = Nothing
    End If
    Exit Function

errHandler:
    Set lfrmForm = Nothing
    sfsub错误处理 "体检界面部件", "clsManageTestForm", "subStart", Err.Number, Err.Description, True
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
    Case "体检登记", "体检管理_体检登记"
        Set lfrmForm = frmRegisterManage
        
    Case "体检管理_体检结果录入"
        Set lfrmForm = FrmInputTestResult
    
    Case "体检管理_体检结论录入"
        Set lfrmForm = frmGiveConclusion
        
    Case "文书打印", "体检管理_文书打印"
        Set lfrmForm = frmPrintPaper
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
    sfsub错误处理 "体检界面部件", "clsManageTestForm", "funcClose", Err.Number, Err.Description, True
End Function


'功能：为主导航界面获取主推信息。
Public Function Func获取主推信息(ByVal para操作名称 As String) As String
    Dim lstr复查间隔 As String
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    Func获取主推信息 = ""
    If para操作名称 = "复检登记" Or para操作名称 = "体检管理_复检登记" Or para操作名称 = "体检管理_复查登记" Then
        '若是"复查登记"操作，则获取待复查人数。
        Set lobjRec = pobj业务对象.Func获取需要复查的体检记录
        Func获取主推信息 = "待复查人员有" & lobjRec.recordcount & "个。"
    Else
        '获取体检但未下结论的记录数。
        Set lobjRec = pobj业务对象.Func获取已下结论但未确定的体检记录("", "", "", "")
        Func获取主推信息 = "待下体检结论的人员有" & lobjRec.recordcount & "个。"
    End If
    
    
    Exit Function
errHandler:
    'sfsub错误处理 "体检界面部件", "clsManageTestForm", "Func获取主推信息", Err.Number, Err.Description, True
End Function

'功能：启动指定操作前，关闭其它窗体与摄像设备的连接。
'作者：杨春。
'时间：2001-5-12。
Private Sub sub关闭其它窗体与摄像设备的连接(ByVal para操作名 As String)
    Dim lobjForm1 As Form
    Dim lobjForm2 As Form
    
    On Error Resume Next
    
    '若业务设置不照相，不需要继续。
    If pobj业务对象.业务设置("是否照像") <> "是" Then Exit Sub
    
    '根据操作名设置不同的待关闭窗体。
    Select Case para操作名
'    Case "初检登记", "体检管理_初检登记"
'        Set lobjForm1 = FrmRegisterAnnual
'        Set lobjForm2 = frmRegisterLater
'    Case "年检登记", "体检管理_年检登记"
'        Set lobjForm1 = FrmRegisterFirst
'        Set lobjForm2 = frmRegisterLater
    
    Case "体检登记", "体检管理_体检登记"
        Set lobjForm1 = FrmRegister
        
    Case "补录体检登记信息", "体检管理_补录体检登记信息"
        Set lobjForm1 = FrmRegister
'        Set lobjForm1 = FrmRegisterFirst
'        Set lobjForm2 = FrmRegisterAnnual
    End Select
        
    '先关闭其他窗体的摄像设备连接。
    If Not lobjForm1 Is Nothing Then
        With lobjForm1
            If .pblnInUse Then
                .cctlCatchPhoto.subDisconnect
            End If
        End With
    End If
    If Not lobjForm2 Is Nothing Then
        With lobjForm2
            If .pblnInUse Then
                .cctlCatchPhoto.subDisconnect
            End If
        End With
    End If
End Sub

'功能：启动指定体检登记界面前，若该界面已启动，恢复其与摄像设备的连接。
'作者：杨春。
'时间：2001-5-12。
Private Sub sub恢复与照相设备的连接(ByVal paraForm As Form)
    On Error Resume Next
    
    '若业务设置不照相，不需要继续。
    If pobj业务对象.业务设置("是否照像") = "是" And Not paraForm Is Nothing Then
        With paraForm
            If .pblnInUse Then
                '重新初始化照相控件。
                .cctlCatchPhoto.funcInitVideo
                '判断初始化是否成功
                If Not .cctlCatchPhoto.VideoIsOk Then
                    .csbMain.Panels(1) = "照相设备初始化失败，请检查原因或到业务设置中设置不进行照相！"
                End If
            End If
        End With
    End If
End Sub
