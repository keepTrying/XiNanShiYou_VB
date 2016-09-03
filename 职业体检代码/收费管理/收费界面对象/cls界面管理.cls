VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls界面管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'修改：2002-6-23（杨春）增加参数“para参数”，供修改划价使用。
Public Function funcStart(para操作名 As String) As Long

    On Error GoTo errhandler

    Dim lfrmForm As Form '待打开的窗体

    funcStart = 0

    '根据操作名设置不同的待打开窗体。

    Select Case para操作名
        Case "收费管理_划价"
            Set lfrmForm = frm划价
            
        Case "收费管理_业务设置"
            Set lfrmForm = frm业务设置
        Case "收费管理_收费管理"
            Set lfrmForm = frm收费管理
        Case "收费管理_号段设置"
            Set lfrmForm = frm号段设置
        Case "收费管理_财务监管"
            Set lfrmForm = frm财务监管
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

        lfrmForm.Show

    End If
    
errhandler:

    Set lfrmForm = Nothing

    If Err.Number = 0 Then Exit Function

    sfsub错误处理 "界面模板", "cls界面模板", "funcStart", Err.Number, Err.Description, True
    Exit Function
    Resume
End Function

Public Function func获取主推信息(ByVal para操作名 As String) As String
    On Error GoTo errhandler
        func获取主推信息 = "无"
    Exit Function
errhandler:
    sfsub错误处理 "收费界面", "cls界面管理", "func获取主推信息", Err.Number, Err.Description, True
End Function

Private Sub Class_Initialize()
    On Error GoTo errhandler
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面", "cls界面管理", "Class_Initialize", Err.Number, Err.Description, False
End Sub

Public Function funcClose(para操作名 As String) As Boolean
    On Error GoTo errhandler
    Dim lfrmForm As Form '待关闭的窗体
    funcClose = False
    '根据操作名设置不同的待关闭窗体。
    Select Case para操作名

        Case "收费管理_划价"

            Set lfrmForm = frm划价

            
        Case "收费管理_业务设置"

            Set lfrmForm = frm业务设置
            
        Case "收费管理_收费管理"

            Set lfrmForm = frm收费管理
            
        Case "收费管理_号段设置"
            Set lfrmForm = frm号段设置
        Case "收费管理_财务监管"
            Set lfrmForm = frm财务监管
            
        Case Else
        
            funcClose = True
            
    End Select
    
    '如果找到窗体则关闭。
    If funcClose = False Then
        
        Unload lfrmForm
        
        '如果窗体未确认退出则返回False。
        funcClose = Not lfrmForm.pblnInUse
    
    End If
errhandler:
    Set lfrmForm = Nothing
    If Err.Number = 0 Then Exit Function
    sfsub错误处理 "收费界面", "cls界面管理", "funcClose", Err.Number, Err.Description, True
End Function


'功能：启动划价界面。
'输入：要修改的收费编号,单位申请编号。
'返回：划价的金额;收费编号。
'创建：2002-6-23（杨春）。
Public Function func划价(Optional ByVal para收费编号 As String = "", _
                        Optional ByVal para申请编号 As String = "", _
                        Optional ByVal para业务分类 As String = "", _
                        Optional ByVal paraTop As Long = 350)
    '修改：2002-6-23（杨春）可以修改划价。
    frm划价.pstr收费编号 = Trim(para收费编号)
    frm划价.pstr单位编号 = Trim(para申请编号)
    frm划价.pstr业务分类 = Trim(para业务分类)
    frm划价.Move 1560, paraTop
    frm划价.Show vbModal
    
    '返回收费编号。
    func划价 = frm划价.pstr收费编号
End Function

