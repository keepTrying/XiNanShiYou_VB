VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls主程序"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'功能：捕获窗体消息的窗口函数（回调函数）。
'注意事项：请勿在主程序中设定断点调试程序。
'作者：罗庆
'创建时间：2001-4-17
Public Function funcClassing(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Static lblnTerminate As Boolean
    '判断消息类型，并进行处理。
    Dim lstrCaption As String
    If sffunc判断集合键值是否存在(pcolWndProc, CStr(hWnd)) Then
        funcClassing = CallWindowProc(pcolWndProc(CStr(hWnd)), hWnd, Msg, wParam, lParam)
    End If
    '如果窗体发生了Destroy消息则从集合中移除该窗体。
    If Msg = WM_DESTROY Then
        If hWnd = plngMainHwnd Then lblnTerminate = True
        SetWindowText plngMainHwnd, pstrSysName  '& "（试用版）"
        If sffunc判断集合键值是否存在(pcolWndProc, CStr(hWnd)) Then
            SetWindowLong hWnd, GWL_WNDPROC, pcolWndProc(CStr(hWnd))
            pcolWndProc.Remove CStr(hWnd)
        End If
        If sffunc判断集合键值是否存在(pcol操作名称, CStr(hWnd)) Then
            Call oesubSave("用户退出" & pcol操作名称(CStr(hWnd)), "退出操作")
            If sffunc判断集合键值是否存在(pcol子窗体句柄, pcol操作名称(CStr(hWnd))) Then
                pcol子窗体句柄.Remove pcol操作名称(CStr(hWnd))
            End If
            pcol操作名称.Remove CStr(hWnd)
        End If
        If hWnd <> plngMainHwnd Then
            
            SetForegroundWindow plngMainHwnd
        End If
    Else
        '如果窗体发生Activate消息则设定其对应的子系统名。
        If Msg = WM_ACTIVATE Then
            If hWnd = plngMainHwnd Then lblnTerminate = False
            If sffunc判断集合键值是否存在(pcol业务名称, CStr(hWnd)) Then
                um当前操作子系统名 = pcol业务名称(CStr(hWnd))
            End If
        End If
        '重新设置主窗体标题栏。
        If Not lblnTerminate Then subSetCaption
    End If
End Function

' 功能：设定主窗体标题栏。
' 输入：无
' 输出：无
' 返回：无
' 注意事项：无
' 作者：罗庆
' 创建时间：2001-4-17
Private Sub subSetCaption()
    On Error Resume Next
    Dim llngHwnd As Long
    Dim lstrCaption As String
    Dim lstrTemp As String
    llngHwnd = GetForegroundWindow
    If llngHwnd = plngMainHwnd Then
        lstrCaption = pstrSysName  '& "（试用版）"
        
    Else
        If sffunc判断集合键值是否存在(pcol操作名称, CStr(llngHwnd)) Then
            'lstrCaption = "道源" & pstrSysName & "（试用版）" & "\" & pcol操作名称(CStr(llngHwnd))
            lstrCaption = pstrSysName & "\" & pcol操作名称(CStr(llngHwnd))
        End If
    End If
    If lstrCaption = "" Then
        lstrCaption = pstrSysName  '& "（试用版）"
    ElseIf llngHwnd <> plngMainHwnd Then
        Unload frm操作列表
        Unload frm字典列表
    End If
    If Err.Number = 0 Then
        If pstrMainCaption <> lstrCaption Then
            pstrMainCaption = lstrCaption
            SetWindowText plngMainHwnd, lstrCaption
        End If
    End If
    Err.Clear
End Sub
