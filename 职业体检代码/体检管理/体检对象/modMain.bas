Attribute VB_Name = "modMain"
Public Const P_SUBSYSNAME = "体检管理"   '子系统名称。

'读配置文件。
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) As Long
'些配置文件。
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpString As Any, _
                ByVal lpFileName As String) As Long


Public pobjDict As Object         '字典对象。

Public Sub Main()
    On Error Resume Next
    '创建字典对象。
    Set pobjDict = CreateObject("字典管理.clsDictionary")
    If Err <> 0 Then
        '创建自定义的字典对象。
        Set pobjDict = New clsDictionary
    End If
    
End Sub


Public Function func错误处理(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case -2147217833
        func错误处理 = "输入数据过长（或过大），已超过系统规定长度（或大小）。"
    Case -2147217913
        func错误处理 = "数据格式不符合数据库要求（比如：日期格式非法；或则系统要求数值类型，而你输入字符类型）！"
    Case -2147217873
        func错误处理 = "系统无法继续处理，因为：" & Chr(13) & Chr(10) & "该次处理相关的信息已被人删除！"
    Case 6
        func错误处理 = "输入数据过大，已超过系统规定大小。"
    Case 94 '无效使用Null。
        func错误处理 = "使用的字典项被人通过字典管理操作删除了，系统无法再继续正常处理。请找系统管理员恢复字典内容。请注意，不要随便删除字典项！"
    Case 336, 337, 338, 429, 430
        func错误处理 = "系统部件已损坏（或已丢失），系统无法再正常运行。请退出系统，并重新安装系统。"
    Case 440 '外部对象错误：类自动错误。
        func错误处理 = "系统部件不正常终止运行。请退出系统，再重新启动系统。"
    Case 91 '对象没有初始化成功。
        func错误处理 = "因为网络阻塞，系统启动功能时无法完成正常的初始化。请退出功能界面，再重新进入功能界面。"
    Case 5
        func错误处理 = "因为网络中断（或阻塞），系统无法正常运行。"
    Case 482
        func错误处理 = "打印失败，因为找不到打印机。请检查打印机是否正常。"
    Case Else
        func错误处理 = paraErrDes
    End Select
End Function

