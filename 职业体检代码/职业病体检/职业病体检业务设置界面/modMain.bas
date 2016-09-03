Attribute VB_Name = "modMain"
Option Explicit
'作者：杨春

Public pobj业务对象 As Object '体检管理业务对象clsManageMedicalExam。
Public pobjDict As Object     '字典对象clsDictionary。
Public flag_database_delete As Boolean '数据库强制删除标志 'german
Public flag_delete_any As Boolean '是否批量删除 'german

Public Sub Main()
    On Error Resume Next
    flag_database_delete = False '默认不强制删除数据库 'german，此标志在复选框的按钮事件中被修改
    flag_delete_any = False '默认不批量删除数据条目 'german，此标志在复选框的按钮事件中被修改
    '创建业务对象。
    Set pobj业务对象 = CreateObject("职业病对象.clsManageMedicalExam")
        
    Err.Clear
    
    '创建字典对象。
    Set pobjDict = CreateObject("字典管理.clsDictionary")
    If Err <> 0 Then
        '字典管理还不可用，创建自己定义的字典管理对象。
        Set pobjDict = New clsDictionary
    End If
    
End Sub

Public Function func错误处理(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case -2147217833
        func错误处理 = "输入数据过长（或过大），已超过系统规定长度（或大小）。"
    Case 6
        func错误处理 = "输入数据过大，已超过系统规定大小。"
    Case -2147217913
        func错误处理 = "日期格式非法！"
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
    Case Else
        func错误处理 = paraErrDes
    End Select
End Function
