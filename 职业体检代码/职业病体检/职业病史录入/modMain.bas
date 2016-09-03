Attribute VB_Name = "modMain"
'***************************************
'名称：职业病史(受检者个人信息)管理工程启动
'函数：func错误处理
'功能：工程启动，错误处理
'作者：Yunle Liu
'时间：2012.03
'***************************************



Option Explicit

Public pobjDict As Object     'clsDictionary。

Public 访问记号 As Integer   '区分职业病史录入和修改
Public pub系统编号 As String
Public pobj业务对象 As Object '体检管理业务对象clsManageMedicalExam。
Public bolenProject As Boolean  '是否已确定体检项目

Public Sub Main()
    On Error Resume Next
     '创建字典对象。
    Set pobjDict = CreateObject("字典管理.clsDictionary")
    Err.Clear
    '创建业务对象。
    Set pobj业务对象 = CreateObject("职业病对象.clsManageMedicalExam")
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "无法创建疫苗管理数据对象。请重新注册“职业病史录入.dll”。"
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "modmain", "Main", Err.Number, Err.Description, False
End Sub

Public Function func错误处理(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case 6
        func错误处理 = "输入数据过大，已超过系统规定大小。"
    Case -2147217833
        func错误处理 = "输入数据过长（或过大），已超过系统规定长度（或大小）。"
    Case -2147217913
        func错误处理 = "日期格式非法！"
    Case -2147217873 '外键不存在。
        func错误处理 = "系统服务继续处理。因为：" & Chr(13) & Chr(10) & "(1) 你正在保存的数据涉及的相关信息已被人删除！" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出本业务操作界面，重新进入。"
    Case 94 '无效使用Null。
        func错误处理 = "使用的字典项被人通过字典管理操作删除了，系统无法再继续正常处理。请找系统管理员恢复字典内容。请注意，不要随便删除字典项！"
    Case 336, 337, 338, 429, 430
        func错误处理 = "系统部件已损坏（或已丢失），系统无法再正常运行。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，并重新安装系统。"
    Case 440 '外部对象错误：类自动错误。
        func错误处理 = "系统部件不正常终止运行。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，重新进入。"
    Case 91 '对象没有初始化成功。
        func错误处理 = "因为网络阻塞，系统启动功能时无法完成正常的初始化。请退出功能界面，再重新进入功能界面。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，重新进入。"
    Case 5
        func错误处理 = "因为网络中断（或阻塞），系统无法正常运行。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，重新进入。"
    Case Else
        func错误处理 = paraErrDes
    End Select
End Function


