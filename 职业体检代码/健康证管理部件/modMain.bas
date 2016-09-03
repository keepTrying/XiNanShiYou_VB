Attribute VB_Name = "modMain"
Option Explicit
Public pobj体检管理 As cls体检管理
Public pobj记忆  As cls用户操作记忆

Sub Main()
    Set pobj体检管理 = New cls体检管理
    
    Set pobj记忆 = New cls用户操作记忆
    pobj记忆.用户编号 = "*"
    pobj记忆.业务名 = "健康证管理"
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
    Case 20504, 20507
        func错误处理 = "在系统路径下未找到要打印的文件模板。"
    Case 20510
        func错误处理 = "水晶报表变量传值错误！"
    Case 20526, 20545
        func错误处理 = "存在妨碍打印的问题。此错误产生的原因及解决方法如下： " & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (1) 没有从 Windows 控制面板中安装打印机？" & Chr(13) & Chr(10) _
                    & "      解决：打开控制面板，双击“打印机”图标，选择“添加打印机”以装入打印机。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (2) 打印机没在线？" & Chr(13) & Chr(10) _
                    & "      解决：检查打印机与计算机的连接是否正常。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (3)  打印机阻塞或缺纸？" & Chr(13) & Chr(10) _
                    & "      解决：解决这些问题。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (4) 试图在只能接受文本的打印机上打印窗体？" & Chr(13) & Chr(10) _
                    & "      解决：切换到一台能打印图形的打印机。"
        
    Case Else
        func错误处理 = paraErrDes
    End Select
End Function

