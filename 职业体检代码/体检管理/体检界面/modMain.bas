Attribute VB_Name = "modMain"
Option Explicit

'增加二代证读写模块，库函数定义 2012-08-28 徐冀川
Public Declare Function Syn_GetCOMBaud Lib "Syn_IDCardRead.dll" (ByVal iComID As Long, ByRef puiBaud As Byte) As Long
Public Declare Function Syn_SetCOMBaud Lib "Syn_IDCardRead.dll" (ByVal iComID As Long, ByVal uiCurrBaud As Long, ByVal uiSetBaud As Long) As Long

Public Declare Function Syn_OpenPort Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long) As Long
Public Declare Function Syn_ClosePort Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long) As Long

Public Declare Function Syn_GetSAMStatus Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long, ByVal iIfOpen As Long) As Long
Public Declare Function Syn_ResetSAM Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long, ByVal iIfOpen As Long) As Long

Public Declare Function Syn_GetSAMID Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long, ByRef pucSAMID As Byte, ByVal iIfOpen As Long) As Long
Public Declare Function Syn_GetSAMIDToStr Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long, ByRef pcSAMID As Byte, ByVal iIfOpen As Long) As Long

Public Declare Function Syn_StartFindIDCard Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long, ByRef pucManaInfo As Byte, ByVal iIfOpen As Long) As Long
Public Declare Function Syn_SelectIDCard Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long, ByRef pucManaMsg As Byte, ByVal iIfOpen As Long) As Long
Public Declare Function Syn_ReadMsg Lib "Syn_IDCardRead.dll" (ByVal iPortID As Long, ByVal iIfOpen As Long, ByRef pIDCardData As Byte) As Long

Public Declare Function Syn_SendSound Lib "Syn_IDCardRead.dll" (ByVal iCmdNo As Long) As Long
Public Declare Function Syn_DelPhotoFile Lib "Syn_IDCardRead.dll" ()
Public Declare Function Syn_SetPhotoPath Lib "Syn_IDCardRead.dll" (ByVal iOption As Long, ByRef cPhotoPath As Byte) As Long
Public pobj业务对象 As Object '体检管理业务对象clsManageMedicalExam。

Public Sub Main()
    On Error GoTo errHandler
    
    '创建业务对象。
    Set pobj业务对象 = CreateObject("体检对象.clsManageMedicalExam")
        
    Exit Sub
errHandler:
End Sub

'功能：填录入板内容。
'作者：杨春。
Public Sub sub填录入板值(ByVal para录入板 As Control, _
                        ByVal paraGUI As cls界面通用对象, _
                        ByVal paraInfo As Collection)
    Dim lstrItem As String
    Dim lstrItemText  As String
    Dim i As Integer
    Dim lint卫生种类 As Integer
    Dim lint行业类别 As Integer
    Dim j As Integer
    
    On Error GoTo errHandler
    
    
    para录入板.pblnTemp = True
    lint卫生种类 = 0
    
    For i = 1 To para录入板.InfoCollection.Count
        '录入项目名称。
        lstrItem = para录入板.InfoCollection(i).Title
        
        If sffunc判断集合键值是否存在(paraInfo, lstrItem) Then
            '设置TrueText。
            para录入板.ItemTrueText(i - 1) = paraInfo(lstrItem)("项目值编号")
            '设置Text。
            para录入板.ItemText(i - 1) = paraInfo(para录入板.InfoCollection(i).Title)("项目值")
            
            If lstrItem = "卫生种类" Then
                lint卫生种类 = i
            ElseIf lstrItem = "行业类别" Then
                lint行业类别 = i
            End If
        Else
            para录入板.ItemTrueText(i - 1) = ""
            para录入板.ItemText(i - 1) = ""
        End If
    Next i
    
    Dim lobjRec As Object
    Dim lstrItemTrueText As String
    '设置行业类别录入框的字典内容的条件。
    If lint卫生种类 > 0 And lint行业类别 > 0 Then

        '获取卫生种类编号。
        lstrItemTrueText = para录入板.ItemTrueText(lint卫生种类 - 1)

        '设置行业类别录入框的字典。
        If lstrItemTrueText <> "" And Not para录入板.InfoCollection(lint卫生种类).DictRecordSet Is Nothing Then
            Set lobjRec = para录入板.InfoCollection(lint卫生种类).DictRecordSet
            If Not lobjRec.EOF Then
                paraGUI.sub初始化字典表 lint行业类别, "Parent=" & lobjRec("InnerId")
            End If
        End If
    End If
  
    para录入板.pblnTemp = False
    Exit Sub
errHandler:
    para录入板.pblnTemp = False
    sfsub错误处理 "体检界面部件", "modMain", "sub填录入板值", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

Public Sub sub显示单位属性(ByVal ciptBase As Control, _
            ByVal para单位申请编号 As String, _
            ByVal paraGUI As cls界面通用对象)
    Dim i As Long
    Dim lcolInfo As Collection
    
    If para单位申请编号 <> "" Then
    
        '获取单位属性。
        On Error Resume Next
        '获取单位属性。
        Set lcolInfo = pobj业务对象.func获取单位属性(para单位申请编号)
        
        ciptBase.pblnTemp = True
        
        ciptBase.Box1("卫生种类").TrueText = ""
        ciptBase.Box1("行业类别").TrueText = ""
        ciptBase.Box1("片区").TrueText = ""
        ciptBase.Box1("经济性质").TrueText = ""
        
        ciptBase.Box1("卫生种类").TrueText = lcolInfo("卫生种类")
        ciptBase.Box1("行业类别").TrueText = lcolInfo("行业类别")
        ciptBase.Box1("片区").TrueText = lcolInfo("片区")
        ciptBase.Box1("经济性质").TrueText = lcolInfo("经济性质")
        
        ciptBase.Box1("卫生种类").Text = lcolInfo("卫生种类名称")
        ciptBase.Box1("行业类别").Text = lcolInfo("行业类别名称")
        ciptBase.Box1("片区").Text = lcolInfo("片区名称")
        ciptBase.Box1("经济性质").Text = lcolInfo("经济性质名称")
        ciptBase.Box1("单位地址").Text = lcolInfo("单位地址")
        
        
        
        Dim lstrItem As String
        Dim lint卫生种类 As Integer
        Dim lint行业类别  As Integer
        
        Err.Clear
        
        '判断是否有卫生种类。
        For i = 1 To ciptBase.InfoCollection.Count
            '录入项目名称。
            lstrItem = ciptBase.InfoCollection(i).Title
            
            If lstrItem = "卫生种类" Then
                lint卫生种类 = i
            ElseIf lstrItem = "行业类别" Then
                lint行业类别 = i
            End If
            If Err <> 0 Then Exit For
        Next i
        
        '设置行业类别录入框的字典内容的条件。
        Dim lstrItemTrueText As String
        Dim lobjRec As Object
        If lint卫生种类 > 0 And lint行业类别 > 0 Then
            '获取卫生种类编号。
            lstrItemTrueText = ciptBase.ItemTrueText(lint卫生种类 - 1)
            '设置行业类别录入框的字典。
            If lstrItemTrueText <> "" And Not ciptBase.InfoCollection(lint卫生种类).DictRecordSet Is Nothing Then
                Set lobjRec = ciptBase.InfoCollection(lint卫生种类).DictRecordSet
                If Not lobjRec.EOF Then
                    paraGUI.sub初始化字典表 lint行业类别, "Parent=" & lobjRec("InnerId")
                End If
            End If
        End If
        
        ciptBase.pblnTemp = False
    End If

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


