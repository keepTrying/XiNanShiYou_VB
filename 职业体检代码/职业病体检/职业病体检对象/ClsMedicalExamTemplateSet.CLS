VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMedicalExamTemplateSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************
'名称：职业体检表设置
'函数：
'功能：体检表选定和照射类型选定是由前一组合框绑定
'       另写组合框填充函数
'作者：Yunle Liu
'时间：2012.03
'*********************************************

Option Explicit
Private mint体检表类型 As String
Private mint体检表类别 As String
Private mint照射源 As String

Public Property Get 体检表类型() As String
    体检表类型 = mint体检表类型
End Property

Public Property Let 体检表类型(ByVal vNewValue As String)
    'If vNewValue > 3 Or vNewValue < 1 Then
    On Error GoTo errorHandle
    'End If
    mint体检表类型 = vNewValue
    Exit Property
errorHandle:
    Exit Property
End Property

'体检表类别：涉核部队，放射工作人员等
Public Property Let 体检表类别(ByVal vNewValue As String)
    mint体检表类别 = vNewValue
    Exit Property
errorHandle:
    Exit Property
End Property
'体检表类别：涉核部队，放射工作人员等
Public Property Get 体检表类别() As String
    体检表类别 = mint体检表类别
    Exit Property
End Property
'功能：返回指定体检表类型的体检表名称集合。
'作者：刘浩
Public Property Get 元素集() As Collection
    Dim lstrSQL语句 As String
    Dim lrst查询返回值 As Object
    Dim lcol数据集 As New Collection
    On Error GoTo errHandler
    If mint体检表类别 = "" Or mint体检表类型 = "" Then Exit Property
    'MsgBox "german1", , "german1" 'german
    Select Case mint体检表类型
        Case "1"
            lstrSQL语句 = "select 体检表名称 from 职业病体检_体检表模板基本信息表 order by 体检表名称 desc"
   '     Case 2
            '复查体检表
   '         lstrSQL语句 = "select 体检表名称 from 职业病体检_体检表模板基本信息表 where isnull(是否复查体检表,0) = 1  order by 代号"
        Case Else
            '不是复查体检表
            lstrSQL语句 = "select 编号,体检表名称 from 职业病体检_体检表模板基本信息表 where isnull(是否复查体检表,0) <> 1 and 体检类别='" & mint体检表类别 & "' and 体检人员类型='" & mint体检表类型 & "' order by 代号"
    End Select
    dasubSetQueryTimeout 6000
    Set lrst查询返回值 = dafuncGetData(lstrSQL语句)
    While Not lrst查询返回值.EOF
        lcol数据集.Add lrst查询返回值("体检表名称").Value, lrst查询返回值("体检表名称").Value
        'lcol数据集.Add lrst查询返回值("编号").Value, lrst查询返回值("编号").Value
        lrst查询返回值.MoveNext
    Wend
    Set 元素集 = lcol数据集
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Property Get 元素集", 6666, lstrError, False

End Property

'功能：返回指定体检表类型的体检表名称集合。
'作者：刘浩
Public Property Get 体检表编号元素集() As Collection
    Dim lstrSQL语句 As String
    Dim lrst查询返回值 As Object
    Dim lcol数据集 As New Collection
    On Error GoTo errHandler
    If mint体检表类型 = "" Or mint体检表类别 = "" Then Exit Property
    Select Case mint体检表类型
        Case "1"
             lstrSQL语句 = "select 编号,体检表名称 from 职业病体检_体检表模板基本信息表 order by 代号"
    '    Case 2
    '        '复查体检表
    '        lstrSQL语句 = "select 体检表名称 from 职业病体检_体检表模板基本信息表 where isnull(是否复查体检表,0) = 1  order by 代号"
        Case Else
            '不是复查体检表
            lstrSQL语句 = "select 编号,体检表名称 from 职业病体检_体检表模板基本信息表 where isnull(是否复查体检表,0) <> 1 and 体检类别='" & mint体检表类别 & "' and 体检人员类型='" & mint体检表类型 & "' order by 代号"
    End Select
    dasubSetQueryTimeout 6000
    Set lrst查询返回值 = dafuncGetData(lstrSQL语句)
    While Not lrst查询返回值.EOF
        'lcol数据集.Add lrst查询返回值("体检表名称").Value, lrst查询返回值("体检表名称").Value
        lcol数据集.Add lrst查询返回值("编号").Value, lrst查询返回值("编号").Value
        lrst查询返回值.MoveNext
    Wend
    Set 体检表编号元素集 = lcol数据集
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Property Get 元素集", 6666, lstrError, False

End Property
'功能：返回指定体检表类型的体检表信息记录。
'返回：recordset[体检表名称，代号，体检单名称，是否复查体检表，诊断处理意见，收费标准]。
'作者：杨春
Public Property Get 详细元素集() As Object
    Dim lstrSQL语句 As String
    Dim lrst查询返回值 As Object
    Dim lcol数据集 As New Collection
    
    On Error GoTo errHandler
    'Select Case mint体检表类型
    '    Case 1
    '        lstrSQL语句 = "select * from 职业病体检_体检表模板基本信息表"
    '    Case 2
    '        '复查体检表
    '        lstrSQL语句 = "select * from 职业病体检_体检表模板基本信息表 where 是否复查体检表 = 1"
    '    Case 3
            '不是复查体检表
            lstrSQL语句 = "select * from 职业病体检_体检表模板基本信息表 where 是否复查体检表 <> 1"
    'End Select
    dasubSetQueryTimeout 6000
    Set 详细元素集 = dafuncGetData(lstrSQL语句)

    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Property Get 详细元素集", 6666, lstrError, False
End Property

Private Sub Class_Initialize()
    mint体检表类型 = "1"
    mint体检表类别 = "1"
End Sub

Private Sub Class_Terminate()
'
End Sub


'获取照射源字典
Public Property Get 照射源() As Object
    Dim lstrSQL语句 As String
    Dim lstrTmp As String
    On Error GoTo errHandler
    dasubSetQueryTimeout 6000
    '将a.编号=2改成4，取得下级节点 2015-04-15 刘伟
    lstrSQL语句 = "select a.编号,a.名称 from 系统管理_字典_字典内容表 a,系统管理_字典_字典表列表 b where a.ID=b.ID and b.名称='职业照射种类字典' and len(a.编号)=4"
    Set 照射源 = dafuncGetData(lstrSQL语句)
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Get 照射源", Err.Number, Err.Description, False
End Property

'获取职业类别字典
Public Property Get 职业类别() As Object
    Dim lstrSQL语句 As String
    Dim lstrTmp As String
    On Error GoTo errHandler
    dasubSetQueryTimeout 6000
    lstrSQL语句 = "select a.编号,a.名称 from 系统管理_字典_字典内容表 a,系统管理_字典_字典表列表 b where b.名称='职业照射种类字典' and a.ID=b.ID and len(a.编号)=4 and left(a.编号,2)=" & mint照射源
    Set 职业类别 = dafuncGetData(lstrSQL语句)
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Get 职业类别", Err.Number, Err.Description, False
End Property

'放射源
Public Property Let int照射源(ByVal vNewValue As Integer)
    mint照射源 = "" & vNewValue
    Exit Property
errorHandle:
    Exit Property
End Property


'获取行业类别字典
Public Property Get 行业类别() As Object
    Dim lstrSQL语句 As String
    Dim lstrTmp As String
    On Error GoTo errHandler
    dasubSetQueryTimeout 6000
    lstrSQL语句 = "select a.编号,a.名称 from 系统管理_字典_字典内容表 a,系统管理_字典_字典表列表 b where a.ID=b.ID and b.名称='行业属性字典' and len(a.编号)=1"
    Set 行业类别 = dafuncGetData(lstrSQL语句)
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Get 行业类别", Err.Number, Err.Description, False
End Property
