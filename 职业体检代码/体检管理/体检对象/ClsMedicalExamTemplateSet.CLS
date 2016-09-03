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
Option Explicit
Private mint体检表类型 As Integer

Public Property Get 体检表类型() As Integer
    体检表类型 = mint体检表类型
End Property

Public Property Let 体检表类型(ByVal vNewValue As Integer)
    If vNewValue > 3 Or vNewValue < 1 Then
        GoTo errorHandle
    End If
    mint体检表类型 = vNewValue
    Exit Property
errorHandle:
    Exit Property
End Property

'功能：返回指定体检表类型的体检表名称集合。
'作者：刘浩
Public Property Get 元素集() As Collection
    Dim lstrSQL语句 As String
    Dim lrst查询返回值 As Object
    Dim lcol数据集 As New Collection
    On Error GoTo errHandler
    
    Select Case mint体检表类型
        Case 1
            lstrSQL语句 = "select 体检表名称 from 体检管理_体检表模板基本信息表 order by 代号"
        Case 2
            '复查体检表
            lstrSQL语句 = "select 体检表名称 from 体检管理_体检表模板基本信息表 where isnull(是否复查体检表,0) = 1  order by 代号"
        Case 3
            '不是复查体检表
            lstrSQL语句 = "select 体检表名称 from 体检管理_体检表模板基本信息表 where isnull(是否复查体检表,0) <> 1  order by 代号"
    End Select
    Set lrst查询返回值 = dafuncGetData(lstrSQL语句)
    While Not lrst查询返回值.EOF
        lcol数据集.Add lrst查询返回值("体检表名称").Value, lrst查询返回值("体检表名称").Value
        lrst查询返回值.MoveNext
    Wend
    Set 元素集 = lcol数据集
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Property Get 元素集", 6666, lstrError, True

End Property

'功能：返回指定体检表类型的体检表信息记录。
'返回：recordset[体检表名称，代号，体检单名称，是否复查体检表，诊断处理意见，收费标准]。
'作者：杨春
Public Property Get 详细元素集() As Object
    Dim lstrSQL语句 As String
    Dim lrst查询返回值 As Object
    Dim lcol数据集 As New Collection
    
    On Error GoTo errHandler
    Select Case mint体检表类型
        Case 1
            lstrSQL语句 = "select * from 体检管理_体检表模板基本信息表"
        Case 2
            '复查体检表
            lstrSQL语句 = "select * from 体检管理_体检表模板基本信息表 where 是否复查体检表 = 1"
        Case 3
            '不是复查体检表
            lstrSQL语句 = "select * from 体检管理_体检表模板基本信息表 where 是否复查体检表 <> 1"
    End Select
    
    Set 详细元素集 = dafuncGetData(lstrSQL语句)

    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Property Get 详细元素集", 6666, lstrError, True
End Property

Private Sub Class_Initialize()
    mint体检表类型 = 1
End Sub

Private Sub Class_Terminate()
'
End Sub
