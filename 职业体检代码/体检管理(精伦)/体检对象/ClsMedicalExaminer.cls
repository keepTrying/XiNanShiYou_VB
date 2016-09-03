VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMedicalExaminer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mstr编号 As String
Private mstr姓名  As String

'功能：给本对象的关键属姓“体检已是编号”赋值，从而通过创建用户对象获取医师姓名。
'作者：杨春
Public Property Let 编号(ByVal vNewValue As String)
    Dim obj用户 As Object
    
    On Error GoTo errHandler
    mstr编号 = vNewValue
    
'    '调试代码。
'    If mstr编号 = "" Then
'        mstr编号 = "0000"
'    End If
    
    '创建“用户”对象，获取姓名、科室。
    Set obj用户 = CreateObject("系统管理.cls用户")
    obj用户.用户编号 = mstr编号
    mstr姓名 = obj用户.员工姓名
    
    Exit Property
errHandler:
    Dim llngErr As Long
    Dim lstrError As String
    llngErr = Err.Number
    lstrError = Err.Description
    If llngErr = P_ERRORCODE_CREATEOBJECT_FAILED Then
        lstrError = "创建“系统管理.cls用户”的对象失败。请退出系统，并重新注册“系统管理.dll”。"
        llngErr = 6666
    End If
'    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Property Let 编号", llngErr, lstrError, True
    Exit Property
    
    Resume
End Property

Public Property Get 编号() As String
    编号 = mstr编号
End Property
Public Property Get 姓名() As String
    姓名 = mstr姓名
End Property

'功能：获取当前体检医师可做得体检项目。
'返回：[编码，名称]
'作者：杨春
Public Property Get 可作体检项目() As Collection
Attribute 可作体检项目.VB_Description = "'返回：[编码，名称]\r\n"
    Dim lobjRec As Object
    Dim lcolItem As Collection
    
    On Error GoTo errHandler
    Set 可作体检项目 = New Collection
    
    '从视图“体检管理_体检医师项目设置数据库”中获取该体检医师的姓名、科室、可作体检项目。
    Set lobjRec = dafuncGetData("select * from 体检管理_体检医师项目设置数据库 where 医师编号='" & mstr编号 & "'")
    Do While Not lobjRec.EOF
        Set colItem = New Collection
        colItem.Add lobjRec("体检项目").Value, "编码"
        colItem.Add lobjRec("体检项目名称").Value, "名称"
        
        可作体检项目.Add colItem, colItem("编码")
        lobjRec.MoveNext
    Loop
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Property Get 可作体检项目", 6666, lstrError, True
    Exit Property
    Resume
End Property

'功能：判断当前对象的医师是否可作指定体检项目。
'返回：true可以做/false不可以做。
'作者：杨春
Public Function func是否可作项目(ByVal para体检项目编码 As String) As Boolean
Attribute func是否可作项目.VB_Description = "'功能：判断当前对象的医师是否可作指定体检项目。\r\n    '返回：true可以做/false不可以做。\r\n"
    Dim lobjRec As Object
    On Error GoTo errHandler
    Set lobjRec = dafuncGetData("select * from 体检管理_体检医师项目设置表 where 医师编号='" & mstr编号 & "' and 体检项目='" & para体检项目编码 & "'")
    If lobjRec.recordcount = 0 Then
        func是否可作项目 = False
    Else
        func是否可作项目 = True
    End If
    
    Exit Function
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "func是否可作项目", Err.Number, Err.Description, True
End Function

'功能：添加当前医师可作的一个体检项目。
'作者：杨春
Public Sub Sub添加体检项目(ByVal para体检项目编码 As String)
    On Error GoTo errHandler

    dafuncGetData "insert into 体检管理_体检医师项目设置表(医师编号,体检项目) values('" & mstr编号 & "','" & para体检项目编码 & "')"
    
    Exit Sub
errHandler:
    Dim lstrError As String
    If Err.Number = -2147217873 Then
        lstrError = "无法保存体检医师设置，原因可能是：" & Chr(13) & Chr(10) & "(1) 该体检医师已被人通过“系统管理”的“员工管理”操作删除；" & Chr(13) & Chr(10) & "(2) 体检项目已被人删除。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 退出本业务界面，重新进入。"
    Else
        lstrError = func错误处理(Err.Number, Err.Description)
    End If
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Sub添加体检项目", 6666, lstrError, True
End Sub

'功能：删除当前医师可作的一个体检项目。
'作者：杨春
Public Sub Sub删除体检项目(ByVal para体检项目编码 As String)
    On Error GoTo errHandler
    dafuncGetData "delete 体检管理_体检医师项目设置表 where 医师编号='" & mstr编号 & "' and 体检项目='" & para体检项目编码 & "'"
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Sub删除体检项目", 6666, lstrError, True
End Sub

'功能：删除当前医师可作的所有体检项目。
'作者：杨春
Public Sub Sub删除所有体检项目()
    dafuncGetData "delete 体检管理_体检医师项目设置表 where 医师编号='" & mstr编号 & "'"
End Sub

'功能：获取当前医师在指定体检记录的体检表上可以作的体检项目。
'输入：para系统编号      体检的系统编号。
'      para体检项目属性  常规 / 化验。
'返回：Collection[clsFactTestItem],key:项目编码。
'          其中clsFactTestItem的属性：体检项目编码，体检项目名称，缺省值，枚举来源，体检结果，体检医师编号
'作者：杨春
Public Function Func获取本人指定体检表上可作的体检项目(ByVal para系统编号 As String, _
                                                       ByVal para体检项目属性 As String) As Collection
Attribute Func获取本人指定体检表上可作的体检项目.VB_Description = "'功能：获取当前医师在指定体检记录的体检表上可以作的体检项目。\r\n    '输入：para系统编号      体检的系统编号。\r\n    '      para体检项目属性  常规 / 化验。\r\n    '返回：Collection[clsFactTestItem],key:项目编码。\r\n    '          其中clsFactTestItem的属性：体"
    Dim lobjRec As Object               'RecordSet。
    Dim lobj体检 As clsMedicalExam
    Dim lcol所有体检项目 As Collection  '体检记录上所有的体检项目。
    Dim lcol返回体检项目 As Collection  '返回的本人可作的体检项目。
    Dim lobj体检项目 As Variant         'clsFactTestItem
    
    On Error GoTo errHandler
    
    '创建"体检"对象。
    Set lobj体检 = New clsMedicalExam
    lobj体检.系统编号 = para系统编号
    
    '获取"体检"对象的属性"体检表.体检项目集。
    Set lcol所有体检项目 = lobj体检.体检表.体检项目集(para体检项目属性)
    
    '找出属于本对象属性"可作体检项目"的体检项目及其已登记结果。
    Set lcol返回体检项目 = New Collection
    For Each lobj体检项目 In lcol所有体检项目
        '判断医师是否可作当前项目。
        Set lobjRec = dafuncGetData("select * from 体检管理_体检医师项目设置表 where 医师编号='" & mstr编号 & "' and 体检项目='" & lobj体检项目.体检项目编号 & "'")
        If lobjRec.recordcount > 0 Then
            '可以作，加入返回集合中。
            lcol返回体检项目.Add lobj体检项目, lobj体检项目.体检项目编号
        End If
    Next
    
    '按返回值要求返回。
    Set Func获取本人指定体检表上可作的体检项目 = lcol返回体检项目
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Func获取本人指定体检表上可作的体检项目", 6666, lstrError, True
End Function

'功能：获取当前医师在指定体检记录的体检表上可以作的体检项目。
'输入：para系统编号      体检的系统编号。
'      para体检项目属性  常规 / 化验。
'返回：Recordset(体检项目编码，体检项目名称，体检结果，缺省值，枚举来源)
'作者：杨春
'修改：2001-4-22
Public Function Func优化的获取本人可作的体检项目(ByVal para系统编号 As String, _
                                                       ByVal para体检项目属性 As String) As Object
    On Error GoTo errHandler
    
    '获取本人在指定体检表上可作的体检项目。
    Set Func优化的获取本人可作的体检项目 = dafuncGetData("select 体检项目编号=a.体检项目,b.体检项目名称,体检结果=case when isnull(体检结果,'')='' then 缺省值 else 体检结果 end,枚举来源,缺省值,标准值,单位,单项结论 from 体检管理_体检结果信息表 a,体检管理_体检医师项目设置数据库 b where 系统编号='" & para系统编号 & "' and 医师编号='" & um用户编号 & "' and a.体检项目=b.体检项目 and (属性='" & para体检项目属性 & "' or '" & para体检项目属性 & "'='') order by 属性,a.体检项目")
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Func优化的获取本人可作的体检项目", 6666, lstrError, True
End Function

'输入：para体检项目属性  常规 / 化验。
'返回：Recordset(体检项目编码，体检项目名称，体检结果，缺省值，枚举来源)
'创建：2002-10-14。
'作者：杨春
Public Function Func获取本人所有体检表上可作的体检项目(ByVal para体检项目属性 As String) As Object
    On Error GoTo errHandler
    
    Set Func获取本人所有体检表上可作的体检项目 = dafuncGetData("select distinct 体检项目编号=a.体检项目,b.体检项目名称,体检结果=缺省值,枚举来源,缺省值,属性,标准值,单位,单项结论='' from 体检管理_体检表模板体检项目表 a,体检管理_体检医师项目设置数据库 b where 医师编号='" & um用户编号 & "' and a.体检项目=b.体检项目 and (属性='" & para体检项目属性 & "' or '" & para体检项目属性 & "'='')  order by 属性,a.体检项目")
    
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Func获取本人所有体检表上可作的体检项目", 6666, lstrError, True
End Function

Public Function Func获取指定体检表上可作的体检项目(ByVal para体检表名 As String, ByVal para体检项目属性 As String) As Object
    On Error GoTo errHandler
    
    Set Func获取指定体检表上可作的体检项目 = dafuncGetData("select distinct 体检项目编号=a.体检项目,b.体检项目名称,体检结果=缺省值,枚举来源,缺省值,属性,标准值,单位,单项结论='' from 体检管理_体检表模板体检项目表 a,体检管理_体检医师项目设置数据库 b where 医师编号='" & um用户编号 & "' and a.体检项目=b.体检项目 and 体检表名称='" & para体检表名 & "' and (属性='" & para体检项目属性 & "' or '" & para体检项目属性 & "'='')  order by 属性,a.体检项目")
    
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExaminer", "Func获取指定体检表上可作的体检项目", 6666, lstrError, True
    Exit Function
    Resume
End Function
