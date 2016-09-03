VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTestItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mstr编码 As String

Public 名称 As String
Public 缺省值 As String
Public 枚举来源 As String
Public 属性 As String
Public 体检大类 As Long
Public 标准值 As String
Public 比较方式 As String
Public 单位 As String
Public 单价 As String 'german

Private mbln是否存在 As Boolean

Public Property Get 是否存在() As Boolean
    是否存在 = mbln是否存在
End Property

Public Property Get 编码() As String
    编码 = mstr编码
End Property

'功能：给本对象关键属性赋值，从而获取其它属性。
'作者：杨春
Public Property Let 编码(ByVal vNewValue As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    mstr编码 = vNewValue
    mbln是否存在 = False
    
    '从“体检管理_体检项目设置表”中获取当前编码对应的体检项目其他属性。
    'german
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    dasubSetQueryTimeout 600
    Set lobjRec = dafuncGetData("select * from 职业病体检_体检项目设置表 where 编码='" & mstr编码 & "'")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If lobjRec.RecordCount > 0 Then
        名称 = lobjRec!名称
        缺省值 = IIf(IsNull(lobjRec!缺省值), "", lobjRec!缺省值)
        枚举来源 = IIf(IsNull(lobjRec!枚举来源), "", lobjRec!枚举来源)
        属性 = IIf(IsNull(lobjRec!属性), "常规", lobjRec!属性)
        体检大类 = IIf(IsNull(lobjRec!体检大类), 0, lobjRec!体检大类)
        比较方式 = IIf(IsNull(lobjRec!比较方式), "", lobjRec!比较方式)
        标准值 = IIf(IsNull(lobjRec!标准值), "", lobjRec!标准值)
        单位 = IIf(IsNull(lobjRec!单位), "", lobjRec!单位)
        
        mbln是否存在 = True
    Else
        名称 = ""
        缺省值 = ""
        枚举来源 = ""
        属性 = "常规"
        单位 = ""
        标准值 = ""
        比较方式 = ""
        
        体检大类 = 0
    End If
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsTestItem", "Property Let 编码", Err.Number, Err.Description, False
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    mbln是否存在 = False
    属性 = "常规"
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "ClsTestItem", "Class_Initialize", Err.Number, Err.Description, False
End Sub

'功能   仅仅修改单价
'作者：何啸天
'时间 ：3.13
Public Sub SubSaveUnitprice()
    'SOL 语句模板 = update 职业病体检_体检项目设置表 set 单价=2222 where(编码='1333')
    On Error GoTo errHandler
    If (单价 = "") Then
        MsgBox "单独修改单价的时候出错，单价没有被填写", 16, "消息"
        Exit Sub
    End If
    dafuncGetData ("update 职业病体检_体检项目设置表 set 单价=" + 单价 + " where(编码='" + mstr编码 + "')")
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "ClsTestItem", "subSave", Err.Number, Err.Description, True
End Sub
'功能：保存体检项目设置信息。
'作者：杨春




Public Sub SubSave()
    Dim lobjRec As Object
    On Error GoTo errHandler
    '判断名称是否重复。
    'german
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    dasubSetQueryTimeout 600
'    Set lobjRec = dafuncGetData("select count(*) from 职业病体检_体检项目设置表 where 名称='" & 名称 & "'" & IIf(mbln是否存在, " and 编码<>'" & mstr编码 & "'", ""))
    '不同的科室有相同的项目时也可以添加或修改（只要在同一科室没有相同的就行） 2016-3-2 by 牟俊
    Dim lob As Object
    Set lob = dafuncGetData("select 编码 from 职业病体检_体检项目设置表 where 名称='" & 名称 & "'" & IIf(mbln是否存在, " and 编码<>'" & mstr编码 & "'", ""))
    If lob.RecordCount > 0 Then
        Dim i As Integer
        Dim a As String
        For i = 1 To lob.RecordCount
            a = lob("编码")
            If Left(a, 2) = Left(mstr编码, 2) Then
                Err.Raise 6666, , "体检项目的名称不允许重复。"
            End If
        lob.MoveNext
        Next
'        Set lobjRec = dafuncGetData("select count(*) from 职业病体检_体检项目设置表 where 名称='" & 名称 & "'" & IIf(mbln是否存在, " and 编码<>'" & mstr编码 & "' and left('" & lob("编码") & "',2)=left('" & mstr编码 & "',2)", ""))
        Set lobjRec = dafuncGetData("select count(*) from 职业病体检_体检项目设置表 where 名称='" & 名称 & "' and '" & Left(a, 2) & "'='" & Left(mstr编码, 2) & "'")
    Else
        Set lobjRec = dafuncGetData("select count(*) from 职业病体检_体检项目设置表 where 名称='" & 名称 & "'" & IIf(mbln是否存在, " and 编码<>'" & mstr编码 & "'", ""))
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    If IIf(IsNull(lobjRec(0)), 0, lobjRec(0)) > 0 Then
'        Err.Raise 6666, , "体检项目的名称不允许重复。"
'    End If
    
    If mbln是否存在 Then
        '修改。
        'german
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        dafuncGetData "update 职业病体检_体检项目设置表 set 名称='" & 名称 & "',缺省值='" & 缺省值 & "',枚举来源='" & 枚举来源 & "',属性='" & IIf(属性 = "", "常规", 属性) & "',体检大类=" & 体检大类 & ",比较方式='" & 比较方式 & "',标准值='" & 标准值 & "',单位='" & 单位 & "' where 编码='" & mstr编码 & "'"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        '新增。
        'german
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        dafuncGetData "insert into 职业病体检_体检项目设置表(编码,名称,缺省值,枚举来源,属性,体检大类,比较方式,标准值,单位,单价) values('" & mstr编码 & "','" & 名称 & "','" & 缺省值 & "','" & 枚举来源 & "','" & IIf(属性 = "", "常规", 属性) & "'," & 体检大类 & ",'" & 比较方式 & "','" & 标准值 & "','" & 单位 & "'," & 单价 & ")"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '修改：2012.3.1 何啸天
        On Error Resume Next
        'dafuncGetData "insert into 职业病体检_体检医师项目设置表(医师编号,体检项目) " _
                & "select 编号,'" & mstr编码 & "' from 系统管理_系统用户信息表 where 编号<>'0000' and 编号<>'gues' and 编号 not in(select 医师编号 from 职业病体检_体检医师项目设置表 where 体检项目='" & mstr编码 & "')"
         dafuncGetData "insert into 职业病体检_体检医师项目设置表(医师编号,体检项目) " _
                & "select 编号,'" & mstr编码 & "' from 系统管理_系统用户信息表 where 编号<>'0000' and 编号<>'gues' and 编号 not in(select 医师编号 from 职业病体检_体检医师项目设置表 where 体检项目='" & mstr编码 & "')"
        
        mbln是否存在 = True
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "ClsTestItem", "subSave", Err.Number, Err.Description, True
End Sub

'功能：批量删除数据
'作者：何啸天
'修改功能： 增加用户可强制删除信息的功能
Public Sub subDelete_any() 'german
    'MsgBox a.Rows
    On Error GoTo errHandler
    dasubBeginTran '事务开始
    dafuncGetData "delete 职业病体检_体检医师项目设置表 where 体检项目='" & mstr编码 & "'"
    dafuncGetData "delete 职业病体检_体检项目设置表 where 编码='" & mstr编码 & "'"
    dafuncGetData "delete from 职业病体检_体检表模板体检项目表 where 体检项目 = '" & mstr编码 & "'"
    dasubCommitTran '事务结束调用
errHandler:
    Dim llngErr As Long    '错误号。
    Dim llngError As String '错误信息。
    llngErr = Err.Number
    llngError = Err.Description
    If llngErr = P_ERRORCODE_VIOLATE Then
        llngErr = 6666
        llngError = "你删除的体检项目已被使用，不允许删除。"
    End If
    sfsub错误处理 "体检对象部件", "ClsTestItem", "subDelete", llngErr, llngError, True
End Sub

'功能：删除体检项目设置信息。
'作者：何啸天
'修改功能： 增加用户可强制删除信息的功能
Public Sub subDelete() 'german
    On Error GoTo errHandler
    dasubBeginTran '事务开始
    '先删除医师权限设置。
    'dafuncGetData "delete 职业病体检_体检医师项目设置表 where 体检项目='" & mstr编码 & "'" '旧
    'dafuncGetData "delete 体检管理_体检项目设置表 where 编码='" & mstr编码 & "'" '旧
    'german
    '数据库改动：职业病体检_体检医师项目设置表 改为 职业病体检_体检医师项目设置表
    '试图 体检管理_体检医师项目设置数据库 中表名修改 系统管理_清库工具 中表名修改
    '////////////////////////////////////////////////////////
    
    dafuncGetData "delete 职业病体检_体检医师项目设置表 where 体检项目='" & mstr编码 & "'"
    dafuncGetData "delete 职业病体检_体检项目设置表 where 编码='" & mstr编码 & "'"
    dafuncGetData "delete from 职业病体检_体检表模板体检项目表 where 体检项目 = '" & mstr编码 & "'"
    
'    If (flag = False) Then
'        '如果没有强制删除，那么就跳过
'        Exit Sub
'    Else '用户执行强制删除
'        'MsgBox "强制删除", , "消息"
'        dafuncGetData "delete 职业病体检_体检医师项目设置表 where 体检项目='" & mstr编码 & "'"
'        dafuncGetData "delete 职业病体检_体检项目设置表 where 编码='" & mstr编码 & "'"
'    End If
    dasubCommitTran '事务结束调用
    '////////////////////////////////////////////////////////
    
    mbln是否存在 = False
    
    Exit Sub
errHandler:
    Dim llngErr As Long    '错误号。
    Dim llngError As String '错误信息。
    llngErr = Err.Number
    llngError = Err.Description
    If llngErr = P_ERRORCODE_VIOLATE Then
        llngErr = 6666
        llngError = "你删除的体检项目已被使用，不允许删除。"
    End If
    sfsub错误处理 "体检对象部件", "ClsTestItem", "subDelete", llngErr, llngError, True
End Sub

