VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFactTestItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'本对象的主键属性。
Private mstr系统编号 As String
Private mstr体检项目编号 As String

'只读属性。
Private mstr体检项目名称 As String
Private mstr属性 As String          '常规/化验。
Private mstr缺省值 As String
Private mstr枚举来源 As String
Private mstr医师姓名 As String
Private mlng体检大类 As Long        '体检大类在字典表中的InnerID。

'可写属性。
Public 体检结果 As String
Public 医师编号 As String
Public 填写日期 As Date

Private mbln是否已存在 As Boolean

'功能：设置“系统编号”属性，获取其他属性。
'作者：杨春
Public Property Get 系统编号() As String
    系统编号 = mstr系统编号
End Property
Public Property Let 系统编号(ByVal paraSysNo As String)
    On Error GoTo errHandler
    mstr系统编号 = paraSysNo
    If mstr系统编号 <> "" And mstr体检项目编号 <> "" Then
        Call subGetProperty
    Else
        subClear
        mbln是否已存在 = False
    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsFactTestItem", "Property Let 系统编号", 6666, lstrError, True
End Property

'功能：设置“体检项目编号”属性，获取其他属性。
'作者：杨春
Public Property Get 体检项目编号() As String
    体检项目编号 = mstr体检项目编号
End Property
Public Property Let 体检项目编号(ByVal paraItemNo As String)
    On Error GoTo errHandler
    mstr体检项目编号 = paraItemNo
    If mstr系统编号 <> "" And mstr体检项目编号 <> "" Then
        Call subGetProperty
    Else
        subClear
        mbln是否已存在 = False
    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsFactTestItem", "Property Let 体检项目编号", 6666, lstrError, True
End Property

'功能：获取只读属性。
Public Property Get 体检项目名称() As String
    体检项目名称 = mstr体检项目名称
End Property
Public Property Get 属性() As String
    属性 = mstr属性
End Property
Public Property Get 缺省值() As String
    缺省值 = mstr缺省值
End Property
Public Property Get 枚举来源() As String
    枚举来源 = mstr枚举来源
End Property
Public Property Get 医师姓名() As String
    医师姓名 = mstr医师姓名
End Property
Public Property Get 体检大类() As String
    体检大类 = mlng体检大类
End Property

Public Property Get 是否已存在() As Boolean
    是否已存在 = mbln是否已存在
End Property
Public Property Let 是否已存在(ByVal vNewValue As Boolean)
    mbln是否已存在 = vNewValue
End Property
'功能：根据主键属性“系统编号、体检项目编号”获取指定体检表上指定体检项目的信息。
'作者：杨春
Private Sub subGetProperty()
    On Error GoTo errHandler
    Dim lobjRec As Object
    Dim lstrSql As String
    
    mbln是否已存在 = False
    
    '获得体检结果、体检医师、填写日期。
    lstrSql = "select * from 体检管理_体检结果数据库 where 系统编号='" & mstr系统编号 & "' and 体检项目='" & mstr体检项目编号 & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    If lobjRec.recordcount > 0 Then
        体检结果 = IIf(IsNull(lobjRec("体检结果")), "", lobjRec("体检结果"))
        填写日期 = Format(IIf(IsNull(lobjRec("填写日期")), Date, lobjRec("填写日期")), "yyyy-mm-dd")
        医师编号 = IIf(IsNull(lobjRec("体检医师")), "", lobjRec("体检医师"))
        
        mstr体检项目名称 = lobjRec("体检项目名称")
        mstr缺省值 = IIf(IsNull(lobjRec("缺省值")), "", lobjRec("缺省值"))
        mstr枚举来源 = IIf(IsNull(lobjRec("枚举来源")), "", lobjRec("枚举来源"))
        mlng体检大类 = IIf(IsNull(lobjRec("体检大类")), 0, lobjRec("体检大类"))
        mstr属性 = IIf(IsNull(lobjRec("属性")), "常规", lobjRec("属性"))
        mstr医师姓名 = IIf(IsNull(lobjRec("体检医师姓名")), "", lobjRec("体检医师姓名"))
        
        mbln是否已存在 = True
    Else
        '该体检项目在本体检记录中还不存在，直接获取体检项目设置信息。
        lstrSql = "select * from 体检管理_体检项目设置表 where 编码='" & mstr体检项目编号 & "'"
        Set lobjRec = dafuncGetData(lstrSql)
        If lobjRec.recordcount > 0 Then
            '体检项目存在。
            mstr体检项目名称 = lobjRec("名称")
            mstr缺省值 = IIf(IsNull(lobjRec("缺省值")), "", lobjRec("缺省值"))
            mstr枚举来源 = IIf(IsNull(lobjRec("枚举来源")), "", lobjRec("枚举来源"))
            mlng体检大类 = IIf(IsNull(lobjRec("体检大类")), 0, lobjRec("体检大类"))
            mstr属性 = IIf(IsNull(lobjRec("属性")), "常规", lobjRec("属性"))
        Else
            '体检项目不存在。
            Err.Raise 6666, , "当前体检记录的体检项目“”已不存在。请进入“业务设置”的“体检项目设置”界面设置体检项目。"
        End If
    End If
    
    lobjRec.Close
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "ClsFactTestItem", "subGetProperty", Err.Number, Err.Description, True
End Sub

'功能：修改体检结果表中指定系统编号的指定项目的体检结果。
'作者：杨春
Public Sub SubSave()
    On Error GoTo errHandler
    Dim lstrSql As String
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select * from 体检管理_体检结果信息表 where 系统编号='" & mstr系统编号 & "' and 体检项目='" & mstr体检项目编号 & "'")
    If lobjRec.recordcount > 0 Then
        '修改记录。
        lstrSql = "update 体检管理_体检结果信息表 set 体检结果='" & 体检结果 & "'" _
            & ",体检医师=" & IIf(医师编号 = "", "null", "'" & 医师编号 & "'") & ",填写日期='" & 填写日期 & "'" _
            & " where 系统编号='" & mstr系统编号 & "' and 体检项目='" & mstr体检项目编号 & "'"
    Else
        '插入记录。
        lstrSql = "insert into 体检管理_体检结果信息表(系统编号,体检项目,体检结果,体检医师,填写日期) values(" _
            & "'" & mstr系统编号 & "','" & mstr体检项目编号 & "','" & 体检结果 & "'," & IIf(医师编号 = "", "null", "'" & 医师编号 & "'") & ",'" & 填写日期 & "')"
    End If
    dafuncGetData lstrSql
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsFactTestItem", "subSave", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：清空本对象属性。
'作者：杨春
Private Sub subClear()
    On Error Resume Next
    体检结果 = ""
    医师编号 = ""
    填写日期 = Format(Date, "yyyy-mm-dd")
    
    mstr体检项目名称 = ""
    mstr缺省值 = ""
    mstr枚举来源 = ""
    mlng体检大类 = 0
    mstr属性 = ""
    mstr医师姓名 = um用户名
End Sub
