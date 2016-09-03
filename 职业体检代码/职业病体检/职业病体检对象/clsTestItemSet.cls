VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestItemSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public 体检项目编码 As String
Public 体检大类 As Long
Public 属性 As String '常规/化验。

Private Sub Class_Initialize()
    体检大类 = -1
    体检项目编码 = ""
    属性 = ""
End Sub

'功能：获取所有体检项目(或指定体检大类、体检项目编码）的信息
'输入：无。
'返回：Recordset[编码，名称，默认值，可选值，属性]。
'作者：何啸天
Public Property Get 体检项目() As Object
    Dim lstrFilter As String
    
    On Error GoTo errHandler
    '拼条件。
    If 体检项目编码 <> "" Then lstrFilter = "编码='" & 体检项目编码 & "'"
    If 体检大类 <> -1 Then lstrFilter = lstrFilter & IIf(lstrFilter = "", "", " and ") & "体检大类=" & 体检大类
    If 属性 <> "" Then lstrFilter = lstrFilter & IIf(lstrFilter = "", "", " and ") & "属性='" & 属性 & "'"
    
    '执行sql语句。
    '作者：何啸天 'GERMAN
    '数据库已经更改
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set 体检项目 = dafuncGetData("select * from 职业病体检_体检项目设置表" & IIf(lstrFilter = "", "", " where " & lstrFilter))
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsTestItemSet", "Property Get 体检项目", 6666, lstrError, False
    
End Property

'功能：根据系统编号，划价
'参数：体检系统编号
'返回：体检项目总价
Public Function fun划价(ByVal paraNo As String) As Collection
    Dim lstrSql As String
    Dim lobjRec As Object
    Dim lcolItem As Collection
    Dim lcolInfo As New Collection
    Dim lsngTotal As Single '总金额。
    lsngTotal = 0
    On Error GoTo errHandler
    dasubSetQueryTimeout 6000
    lstrSql = "select * from 职业病体检_体检收费视图 where 系统编号='" & paraNo & "'"
    dasubSetQueryTimeout 600
    Set lobjRec = dafuncGetData(lstrSql)
    If lobjRec.RecordCount > 0 Then
        lobjRec.MoveFirst
        Set lcolItem = New Collection
        lcolItem.Add lobjRec("单位申请编号").Value, "单位编号"
        lcolItem.Add lobjRec("单位名称").Value, "单位名称"
        lcolItem.Add "011", "体检项目"
        lcolItem.Add lobjRec("体检类型").Value, "体检类型"
        lcolItem.Add lobjRec("体检类别").Value, "体检类别"
        '加单价总和
        Do While Not lobjRec.EOF
            If IsNull(lobjRec("单价")) = False Then
            lsngTotal = Format(lsngTotal + lobjRec("单价"), "0.00")
            End If
            lobjRec.MoveNext
        Loop
        
        lcolItem.Add lsngTotal, "单价"
        lcolItem.Add 1, "数量"
        
        lcolInfo.Add lcolItem, lcolItem("体检项目")
         
    End If
    Set fun划价 = lcolInfo
    Exit Function
errHandler:
    sfsub错误处理 "体检对象", "clstestitemset", "fun划价", Err.Number, Err.Description, False
End Function

''功能：向 职业病体检_体检基本信息表 中填充收费金额
''参数：系统编号，收费金额
''返回：boolean
'Public Function fun填金额(ByVal paraNo As String, ByVal totalpay As Double) As Boolean
'    Dim lstrSql As String
'    Dim lobjRec As Object
'    On Error GoTo errHandler
'    lstrSql = "select * from 职业病体检_体检基本信息表 where 系统编号='" & paraNo & "'"
'    Set lobjRec = dafuncGetData(lstrSql)
'    If Not lobjRec.RecordCount = 1 Then
'        MsgBox "未找到该系统编号或该系统编号不唯一，请检查！"
'        fun填金额 = False
'        Exit Function
'    End If
'    lstrSql = "update 职业病体检_体检基本信息表 set 收费金额=" & totalpay & " where 系统编号='" & paraNo & "'"
'    dafuncGetData (lstrSql)
'    fun填金额 = True
'    Exit Function
'errHandler:
'    sfsub错误处理 "体检对象", "clstestitemset", "fun填金额", Err.Number, Err.Description, False
'End Function
