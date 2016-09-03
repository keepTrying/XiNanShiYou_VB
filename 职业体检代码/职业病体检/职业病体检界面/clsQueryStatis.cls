VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryStatis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function func读取卫生种类() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select * from 系统管理_字典_字典内容表 where ID = 12 and parent=0 and 名称 like '%卫生';")
    Set func读取卫生种类 = rs
End Function

Public Function func读取行业类别() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select * from 系统管理_行业类别字典视图")
    Set func读取行业类别 = rs
End Function

Public Function func读取辖区() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select * from 系统管理_字典_字典内容表 where id=14 and parent = 0;")
    Set func读取辖区 = rs
End Function

Public Function func读取工种() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select InnerID,名称 from 系统管理_字典_字典内容表 " _
        & "where ID = (select ID from 系统管理_字典_字典表列表 where 名称 = '工种字典');")
    Set func读取工种 = rs
End Function

Public Function func危害因素() As Object
    Set func危害因素 = dafuncGetData("select InnerID,名称 from 系统管理_字典_字典内容表 " _
            & "where ID = (select ID from 系统管理_字典_字典表列表 where 名称 = '危害种类字典');")
End Function

Public Function func读取科室() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select InnerID,名称 from 系统管理_字典_字典内容表 " _
       & "where ID = (select ID from 系统管理_字典_字典表列表 where 名称 = '职业病体检科室字典');")
    Set func读取科室 = rs
End Function

Public Function func读取体检项目() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select * from 系统管理_字典_字典内容表 where ID = 84 and 编号 not in(13,14);")
    Set func读取体检项目 = rs
End Function

Public Function func读取体检类型() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select InnerID,名称 from 系统管理_字典_字典内容表 " _
        & "where ID = (select ID from 系统管理_字典_字典表列表 where 名称 = '体检人类别字典');")
    Set func读取体检类型 = rs
End Function

Public Function func读取体检类别() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select * from 系统管理_字典_字典内容表 where id=72;")
    Set func读取体检类别 = rs
End Function

Public Function func读取体检表() As Object
    Dim rs As Object
    Set rs = dafuncGetData("select * from 职业病体检_体检表模板基本信息表;")
    Set func读取体检表 = rs
End Function

'根据医师姓名返回编号
Public Function func返回医师信息(ByVal para姓名 As String, Optional ByVal para编号 As String) As Object
    Dim rs As Object
    Set rs = dafuncGetData("exec sel返回医师信息 '" & para姓名 & "','" & para编号 & "'")
    Set func返回医师信息 = rs
End Function

'根据条件返回查询结果集
Public Function func返回查询信息(ByVal paraSQL As String) As Object
    Dim lobjRec As Object '查询结果集
    
    Set lobjRec = dafuncGetData(paraSQL)
    Set func返回查询信息 = lobjRec
    
End Function




