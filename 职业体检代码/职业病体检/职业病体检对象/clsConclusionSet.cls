VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConclusionSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstr科室编号 As String

Private mstr科室名称 As String

Private mstr结论模板 As String

Private mstr医师编号 As String

Private mstr结论日期 As String

Private mstr结论标准 As String


Public Function func获取最终结论模板() As Object
    
    Dim lobjRec As Object
    Dim sql As String
    
    sql = "select * from 系统管理_字典_体检结论模板表 where 科室编号='16'"
    dasubSetQueryTimeout 600
    Set lobjRec = dafuncGetData(sql)
    
    Set func获取最终结论模板 = lobjRec
    Set lobjRec = Nothing
    
End Function

Public Function func是否存在(ByVal para结论模板 As String, ByVal para结论标准 As String) As Boolean
    
    Dim lobjRec As Object
    Dim sql As String
    Dim flag As Boolean
    
    sql = "select * from 系统管理_字典_体检结论模板表 where 结论模板='" & para结论模板 & "' and 结论标准='" & para结论标准 & "' and 科室编号='16'"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(sql)
    If lobjRec.RecordCount > 0 Then
        func是否存在 = True
    Else
        func是否存在 = False
    End If
    Set lobjRec = Nothing

End Function

Public Function func添加最终结论模板(ByVal para科室编号 As String, ByVal para科室名称 As String, ByVal para结论模板 As String, ByVal para医生编号 As String, ByVal para结论标准 As String) As Boolean
    
    Dim lobjRec As Object
    Dim sql As String
    On Error GoTo errHandler
    
    sql = "insert into 系统管理_字典_体检结论模板表 values('" & para科室编号 & "','" & para科室名称 & "','" & para结论模板 & "','" & para医生编号 & "','" & Date & "','" & para结论标准 & "')"
    
    Set lobjRec = dafuncGetData(sql)
    
    func添加最终结论模板 = True
    Set lobjRec = Nothing
    Exit Function
errHandler:
    func添加最终结论模板 = False
End Function

Public Function func修改最终结论模板(ByVal para科室编号 As String, ByVal para科室名称 As String, ByVal para结论模板 As String, ByVal para医生编号 As String, ByVal para结论日期 As String, ByVal para结论标准 As String) As Boolean
    
    Dim lobjRec As Object
    Dim sql As String
    On Error GoTo errHandler
    
    sql = "update 系统管理_字典_体检结论模板表 set 结论模板='" & para结论模板 & "',结论标准='" & para结论标准 & "' where 科室编号='" & para科室编号 & "' and 医生编号='" & para医生编号 & "' and 结论日期='" & para结论日期 & "'"
    
    Set lobjRec = dafuncGetData(sql)
    
    func修改最终结论模板 = True
    Set lobjRec = Nothing
    Exit Function
errHandler:
    func修改最终结论模板 = False
End Function

Public Function func删除最终结论模板(ByVal para结论模板 As String, ByVal para结论标准 As String) As Boolean
    
    Dim lobjRec As Object
    Dim sql As String
    On Error GoTo errHandler
    
    sql = "delete from 系统管理_字典_体检结论模板表 where 结论模板='" & para结论模板 & "' and 结论标准='" & para结论标准 & "'"
    
    Set lobjRec = dafuncGetData(sql)
    
    func删除最终结论模板 = True
    Set lobjRec = Nothing
    Exit Function
errHandler:
    func删除最终结论模板 = False
End Function

'功能：读取最终结论模板
'作者：翁乔
'时间：2012-06-01
Public Function func读取所有最终结论模板(ByVal paraSQL As String) As Object

    Set func读取所有最终结论模板 = dafuncGetData(paraSQL)

End Function
