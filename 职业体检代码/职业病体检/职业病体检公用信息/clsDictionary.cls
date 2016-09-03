VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'作者：杨春
'功能：临时替代总体结果的字典管理。

'输入：字典名，条件。
Public Function Fetch(Optional ByVal strName As String = "", Optional strCondition As String = "1=1") As Object
    Dim lobjRec As Object
    
    '获取字典id。
    Set lobjRec = dafuncGetData("select ID from 系统管理_字典_字典表列表 where 名称='" & strName & "'")
    If lobjRec.recordcount = 0 Then
        Err.Raise 6666, , "指定字典不存在。"
    End If
    
    '获取字典内容。
    Set Fetch = dafuncGetData("select * from 系统管理_字典_字典内容表 where ID=" & lobjRec!ID & IIf(strCondition = "", "", " and " & strCondition))
End Function


Public Function FetchEx(ByVal strName As String, Optional strCondition As String = "") As Object
    Dim lobjRec As Object
    
    '获取字典id。
    Set lobjRec = dafuncGetData("select 数据表名 from 系统管理_字典_其它字典配置表 where 名称='" & strName & "'")
    If lobjRec.recordcount = 0 Then
        Err.Raise 6666, , "指定字典不存在。"
    End If
    
    '获取字典内容。
    Set FetchEx = dafuncGetData("select * from " & lobjRec(0) & IIf(strCondition = "", "", " where " & strCondition))

End Function

