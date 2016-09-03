VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatchManege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'2012-04-25 翁乔
'为每个科室的批量操作提供相应函数

Option Explicit

'五官科的批量保存
'作者：翁乔

Public Function func批量保存体检人员(ByVal paraSysNo As Collection, _
                                     ByVal paraDoctor As String, _
                                     ByVal paraFillDate As Date, _
                                     ByVal paraItem As Collection, _
                                     ByVal paraResult As Collection, _
                                     ByVal paraTableName As String) As Boolean

On Error GoTo errHandler
    Dim i As Integer
    Dim j As Integer
    Dim paraConclusion As String
    Dim sql As String

    For j = 1 To paraSysNo.Count
        For i = 1 To paraItem.Count
            sql = "exec autoConclusion '" & paraSysNo(j) & "','" & paraItem(i) & "','" & paraResult(i) & "','" & paraDoctor & "','" & paraFillDate & "','" & paraConclusion & "','" & paraTableName & "'"
        dafuncGetData (sql)
        Next
    Next j

    func批量保存体检人员 = True
    
    Exit Function
errHandler:
    MsgBox ("保存失败，事务回滚。")
    func批量保存体检人员 = False
End Function

'Private Function sub保存单项体检结果(ByVal paraSysNo As String, _
'                                     ByVal paraItem As String, _
'                                     ByVal paraResult As String, _
'                                     ByVal paraDoctor As String, _
'                                     ByVal paraFillDate As Date, _
'                                     ByVal paraConclusion As String, _
'                                     ByVal paraTableName As String)
'    Dim lstrSql As String
'    Dim lobjRec As Object
'    Dim lstrItemID As String
'
'    '查找对应paraItem体检项目的项目编号
'    lstrSql = "select * from 职业病体检_体检项目设置表 where 名称='" & paraItem & "'"
'    Set lobjRec = dafuncGetData(lstrSql)
'    lstrItemID = lobjRec("编码")
'
'    '检查该系统编号和体检项目是否存在。
'    '如果没有记录，则insert一条结果记录；否则，upDate该记录。
'    lstrSql = "select * from " & paraTableName & " where 系统编号='" & paraSysNo & "' and 体检项目='" & lstrItemID & "'"
'    Set lobjRec = dafuncGetData(lstrSql)
'    If lobjRec.recordcount = 0 Then
'        lstrSql = "insert into " & paraTableName & " values('" & paraSysNo & "','" & lstrItemID & "','" & paraResult & "','" & paraDoctor & "','" & paraFillDate & "','" & paraConclusion & "')"
'        dafuncGetData (lstrSql)
'    Else
'        lstrSql = "upDate " & paraTableName & " set 体检结果='" & paraResult & "',体检医师='" & paraDoctor & "',填写时间='" & paraFillDate & "',单项结论='" & paraConclusion & "' where 系统编号='" & paraSysNo & "' and 体检项目='" & lstrItemID & "'"
'        dafuncGetData (lstrSql)
'    End If
'
'End Function



