VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls服务器设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CF3362800EB"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Mstr服务启动时间 As String          '定义变量记录服务启时间
Public Mstr照片目录 As String              '定义量记录照片目录
Public Mstr服务间隔时间 As String          '定义变量记录服务间隔时间
Public Mstr服务时间终 As String
Public MlngDataDay As Long                 '数据期限
Public mblnClear As Boolean
Public Mstr最后运行时间 As String

'功能：将服务器的设置值保存到数据库中
'输入：无
'输出：保存到数据库中
'作者：徐冀川
'时间：2002/05/29
'##ModelId=3CF42A2C01DF
Public Sub sub保存设置()
    Dim lstrSql As String           '定义变量存放SQL语句
    
    '数据期限
    lstrSql = "update  系统管理_导照片服务设置信息表 set 设置值='" & MlngDataDay & _
              "' where 设置项='数据期限'"
    dafuncGetData (lstrSql)
    '清空
    lstrSql = "update  系统管理_导照片服务设置信息表 set 设置值='" & IIf(mblnClear, 1, 0) & _
              "' where 设置项='清空'"
    dafuncGetData (lstrSql)

    '保存启动时间
    lstrSql = "update  系统管理_导照片服务设置信息表 set 设置值='" & Mstr服务启动时间 & _
              "' where 设置项='" & "服务启动时间" & "'"
    dafuncGetData (lstrSql)
    
    '保存间隔时间
    lstrSql = "update  系统管理_导照片服务设置信息表 set 设置值='" & Mstr服务间隔时间 & _
              "' where 设置项='" & "服务间隔时间" & "'"
    dafuncGetData (lstrSql)
    
    '保存服务时间终
    lstrSql = "update  系统管理_导照片服务设置信息表 set 设置值='" & Mstr服务时间终 & _
              "' where 设置项='" & "服务时间终" & "'"
    dafuncGetData (lstrSql)
    
    '保存照片目录
    lstrSql = "update  系统管理_导照片服务设置信息表 set 设置值='" & Mstr照片目录 & _
              "' where 设置项= '" & "照片目录路径" & "'"
    dafuncGetData (lstrSql)
    
End Sub

'功能：将数据库中存储设置值，获取给系统。
'输入：无
'输出：无
'作者：徐冀川
'时间：2002/05/29
'##ModelId=3CF42A3A033D
Public Sub sub获取设置()
    Dim lobjRec As Object         '定义记录集对象
    Dim lstrSql As String         '定义SQL语句
        
    On Error Resume Next
    lstrSql = "select * from 系统管理_导照片服务设置信息表 where 设置项='数据期限'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    If lobjRec.RecordCount > 0 Then
        MlngDataDay = Val(IIf(IsNull(lobjRec("设置值")), "", lobjRec("设置值")))
        If MlngDataDay < 1 Then MlngDataDay = 30
    End If
    
    lstrSql = "select * from 系统管理_导照片服务设置信息表 where 设置项='清空'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    If lobjRec.RecordCount > 0 Then
        If IIf(IsNull(lobjRec("设置值")), "0", lobjRec("设置值")) = 1 Then
            mblnClear = True
        Else
            mblnClear = False
        End If
    End If
    
    lstrSql = "select * from 系统管理_导照片服务设置信息表 where 设置项='" & "服务启动时间" & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    If lobjRec.RecordCount > 0 Then
        Mstr服务启动时间 = IIf(IsNull(lobjRec("设置值")), "2", lobjRec("设置值"))
    End If
    
    
    lstrSql = "select * from 系统管理_导照片服务设置信息表 where 设置项='" & "照片目录路径" & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    If lobjRec.RecordCount > 0 Then
        Mstr照片目录 = IIf(IsNull(lobjRec("设置值")), "", lobjRec("设置值"))
    End If
    
    
    lstrSql = "select * from 系统管理_导照片服务设置信息表 where 设置项='" & "服务间隔时间" & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    If lobjRec.RecordCount > 0 Then
        Mstr服务间隔时间 = IIf(IsNull(lobjRec("设置值")), "1", lobjRec("设置值"))
    End If
    
    lstrSql = "select * from 系统管理_导照片服务设置信息表 where 设置项='" & "服务时间终" & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    If lobjRec.RecordCount > 0 Then
        Mstr服务时间终 = IIf(IsNull(lobjRec("设置值")), "1", lobjRec("设置值"))
    End If
    
    lstrSql = "select * from 系统管理_导照片服务设置信息表 where 设置项='" & "最后运行时间" & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    If lobjRec.RecordCount > 0 Then
        Mstr最后运行时间 = IIf(IsNull(lobjRec("设置值")), "1", lobjRec("设置值"))
    End If
    
End Sub
