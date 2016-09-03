VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mobjMdbFile As clsMdbFile  '连接mdb库文件。
Private mobjTransTable As Object   '包传输.clsTransTable

'功能：设置mdb库。
Public Property Get MdbFile() As clsMdbFile
    Set MdbFile = mobjMdbFile
End Property
Public Property Set MdbFile(ByVal vNewValue As clsMdbFile)
    On Error GoTo errHandler
    
    Set mobjMdbFile = vNewValue
    
    With mobjTransTable
        '设置源指向mdb库。
        .pstrSourceServer = mobjMdbFile.FileName     '源服务器名。
        .pstrSourceDatabase = ""                     '源数据库名。
        .pstrSourceType = "MDB"              '源数据库类型：SQL/MDB。
        .pstrSourceUserID = "Admin"
        .pstrSourcePwd = ""
                
        '建立与库的连接。
        .subConnect
    End With

    Exit Property
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsTransSystem", "Property Set MdbFile", Err.Number, Err.Description, True
    Exit Property
    Resume
End Property

'功能：获取本对象可以传输的系统表个数，便于调用者设置进度条最大值。
Public Property Get 表个数() As Long
    表个数 = 25
End Property

Private Sub Class_Initialize()
    Dim lstrServer  As String
    Dim lstrDatabase  As String
    
    On Error GoTo errHandler
    
    Set mobjTransTable = CreateObject("包传输.clsTransTable")
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrDatabase = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    
    With mobjTransTable
        '设置目的服务器、库、口令。
        .pstrDestnationServer = lstrServer
        .pstrDestnationDatabase = lstrDatabase
        .pstrDestnationType = "SQL"
        .pstrDestnationUserID = "user26"
        .pstrDestnationPwd = "welcome"
        
    End With
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsTransSystem", "Class_Initialize", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mobjMdbFile = Nothing
    Set mobjTransTable = Nothing
End Sub

'功能：传输系统管理的所有表内容。
'假设：导入（mdb或sqlserver）库之前，必须先调用本对象提供的相应方法清空库中的系统表。
'输入：para进度条 as ProgressBar 用来显示传输进度。
'      paraInput as boolean      True输入/False输出
'作者：杨春。
Public Sub sub传输所有系统表(ByVal para进度条 As Object, ByVal paraInput As Boolean)
    On Error GoTo errHandler
    
    '依次传输各表。
    '1
    mobjTransTable.sub传输表内容 "系统管理_系统基本配置表"
    para进度条.Value = para进度条.Value + 1
    '2
    mobjTransTable.sub传输表内容 "系统管理_系统安装信息表"
    para进度条.Value = para进度条.Value + 1
    '3
    mobjTransTable.sub传输表内容 "系统管理_员工基本信息表"
    para进度条.Value = para进度条.Value + 1
    '4
    mobjTransTable.sub传输表内容 "系统管理_可用操作信息表"
    para进度条.Value = para进度条.Value + 1
    '5
    mobjTransTable.sub传输表内容 "系统管理_字典_字典管理级别表"
    para进度条.Value = para进度条.Value + 1
    '6
    mobjTransTable.sub传输表内容 "系统管理_字典_用户管理级别表"
    para进度条.Value = para进度条.Value + 1
    '7
    
    mobjTransTable.sub传输表内容 "系统管理_字典_字典表列表"
    para进度条.Value = para进度条.Value + 1
    '8
    mobjTransTable.sub传输表内容 "系统管理_字典_字典内容表"
    para进度条.Value = para进度条.Value + 1
    '9
    mobjTransTable.sub传输表内容 "系统管理_字典_其它字典配置表"
    para进度条.Value = para进度条.Value + 1
    '10
    mobjTransTable.sub传输表内容 "系统管理_科室字典表"
    para进度条.Value = para进度条.Value + 1
    '11
    mobjTransTable.sub传输表内容 "系统管理_系统用户信息表"
    para进度条.Value = para进度条.Value + 1
    '12
    mobjTransTable.sub传输表内容 "系统管理_用户分组信息表"
    para进度条.Value = para进度条.Value + 1
    '13
    mobjTransTable.sub传输表内容 "系统管理_用户管辖片区表"
    para进度条.Value = para进度条.Value + 1
    '14
    mobjTransTable.sub传输表内容 "系统管理_用户组权限表"
    para进度条.Value = para进度条.Value + 1
    '15
    mobjTransTable.sub传输表内容 "系统管理_工作站基本信息表"
    para进度条.Value = para进度条.Value + 1
    '16
    mobjTransTable.sub传输表内容 "系统管理_平台操作组分类表"
    para进度条.Value = para进度条.Value + 1
    '17
    mobjTransTable.sub传输表内容 "系统管理_平台操作分组表"
    para进度条.Value = para进度条.Value + 1
    '18
    mobjTransTable.sub传输表内容 "系统管理_用户操作权限表"
    para进度条.Value = para进度条.Value + 1
    '19
    mobjTransTable.sub传输表内容 "系统管理_录入信息配置表"
    para进度条.Value = para进度条.Value + 1
    '20
    mobjTransTable.sub传输表内容 "系统管理_业务主推信息表"
    para进度条.Value = para进度条.Value + 1
    '21
    mobjTransTable.sub传输表内容 "系统管理_数据冲突操作表"
    para进度条.Value = para进度条.Value + 1
    '22
    mobjTransTable.sub传输表内容 "系统管理_系统锁定记录表"
    para进度条.Value = para进度条.Value + 1
    '23
    mobjTransTable.sub传输表内容 "系统管理_组管辖片区卫生种类表"
    para进度条.Value = para进度条.Value + 1
    '24
    mobjTransTable.sub传输表内容 "系统管理_系统图片管理表"
    para进度条.Value = para进度条.Value + 1

    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsTransSystem", "sub传输所有系统表", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：清空mdb库中系统管理的相关表内容。
'作者：杨春
Public Sub sub清空Mdb库()
    On Error GoTo errHandler
    
    With mobjMdbFile
        '1
        .subExecuteSql "delete from [系统管理_系统安装信息表]"
        '2
        .subExecuteSql "delete from [系统管理_系统基本配置表]"
        '3
        .subExecuteSql "delete from [系统管理_用户操作权限表]"
        '4
        .subExecuteSql "delete from [系统管理_业务主推信息表]"
        '5
        .subExecuteSql "delete from [系统管理_可用操作信息表]"
        '6
        .subExecuteSql "delete from [系统管理_平台操作组分类表]"
        '7
        .subExecuteSql "delete from [系统管理_平台操作分组表]"
        '8
        .subExecuteSql "delete from [系统管理_录入信息配置表]"
        '9
        .subExecuteSql "delete from [系统管理_员工基本信息表]"
        '10
        .subExecuteSql "delete from [系统管理_系统用户信息表]"
        '11
        .subExecuteSql "delete from [系统管理_科室字典表]"
        
        '12
        .subExecuteSql "delete from [系统管理_数据冲突操作表]"
        '13
        .subExecuteSql "delete from [系统管理_系统锁定记录表]"
        '14
        .subExecuteSql "delete from [系统管理_系统图片管理表]"
        '15
        .subExecuteSql "delete from [系统管理_用户分组信息表]"
        '16
        .subExecuteSql "delete from [系统管理_用户管辖片区表]"
        '17
        .subExecuteSql "delete from [系统管理_字典_字典内容表]"
        '18
        .subExecuteSql "delete from [系统管理_字典_字典表列表]"
        '19
        .subExecuteSql "delete from [系统管理_字典_字典管理级别表]"
        '20
        .subExecuteSql "delete from [系统管理_字典_用户管理级别表]"
        '21
        .subExecuteSql "delete from [系统管理_字典_其它字典配置表]"
        '22
        .subExecuteSql "delete from [系统管理_组管辖片区卫生种类表]"
        '23
        .subExecuteSql "delete from [系统管理_工作站基本信息表]"
        '24
        .subExecuteSql "delete from [系统管理_用户组权限表]"
        
    End With
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsTransSystem", "sub清空Mdb库", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：清空mdb库中系统管理的相关表内容。
'作者：杨春
Public Sub sub清空Sql库()
    On Error GoTo errHandler
    
    '执行存储过程，清空所有的数据。
    dafuncGetData "exec 系统管理_清除数据"
    
'    dafuncGetData "delete from [系统管理_系统锁定记录表]"
'    dafuncGetData "delete from [系统管理_用户操作权限表]"
'    dafuncGetData "delete from [系统管理_数据冲突操作表]"
'    dafuncGetData "delete from [系统管理_可用操作信息表]"
'    dafuncGetData "delete from [系统管理_系统图片管理表]"
'    dafuncGetData "delete from [系统管理_系统安装信息表]"
'
'    dafuncGetData "delete from [系统管理_系统基本配置表]"
'
'    dafuncGetData "delete from [系统管理_业务主推信息表]"
'
'    dafuncGetData "delete from [系统管理_平台操作组分类表]"
'
'    dafuncGetData "delete from [系统管理_平台操作分组表]"
'
'    dafuncGetData "delete from [系统管理_录入信息配置表]"
'
'
'    dafuncGetData "delete from [系统管理_片区字典表]"
'    dafuncGetData "delete from [系统管理_用户分组信息表]"
'    dafuncGetData "delete from [系统管理_系统错误记录表]"
'    dafuncGetData "delete from [系统管理_用户管辖片区表]"
'    dafuncGetData "delete from [系统管理_系统用户信息表]"
'    dafuncGetData "delete from [系统管理_员工基本信息表]"
'    dafuncGetData "delete from [系统管理_科室字典表]"
'
'    dafuncGetData "delete from [系统管理_字典_用户管理级别表]"
'    dafuncGetData "delete from [系统管理_字典_字典表列表]"
'    dafuncGetData "delete from [系统管理_字典_字典内容表]"
'    dafuncGetData "delete from [系统管理_字典_其它字典配置表]"
'    dafuncGetData "delete from [系统管理_字典_字典管理级别表]"
'
'    dafuncGetData "delete from [系统管理_工作站基本信息表]"
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsTransSystem", "sub清空Sql库", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：导入或导出指定表名的内容。
'输入：para表名 as string 数据库中表的名称。
'作者：刘浩
'Private Sub sub传输表内容2(ByVal para表名 As String)
'    Dim lobjRec As ADODB.Recordset        '从源数据库中读出的数据
'    Dim lstrQuerySql As String
'    Dim lstrSQL As String       '拼出的SQL语句
'    Dim lstrExecSql As String
'
'    Dim lstrItem As String
'    Dim i As Long
'
'    On Error GoTo errHandler
'
'    If mobjMdbFile.InputOrOutput Then
'
'        '设置Identity Insert=on
'        lstrExecSql = "select objectproperty(object_id('" & para表名 & "'),'TableHasIdentity')"
'        Set lobjRec = dafuncGetData(lstrExecSql)
'        If lobjRec.Fields(0) = 1 Then
'            dafuncGetData "SET IDENTITY_INSERT " & para表名 & " ON"
'        End If
'
'        '导入数据前.  先从MDB取出需要导入的数据
'        lstrExecSql = "select * from " & para表名
'        Set lobjRec = mobjMdbFile.funcExecuteSql(lstrExecSql)
'
'        '拼插入sql语句。
'        lstrSQL = "insert into " & para表名 & "("
'        For i = 0 To lobjRec.Fields.Count - 1
'            lstrSQL = lstrSQL & lobjRec.Fields(i).Name & ","
'        Next
'        lstrSQL = Left(lstrSQL, Len(lstrSQL) - 1) & ") "
'
'        While Not lobjRec.EOF
'            '拼插入语句。
'            lstrExecSql = lstrSQL & " values("
'            For i = 0 To lobjRec.Fields.Count - 1
'                lstrExecSql = lstrExecSql & "'" & lobjRec(i).Value & "',"
'            Next
'            lstrExecSql = Left(lstrExecSql, Len(lstrExecSql) - 1) & ")"
'
'            dafuncGetData lstrExecSql
'            lobjRec.MoveNext
'        Wend
'
'        '设置Identity Insert=Off
'        lstrSQL = "select objectproperty(object_id('" & para表名 & "'),'TableHasIdentity')"
'        Set lobjRec = dafuncGetData(lstrSQL)
'        If lobjRec.Fields(0) = 1 Then
'            dafuncGetData "SET IDENTITY_INSERT " & para表名 & " OFF"
'        End If
'
'    Else
'        '导出数据前，先从SQLserver取出需要导入的数据
'        lstrSQL = "select * from " & para表名
'        Set lobjRec = dafuncGetData(lstrSQL)
'
'        lstrSQL = "insert into " & para表名 & "("
'        For i = 0 To lobjRec.Fields.Count - 1
'            lstrSQL = lstrSQL & lobjRec.Fields(i).Name & ","
'        Next
'        lstrSQL = Left(lstrSQL, Len(lstrSQL) - 1) & ")"
'        While Not lobjRec.EOF
'            lstrExecSql = lstrSQL & " values("
'            For i = 0 To lobjRec.Fields.Count - 1
'                If lobjRec.Fields(i).Type = adDate Or lobjRec.Fields(i).Type = adDBDate Or lobjRec.Fields(i).Type = adDBTime Then
'                    lstrExecSql = lstrExecSql & "#" & lobjRec(i).Value & "#,"
'                Else
'                    lstrExecSql = lstrExecSql & "'" & lobjRec(i).Value & "',"
'                End If
'            Next
'            lstrExecSql = Left(lstrExecSql, Len(lstrExecSql) - 1) & ")"
'
'            '写入mdb库。
'            mobjMdbFile.subExecuteSql lstrExecSql
'
'            lobjRec.MoveNext
'        Wend
'    End If
'    Exit Sub
'errHandler:
'    sfsub错误处理 "体检对外接口部件", "clsTransSystem", "sub传输表内容", Err.Number, Err.Description, True
'    Exit Sub
'    Resume
'End Sub


