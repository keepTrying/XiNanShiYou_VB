VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const P_SOURCE_CONNECT_ID = 1
Private Const P_DESTINATION_CONNECT_ID = 2
'连接属性。
Public pstrDestnationServer As String   '目的服务器名。
Public pstrDestnationDatabase As String '目的数据库名。
Public pstrDestnationType As String     '目的数据库类型：SQL/MDB。
Public pstrDestnationUserID As String
Public pstrDestnationPwd As String

Public pstrSourceServer As String       '源服务器名。
Public pstrSourceDatabase As String     '源数据库名。
Public pstrSourceType As String         '源数据库类型：SQL/MDB。
Public pstrSourceUserID As String
Public pstrSourcePwd As String

Private mobjDtsPa As DTS.Package2       '传输表内容的包。

'功能：类初始化时，创建包对象。
'作者：刘浩。
Private Sub Class_Initialize()
    On Error GoTo errHandler
    '创建包。
    Set mobjDtsPa = CreateObject("DTS.Package")
    mobjDtsPa.Name = "指定表内容数据传输"

    Exit Sub
errHandler:
    sfsub错误处理 "包传输", "clsTransTable", "Class_Initialize", Err.Number, Err.Description, True
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    Set mobjDtsPa = Nothing
End Sub

'功能：与源库、目的库建立连接。
'假设：已给本对象的连接属性赋值。
'作者：刘浩。
Public Sub subConnect()
    On Error GoTo errHandler
    
    '建立源数据库连接。
    subAddConnection mobjDtsPa, P_SOURCE_CONNECT_ID, pstrSourceServer, pstrSourceUserID, pstrSourcePwd, pstrSourceDatabase, pstrSourceType
           
    '建立目的数据库连接。
    subAddConnection mobjDtsPa, P_DESTINATION_CONNECT_ID, pstrDestnationServer, pstrDestnationUserID, pstrDestnationPwd, pstrDestnationDatabase, pstrDestnationType
    
    Exit Sub
errHandler:
    sfsub错误处理 "包传输", "clsTransTable", "subConnect", Err.Number, Err.Description, True
End Sub

'功能：负责具体的一张表的内容的传输（将源表内容按要求导出到目地表中）
'输入：paraTableName 需传输的源表名（约定源表和目地表的表名和结构相同）
'      paraIsUpdate  是否修改。
'作者：刘浩。
Public Sub sub传输表内容(ByVal paraTableName As String, Optional paraIsUpdate As Boolean = False)
    Dim lcolInfo As Collection     '源表名、目的表名、字段名、关键字段。
    Dim lcolItem As Collection
    Dim lobjCn As ADODB.Connection '重新连接sql库。
    Dim lobjRec As Object          'ADO.Recordset。
    Dim lstrMsg As String
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '设置包执行需要的参数：添加源表名，目地表名。
    Set lcolInfo = New Collection
    lcolInfo.Add paraTableName, "源表名"
    lcolInfo.Add paraTableName, "目地表名"
    
    '获取要传输表的字段名。
    Set lcolItem = New Collection
    Set lobjRec = dafuncGetData("select * from " & paraTableName)
    For i = 0 To lobjRec.Fields.Count - 1
        lcolItem.Add lobjRec.Fields(i).Name, lobjRec.Fields(i).Name
    Next i
    lcolInfo.Add lcolItem, "字段名"
    lobjRec.Close
    
    '重新与sql库建立连接，设置CursorLocation=asUseServer。
    Set lobjCn = New ADODB.Connection
    lobjCn.ConnectionString = dapstrDSN
    lobjCn.CursorLocation = adUseServer
    lobjCn.Open
    
    '获取要传输表的关键字段名。
    Set lcolItem = New Collection
    Set lobjRec = lobjCn.OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, paraTableName))
    While Not lobjRec.EOF
        lcolItem.Add lobjRec.Fields("COLUMN_NAME").Value
        lobjRec.MoveNext
    Wend
    lcolInfo.Add lcolItem, "主键字段"
    lobjRec.Close
    
    '开始表内容传输。
    sub包执行 lcolInfo, paraIsUpdate, " where 2 = 2 "
    
    Exit Sub
errHandler:
    sfsub错误处理 "包传输", "clsTransTable", "sub传输表内容", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：添加连接。
'输入：ParaPackage as dts.Package2；
'      paraConnectionID  连接ID ；
'      paraServerName    服务器名 ；
'      paraUserID        连接库的用户名；
'      paraPassword      访问库的口令；
'      paraCatalog       数据库名
'      paraType          库类型：Sql/MDB
Private Sub subAddConnection(ByRef paraPackage As Object, _
                         ByVal paraConnectionID As Integer, _
                         ByVal paraServerName As String, _
                         ByVal paraUserID As String, _
                         ByVal paraPassword As String, _
                         ByVal paraCatalog As String, _
                         ByVal paraType As String)
    On Error GoTo errHandler

    Dim lobjConn As DTS.Connection2
    
    If UCase(paraType) = "SQL" Then
        Set lobjConn = paraPackage.Connections.New("SQLOLEDB")
        With lobjConn
            .DataSource = paraServerName
            .ID = paraConnectionID
            .UserID = paraUserID
            .Catalog = paraCatalog
            .Reusable = True
            .Password = paraPassword
        End With
    Else
        'Microsoft.Jet.OLEDB.4.0
        Set lobjConn = paraPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
        With lobjConn
            .DataSource = paraServerName
            .ID = paraConnectionID
            .UserID = paraUserID
            .Catalog = paraCatalog
            .Reusable = True
            .Password = paraPassword
        End With
    End If
    
    paraPackage.Connections.Add lobjConn
    Set lobjConn = Nothing
    Exit Sub
errHandler:
    sfsub错误处理 "包传输", "clsTransTable", "subAddConnection", Err.Number, Err.Description, True
End Sub
'功能：传输源表数据到目地表。
'输入：paraTable as collection[源表名，目地表名，[关键字]，[字段名]]，
'      paraIsUpdate as boolean （true 在源表和目地表有相同关键字时，覆盖目地表 / false 保持目地表内容）
'      paraFilter as string
'作者：刘浩。
Private Sub sub包执行(ByVal paraTable As Collection, _
                      ByVal paraIsUpdate As Boolean, _
                      ByVal paraFilter As String)
    Dim lstrSourceName    As String                   '源表名
    Dim lstrDestName      As String                   '目地表名
    Dim lcolFields         As Collection              '字段名
    Dim lcolKeys           As Collection              '主键字段
    
    Dim lobjDtsTask       As DTS.Task                 'DTS任务
    Dim lobjDtsQueryTask  As DTS.DataDrivenQueryTask2 '数据抽取查询任务
    Dim lobjDtsTran As DTS.Transformation             'DTS转换
    Dim lobjDTsStep As DTS.Step2                      'DTS任务执行步骤
    Dim lobjScript As DTSPump.DataPumpTransformScript '设置需要处理的ActiveX语句
    Dim lobjDtsLookUp As DTS.Lookup                   'DTS查找对象
    
    Dim lstrLook As String                     '查找对象名
    Dim lstrSql As String
    Dim lstrMsg As String
    Dim i As Long, j As Long
    
    On Error GoTo errHandler
    
    lstrSourceName = paraTable("源表名")
    lstrDestName = paraTable("目地表名")
    Set lcolFields = paraTable("字段名")
    Set lcolKeys = paraTable("主键字段")
    
    '添加数据驱动查询任务。
    If mobjDtsPa.Tasks.Count > 0 Then
        mobjDtsPa.Tasks.Remove "数据传输"
    End If
    Set lobjDtsTask = mobjDtsPa.Tasks.New("DTSDataDrivenQueryTask")
    lobjDtsTask.Name = "数据传输"
    
    Set lobjDtsQueryTask = lobjDtsTask.CustomTask
    With lobjDtsQueryTask
        '设置源连接ID。
        .SourceConnectionID = P_SOURCE_CONNECT_ID
        '设置在源连接上的查询语句。
        .SourceSQLStatement = "select * from " & lstrSourceName & paraFilter
        
        '设置目的连接ID。
        .DestinationConnectionID = P_DESTINATION_CONNECT_ID
        '设置目的表名
        .DestinationObjectName = lstrDestName
        
        '定义INSERT查询语句。
        lstrSql = "insert into " & lstrDestName & " values("
        '设置插入与语句的字段。
        For i = 1 To lcolFields.Count
            lstrSql = lstrSql & "?,"
            .InsertQueryColumns.AddColumn lcolFields(i), i
        Next i
        lstrSql = Left(lstrSql, Len(lstrSql) - 1) & ")"
        .InsertQuery = lstrSql
        
        '定义update查询语句。
        lstrSql = "update " & lstrDestName & " set "
        For i = 1 To lcolFields.Count
            lstrSql = lstrSql & lcolFields(i) & " =?,"
            .UpdateQueryColumns.AddColumn lcolFields(i), i
        Next
        lstrSql = Left(lstrSql, Len(lstrSql) - 1) & " where "
        
        j = i + 1
        For i = 1 To lcolKeys.Count
            lstrSql = lstrSql & lcolKeys(i) & " = ? and "
            .UpdateQueryColumns.AddColumn lcolKeys(i), j
            j = j + 1
        Next i
        lstrSql = Left(lstrSql, Len(lstrSql) - 4)
        .UpdateQuery = lstrSql
        
    End With
    
    '设置lookup对象。
    For i = 1 To lcolKeys.Count
        Set lobjDtsLookUp = lobjDtsQueryTask.Lookups.New("look" & CStr(i))
        With lobjDtsLookUp
            .ConnectionID = P_DESTINATION_CONNECT_ID
            .Query = "select " & lcolKeys(i) & " from " & lstrDestName & " where " & lcolKeys(i) & " = ? "
        End With
        lobjDtsQueryTask.Lookups.Add lobjDtsLookUp
    Next i
    
    '设置需要传输的字段。
    Set lobjDtsTran = lobjDtsQueryTask.Transformations.New("DTS.DataPumpTransformScript")
    With lobjDtsTran
        .Name = "transform"
        For i = 1 To lcolFields.Count
            .SourceColumns.AddColumn lcolFields(i), i
            .DestinationColumns.AddColumn lcolFields(i), i
        Next i
    End With
    
    '设置传输的vbScript。
    Set lobjScript = lobjDtsTran.TransformServer
    With lobjScript
        .Text = "'**********************************************************************" & vbCrLf
        .Text = .Text & "'  Visual Basic 转换脚本" & vbCrLf
        .Text = .Text & "'************************************************************************" & vbCrLf & vbCrLf
        .Text = .Text & "Function Main()" & vbCrLf
        For i = 1 To lcolFields.Count
            .Text = .Text & Space(8) & "DTSDestination(""" & lcolFields(i) & """) = " & "DTSSource(""" & lcolFields(i) & """)" & vbCrLf
        Next i
        '源表，目地表中没有相同值时，在目地表是插入一条新记录。 否则，跟据参数决定是修改记录或跳过当前记录。
        .Text = .Text & vbCrLf
        .Text = .Text & Space(8) & "if 1 = 1 "
        For i = 1 To lcolKeys.Count
            lstrLook = "look" & CStr(i)
            .Text = .Text & "and DTSDestination(""" & lcolKeys(i) & """) = DTSLookups(""" & lstrLook & _
                """).Execute(DTSSource(""" & lcolKeys(i) & """)) "
        Next i
        .Text = .Text & "then" & vbCrLf
        If paraIsUpdate Then
            .Text = .Text & Space(16) & "Main = DTSTransformStat_UpdateQuery" & vbCrLf
        Else
            .Text = .Text & Space(16) & "Main = DTSTransformStat_SkipRow" & vbCrLf
        End If
        .Text = .Text & Space(8) & "else" & vbCrLf
        .Text = .Text & Space(16) & "Main = DTSTransformStat_InsertQuery" & vbCrLf
        .Text = .Text & Space(8) & "end if" & vbCrLf
        
        .Text = .Text & "End Function"
        .FunctionEntry = "Main"
        .Language = "VBScript"
    End With
    
    '添加数据转换对象到集合。
    lobjDtsQueryTask.Transformations.Add lobjDtsTran
    
    '添加任务到任务集合。
    mobjDtsPa.Tasks.Add lobjDtsTask
    
    '设置任务的执行步骤。
    If mobjDtsPa.Steps.Count > 0 Then
        mobjDtsPa.Steps.Remove "Step1"
    End If
    Set lobjDTsStep = mobjDtsPa.Steps.New
    lobjDTsStep.Name = "Step1"
    lobjDTsStep.ExecuteInMainThread = True
    lobjDTsStep.TaskName = "数据传输"
    mobjDtsPa.Steps.Add lobjDTsStep
    
    mobjDtsPa.Execute
    
    '获取错误信息。
    lstrMsg = funcStepErrors(mobjDtsPa)
    If lstrMsg <> "" Then
        Err.Raise 6666, , "传输“" & lstrSourceName & "”失败！" & lstrMsg
    End If
    
    Set lobjDTsStep = Nothing
    Set lobjDtsTask = Nothing
    Set lobjDtsTran = Nothing
    
    Exit Sub
errHandler:
    sfsub错误处理 "包传输", "clsTransTable", "sub包执行", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：累计统计包执行过程中产生的错误。并以字符串的形式返回所有的错误号和错误描述。
'输入：paraPackage   执行包。
'作者：刘浩。
Private Function funcStepErrors(ByVal paraPackage As DTS.Package) As String
    Dim lobjStep    As DTS.Step
    Dim lstrMessage As String
    Dim llngErrNum  As Long
    Dim lstrDescr   As String
    Dim lstrSource  As String
    
    On Error GoTo errHandler

    lstrMessage = ""
    
    '依次判断各步是否执行成功。
    For Each lobjStep In paraPackage.Steps
        If lobjStep.ExecutionStatus = DTSStepExecStat_Completed Then
            If lobjStep.ExecutionResult = DTSStepExecResult_Failure Then
            
                '获取当前步的错误信息。
                lobjStep.GetExecutionErrorInfo llngErrNum, lstrSource, lstrDescr
                
                '拼错误信息串。
                lstrMessage = lstrMessage & vbCrLf & "Step '" & lobjStep.Name & "' failed, error:  " & _
                          CStr(llngErrNum) & vbCrLf & lstrDescr
            End If
        End If
    Next
    '返回错误信息。
    funcStepErrors = lstrMessage
    
    Exit Function
errHandler:
    sfsub错误处理 "包传输", "clsTransTable", "funcStepErrors", Err.Number, Err.Description, True
End Function




