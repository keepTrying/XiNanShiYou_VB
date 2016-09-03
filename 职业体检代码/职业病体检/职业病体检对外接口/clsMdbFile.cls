VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMdbFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrMdbFileName As String    'mdb库文件名，带路径和后缀。

Private mblnInputOrOutput As Boolean '文件操作方向：True:输入/Fasle:O输出。

Private mobjMdbConnect As ADODB.Connection    '与mdb库的连接。
Private mobjMdbCommand As ADODB.Command       '执行mdb库上的命令。

Private mstrServer As String
Private mstrData As String

Public Property Get FileName() As String
    FileName = mstrMdbFileName
End Property

'功能：设置文件名（带路径），同时与该文件建立连接。
'作者：刘浩。
Public Property Let FileName(ByVal vNewValue As String)
    Dim lobjFso As Object
    Dim lobjFile As Object
    '先关闭旧连接。
    On Error Resume Next
    If Not mobjMdbConnect Is Nothing And mstrMdbFileName <> "" Then
        mobjMdbConnect.Close
    End If
    
    '设置数据库文件的属性为可读可写。
    Set lobjFso = CreateObject("Scripting.FileSystemObject")
    If Len(vNewValue) <> 0 Then
        Set lobjFile = lobjFso.GetFile(vNewValue)
        lobjFile.Attributes = 0
    End If
    
    On Error GoTo errHandler
    mstrMdbFileName = vNewValue
    If mstrMdbFileName <> "" Then
        '建立与mdb库的连接。
        Set mobjMdbConnect = New ADODB.Connection
        With mobjMdbConnect
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & mstrMdbFileName
            .CursorLocation = adUseClient
            .Open
        End With
        
        '创建命令对象。
        Set mobjMdbCommand = New ADODB.Command
        Set mobjMdbCommand.ActiveConnection = mobjMdbConnect
    End If
    
    Set lobjFile = Nothing
    Set lobjFso = Nothing
    Exit Property
errHandler:
    sfsub错误处理 "体检对外接口", "clsMdbFile", "Property Let FileName", Err.Number, Err.Description, True
End Property

'功能：设置（获取）文件的操作方向：输入/输出。
'输入：paraInput as boolean  true输入/false输出。
'作者：杨春。
Public Property Get InputOrOutput() As Boolean
    InputOrOutput = mblnInputOrOutput
End Property
Public Property Let InputOrOutput(ByVal paraInput As Boolean)
    mblnInputOrOutput = paraInput
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    mstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    mstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    
    Exit Sub
    
errHandler:
    sfsub错误处理 "体检对外接口", "clsMdbFile", "Class_Initialize", Err.Number, Err.Description, True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mobjMdbCommand = Nothing
    Set mobjMdbConnect = Nothing
End Sub

'功能：执行查询sql语句。
'作者：刘浩。
Public Function funcExecuteSql(ByVal paraSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    mobjMdbCommand.CommandText = paraSql
    Set funcExecuteSql = mobjMdbCommand.Execute
    
    funcExecuteSql.ActiveConnection = Nothing
    
    Exit Function
errHandler:
    sfsub错误处理 "体检对外接口", "clsMdbFile", "funcExecuteSql", Err.Number, Err.Description, True
End Function

'功能：执行动作sql语句。
'作者：刘浩。
Public Sub subExecuteSql(ByVal paraSql As String)
    On Error GoTo errHandler
    mobjMdbCommand.CommandText = paraSql
    mobjMdbCommand.Execute

    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口", "clsMdbFile", "subExecuteSql", Err.Number, Err.Description, True
End Sub


'功能：清空当前mdb库中所有表内容。
'作者：刘浩。
Public Sub Sub清空文件()
    On Error GoTo errHandler
    
    subExecuteSql "delete from 体检管理_数据传输类型表"
    
    subExecuteSql "delete from 体检管理_数据传输范围表"
    
    subExecuteSql "delete from [体检管理_数据分类信息表] "
    
    subExecuteSql "delete from [体检管理_体检人员基本信息表] "
    
    subExecuteSql "delete from [体检管理_体检基本信息表] "
    
    subExecuteSql "delete from [体检管理_体检附加信息表] "
    
    subExecuteSql "delete from [体检管理_体检结果信息表] "
    
    subExecuteSql "delete from [体检管理_化验单基本信息表] "
    
    subExecuteSql "delete from [体检管理_化验单不正常结果信息表] "
    
    subExecuteSql "delete from [体检管理_体检表模板基本信息表] "
    
    subExecuteSql "delete from [体检管理_体检表模板附加项目信息表] "
    
    subExecuteSql "delete from [体检管理_体检表模板体检项目表] "
    
    subExecuteSql "delete from [体检管理_体检表模板体检结论表] "
    
    subExecuteSql "delete from [体检管理_体检项目设置表] "
    
    subExecuteSql "delete from [体检管理_体检医师项目设置表] "
    
    subExecuteSql "delete from [体检管理_体检结论判断条件表] "
    
    subExecuteSql "delete from [体检管理_编号生成规则表] "
    
    subExecuteSql "delete from [体检管理_体检人员附加项目设置表] "
    
    subExecuteSql "delete from [体检管理_体检访问标志表] "
    
    subExecuteSql "delete from [体检管理_化验单试管编号表] "
    
    subExecuteSql "delete from [体检管理_化验单不正常结果试管编号表] "
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口", "clsMdbFile", "Sub清空文件", Err.Number, Err.Description, True
End Sub


'功能：导入或导出"体检管理_体检表模板附加项目信息表" 的内容
'作者：刘浩。
Public Sub sub传输体检表模板附加项目信息表()
    Dim lobjRec As Object       '从源数据库中读出的数据(过滤掉了不需要被导入或导出的数据)
    Dim lstrSQL As String       '
    
    On Error GoTo errHandler
    
    '分析数据导出的过滤条件
    lstrSQL = "select * from 体检管理_体检表模板附加项目信息表 "
    
    If mblnInputOrOutput Then
        '导入数据时,从MDB获取数据。
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        '把数据写入SQL数据库中。
        While Not lobjRec.EOF
            '如果要导入的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_体检表模板附加项目信息表 where 体检表名称 = '" & _
                lobjRec.Fields("体检表名称") & "' and  附加项目 = '" & lobjRec.Fields("附加项目") & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_体检表模板附加项目信息表(体检表名称,附加项目,序号,是否必录) values('" & _
                lobjRec.Fields("体检表名称").Value & "','" & _
                lobjRec.Fields("附加项目").Value & "','" & _
                lobjRec.Fields("序号").Value & "','" & _
                lobjRec.Fields("是否必录").Value & "')"
            
            dafuncGetData lstrSQL
            
            lobjRec.MoveNext
        Wend
    Else
        '导出数据时,从SQL数据
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检表模板附加项目信息表(体检表名称,附加项目,序号,是否必录) values('" & _
                lobjRec.Fields("体检表名称").Value & "','" & _
                lobjRec.Fields("附加项目").Value & "','" & _
                lobjRec.Fields("序号").Value & "','" & _
                lobjRec.Fields("是否必录").Value & "')"
            
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输体检表模板附加项目信息表", Err.Number, Err.Description, True
End Sub

'功能： 导入或导出"sub体检管理_体检人员基本信息表" 的内容
'输入： paraRange As Collection [数据范围名，数据范围值]。
'作者： 刘浩。
Public Sub sub传输体检人员基本信息表(paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据(过滤掉了不需要被导入或导出的数据)
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    '分析数据导出的过滤条件
    lstrSQL = "select * from 体检管理_体检人员基本信息表 where 健康档案编号 in(select distinct 健康档案编号 from 体检管理_体检基本数据库  where 1 = 1 "
    lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput) & ")"
    
    If mblnInputOrOutput Then
        '导入数据时,从MDB获取数据。
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        '把数据写入SQL数据库中。
        While Not lobjRec.EOF
            '只导入不存在的体检人员。
            lstrSQL = "if not exists(select 健康档案编号 from 体检管理_体检人员基本信息表 where 健康档案编号 = '" & lobjRec("健康档案编号").Value & "') insert into 体检管理_体检人员基本信息表(健康档案编号,公民身份号码,姓名,性别,出生日期,单位申请编号," & _
                "单位名称,建档日期,卫生种类,片区,行业类别) values('" & lobjRec("健康档案编号").Value & "','" & _
                lobjRec("公民身份号码").Value & "','" & _
                lobjRec("姓名").Value & "','" & lobjRec("性别").Value & "'," & _
                IIf(IsNull(lobjRec("出生日期").Value), "NUll", "'" & Format(lobjRec("出生日期").Value, "yyyy-mm-dd") & "'") & ",'" & _
                lobjRec("单位申请编号").Value & "','" & _
                lobjRec("单位名称").Value & "'," & _
                IIf(IsNull(lobjRec("建档日期").Value), "NUll", "'" & Format(lobjRec("建档日期").Value, "yyyy-mm-dd") & "'") & ",'" & _
                lobjRec("卫生种类").Value & "','" & _
                lobjRec("片区").Value & "','" & lobjRec("行业类别").Value & "')"
                
                dafuncGetData lstrSQL
           
            lobjRec.MoveNext
        Wend
    Else
        '导出数据时,从SQL数据
        Set lobjRec = dafuncGetData(lstrSQL)
        
        '把数据写入MDB文件
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检人员基本信息表(健康档案编号,公民身份号码,姓名,性别,出生日期,单位申请编号," & _
                    "单位名称,建档日期,卫生种类,片区,行业类别) values('" & lobjRec("健康档案编号").Value & "','" & _
                    lobjRec("公民身份号码").Value & "','" & _
                    lobjRec("姓名").Value & "','" & lobjRec("性别").Value & "'," & _
                    IIf(IsNull(lobjRec("出生日期").Value), "NUll", "#" & lobjRec("出生日期").Value & "#") & ",'" & _
                    lobjRec("单位申请编号").Value & "','" & _
                    lobjRec("单位名称").Value & "'," & _
                    IIf(IsNull(lobjRec("建档日期").Value), "NUll", "#" & lobjRec("建档日期").Value & "#") & ",'" & _
                    lobjRec("卫生种类").Value & "','" & _
                    lobjRec("片区").Value & "','" & lobjRec("行业类别").Value & "')"
                    
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub传输体检人员基本信息表", Err.Number, Err.Description, True
End Sub

'功能： 导入或导出"体检管理_体检基本信息表" 的内容
'输入： paraRange As Collection [数据范围名，数据范围值]。
'作者： 刘浩。
Public Sub sub传输体检基本信息表(paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据(过滤掉了不需要被导入或导出的数据)
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据时,从MDB获取数据
        lstrSQL = "select * from 体检管理_体检基本数据库 where 1 = 1 "
        lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
        
        Set lobjRec = funcExecuteSql(lstrSQL)
                
        '把数据写入SQL数据库中
        While Not lobjRec.EOF
            '只导入不存在的体检记录。
            lstrSQL = "if not exists(select * from 体检管理_体检基本信息表 where 系统编号 = '" & lobjRec("系统编号").Value & "') "
            lstrSQL = lstrSQL & "insert into 体检管理_体检基本信息表(系统编号,健康档案编号,试管编号," & _
                "体检表名称,体检类别,体检日期," & _
                "体检结论,诊断和处理意见,下结论医师," & _
                "复查体检表名,复查系统编号,体检状态,收费批号,下结论日期) " & _
                "values('" & lobjRec("系统编号").Value & "','" & _
                lobjRec("健康档案编号").Value & "','" & lobjRec("试管编号").Value & "','" & _
                lobjRec("体检表名称").Value & "','" & lobjRec("体检类别").Value & "'," & _
                IIf(IsNull(lobjRec("体检日期").Value), "NULL", "'" & Format(lobjRec("体检日期").Value, "yyyy-mm-dd") & "'") & ",'" & _
                lobjRec("体检结论").Value & "','" & _
                lobjRec("诊断和处理意见").Value & "'," & _
                IIf(IsNull(lobjRec("下结论医师").Value), "null", "'" & lobjRec("下结论医师").Value & "'") & ",'" & _
                lobjRec("复查体检表名").Value & "','" & lobjRec("复查系统编号").Value & "','" & _
                lobjRec("体检状态").Value & "','" & lobjRec("收费批号").Value & "'," & IIf(IsNull(lobjRec!下结论日期), "null", "'" & lobjRec("下结论日期") & "'") & ")"
        
            dafuncGetData lstrSQL
            
            lobjRec.MoveNext
        Wend
    Else
        '导出数据时,从SQL获取数据
        lstrSQL = "select * from 体检管理_体检基本数据库 where 1 = 1 "
        lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
        Set lobjRec = dafuncGetData(lstrSQL)
                
        '把数据写入MDB文件
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检基本信息表(系统编号,健康档案编号,试管编号," & _
                    "体检表名称,体检类别,体检日期," & _
                    "体检结论,诊断和处理意见,下结论医师," & _
                    "复查体检表名,复查系统编号,体检状态,收费批号,下结论日期) " & _
                    "values('" & lobjRec("系统编号").Value & "','" & _
                    lobjRec("健康档案编号").Value & "','" & lobjRec("试管编号").Value & "','" & _
                    lobjRec("体检表名称").Value & "','" & lobjRec("体检类别").Value & "'," & _
                    IIf(IsNull(lobjRec("体检日期").Value), "NULL", "#" & lobjRec("体检日期").Value & "#") & ",'" & _
                    lobjRec("体检结论").Value & "','" & _
                    lobjRec("诊断和处理意见").Value & "'," & _
                    IIf(IsNull(lobjRec("下结论医师").Value), "null", "'" & lobjRec("下结论医师").Value & "'") & ",'" & _
                    lobjRec("复查体检表名").Value & "','" & lobjRec("复查系统编号").Value & "','" & _
                    lobjRec("体检状态").Value & "','" & lobjRec("收费批号").Value & "'," & IIf(IsNull(lobjRec!下结论日期), "null", "#" & lobjRec("下结论日期") & "#") & ")"
            
            subExecuteSql lstrSQL

            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub传输体检基本信息表", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能： 导入或导出"体检管理_体检附加信息表" 的内容
'输入： paraRange As Collection [数据范围名，数据范围值]
Public Sub sub传输体检附加信息表(paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据(过滤掉了不需要被导入或导出的数据)
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    lstrSQL = "select 体检管理_体检附加信息表.* from 体检管理_体检附加信息表 ,体检管理_体检基本数据库 " & _
                "where 体检管理_体检附加信息表.系统编号 = 体检管理_体检基本数据库.系统编号 "
    
    If mblnInputOrOutput Then
        '导入数据前.  先根据过滤条件从MDB取出需要导入的数据
        lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
        
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '只导入不存在的纪录。
            lstrSQL = "if not exists(select * from 体检管理_体检附加信息表 where 系统编号 = '" & lobjRec("系统编号").Value & "' and " & _
                      "附加项目 = '" & lobjRec("附加项目").Value & "') "
            lstrSQL = lstrSQL & "insert into 体检管理_体检附加信息表(系统编号,附加项目,项目值,项目值编号) " & _
                "values('" & lobjRec("系统编号").Value & "','" & lobjRec("附加项目").Value & _
                "','" & lobjRec("项目值").Value & "','" & lobjRec("项目值编号").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '数据导出前,先根据过滤条件从SQLserver取出需要导入的数据
        lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
        Set lobjRec = dafuncGetData(lstrSQL)
             
        '把数据依次添加到MDB数据库
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检附加信息表(系统编号,附加项目,项目值,项目值编号) " & _
                    "values('" & lobjRec("系统编号").Value & "','" & lobjRec("附加项目").Value & _
                    "','" & lobjRec("项目值").Value & "','" & lobjRec("项目值编号").Value & "')"
            
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub传输体检附加信息表", Err.Number, Err.Description, True
End Sub

'功能： 导入或导出"体检管理_体检结果信息表" 的内容
'输入： paraRange As Collection [数据范围名，数据范围值]
'作者： 刘浩。
Public Sub sub传输体检结果信息表(paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据(过滤掉了不需要被导入或导出的数据)
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先根据过滤条件从MDB取出需要导入的数据
        lstrSQL = "select 体检管理_体检结果信息表.* from 体检管理_体检结果信息表,体检管理_体检基本数据库 " & _
            "Where 体检管理_体检结果信息表.系统编号 = 体检管理_体检基本数据库.系统编号 "
        lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
        
        Set lobjRec = funcExecuteSql(lstrSQL)
                
        While Not lobjRec.EOF
            '只导入不存在的体检结果信息。
            lstrSQL = "if not exists(select * from 体检管理_体检结果信息表 " & _
                    "where 系统编号 = '" & lobjRec("系统编号").Value & "' " & _
                        "and 体检项目 = '" & lobjRec("体检项目").Value & "') "
            lstrSQL = lstrSQL & "insert into 体检管理_体检结果信息表(系统编号,体检项目,体检结果,体检医师,填写日期) " & _
                "values('" & lobjRec("系统编号").Value & "','" & _
                lobjRec("体检项目").Value & "','" & _
                lobjRec("体检结果").Value & "'," & _
                IIf(IsNull(lobjRec("体检医师").Value), "null", "'" & lobjRec("体检医师").Value & "'") & "," & _
                IIf(IsNull(lobjRec("填写日期").Value), "NULL", "'" & lobjRec("填写日期").Value & "'") & ")"
            
            dafuncGetData lstrSQL
            
            lobjRec.MoveNext
        Wend
    Else
        '数据导出前,先根据过滤条件从SQLserver取出需要导入的数据
        lstrSQL = "select 体检管理_体检结果信息表.* from 体检管理_体检结果信息表,体检管理_体检基本数据库 " & _
            "Where 体检管理_体检结果信息表.系统编号 = 体检管理_体检基本数据库.系统编号 "
        lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
        Set lobjRec = dafuncGetData(lstrSQL)
        
        '把数据依次添加到MDB数据库
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检结果信息表(系统编号,体检项目,体检结果,体检医师,填写日期) " & _
                    "values('" & lobjRec("系统编号").Value & "','" & _
                    lobjRec("体检项目").Value & "','" & _
                    lobjRec("体检结果").Value & "'," & _
                    IIf(IsNull(lobjRec("体检医师").Value), "null", "'" & lobjRec("体检医师").Value & "'") & "," & _
                    IIf(IsNull(lobjRec("填写日期").Value), "NULL", "#" & lobjRec("填写日期").Value & "#") & ")"
            
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub传输体检结果信息表", Err.Number, Err.Description, True
End Sub

'功能： 导入或导出"体检管理_体检表模板基本信息表" 的内容
'作者： 刘浩。
Public Sub sub传输体检表模板基本信息表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object      '用于判断目地数据库中是否已存在当前将要被导入的记录
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检表模板基本信息表 "
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果需要导入的数据已存在,修改原有记录.否则添加一条新记录(本表中的记录有被其它表引用的情况)
            lstrSQL = "select * from 体检管理_体检表模板基本信息表 where 体检表名称 = '" & lobjRec("体检表名称").Value & "'"
            Set lobjRec1 = dafuncGetData(lstrSQL)
            If lobjRec1.RecordCount <> 0 Then
                lstrSQL = "update 体检管理_体检表模板基本信息表 " & _
                    "set 试管编号字母 = '" & lobjRec("试管编号字母").Value & "' " & _
                    ",诊断处理意见 = '" & lobjRec("诊断处理意见").Value & "' " & _
                    ",体检单名称 = '" & lobjRec("体检单名称").Value & "' " & _
                    ",是否复查体检表 =" & IIf(IsNull(lobjRec("是否复查体检表")), 0, lobjRec("是否复查体检表").Value) & _
                    ",代号 = '" & lobjRec("代号").Value & "' " & _
                    ",收费标准 = '" & lobjRec("收费标准").Value & "' where 体检表名称 = '" & lobjRec("体检表名称").Value & "'"
            Else
                lstrSQL = "insert into 体检管理_体检表模板基本信息表(体检表名称,试管编号字母,诊断处理意见," & _
                        "体检单名称,是否复查体检表,代号,收费标准) " & _
                    "values('" & lobjRec("体检表名称").Value & "','" & lobjRec("试管编号字母").Value & "','" & _
                    lobjRec("诊断处理意见").Value & "','" & lobjRec("体检单名称").Value & "','" & _
                    lobjRec("是否复查体检表").Value & "','" & lobjRec("代号").Value & "','" & _
                    lobjRec("收费标准").Value & "')"
            End If
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检表模板基本信息表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        '把数据依次添加到MDB数据库
        While Not lobjRec.EOF
                lstrSQL = "insert into 体检管理_体检表模板基本信息表(体检表名称,试管编号字母,诊断处理意见," & _
                        "体检单名称,是否复查体检表,代号,收费标准) " & _
                    "values('" & lobjRec("体检表名称").Value & "','" & lobjRec("试管编号字母").Value & "','" & _
                    lobjRec("诊断处理意见").Value & "','" & lobjRec("体检单名称").Value & "'," & _
                    IIf(IsNull(lobjRec("是否复查体检表")), 0, lobjRec("是否复查体检表").Value) & ",'" & lobjRec("代号").Value & "','" & _
                    lobjRec("收费标准").Value & "')"
            
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输体检表模板基本信息表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_化验单基本信息表" 的内容
'输入：paraRange As Collection[数据范围名,数据范围值]
'作者：刘浩。
Public Sub sub传输化验单基本信息表(ByVal paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object      '用于判断目地数据库中是否已存在当前将要被导入的记录
    On Error GoTo errHandler
    
    lstrSQL = "select distinct 体检管理_化验单基本信息表.* from 体检管理_化验单基本信息表,体检管理_化验单试管编号数据库,体检管理_体检基本数据库 " & _
            "where 体检管理_化验单基本信息表.化验单编号 = 体检管理_化验单试管编号数据库.化验单编号 and " & _
            "体检管理_化验单试管编号数据库.试管编号 = 体检管理_体检基本数据库.试管编号 and " & _
            "体检管理_化验单试管编号数据库.化验日期 = 体检管理_体检基本数据库.体检日期 "
    lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
    If mblnInputOrOutput Then
        '导入数据前.  先MDB取出需要导入的数据
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果需要导入的数据已存在,修改原有记录.否则添加一条新记录(本表中的记录有被其它表引用的情况)
            lstrSQL = "if not exists(select * from 体检管理_化验单基本信息表 where 化验单编号 = '" & lobjRec("化验单编号").Value & "') "
            lstrSQL = lstrSQL & "insert into 体检管理_化验单基本信息表(化验单编号,试管编号范围,化验项目,化验日期,录入日期," & _
                    "化验人,复核人,复核日期,复核错误信息,化验单状态) " & _
                    "values('" & lobjRec("化验单编号").Value & "','" & lobjRec("试管编号范围").Value & "','" & _
                    lobjRec("化验项目").Value & "'," & _
                    IIf(IsNull(lobjRec("化验日期").Value), "NULL", "'" & lobjRec("化验日期").Value & "'") & "," & _
                    IIf(IsNull(lobjRec("录入日期").Value), "NULL", "'" & lobjRec("录入日期").Value & "'") & "," & _
                    IIf(IsNull(lobjRec("化验人").Value), "null", "'" & lobjRec("化验人").Value & "'") & "," & _
                    IIf(IsNull(lobjRec("复核人").Value), "null", "'" & lobjRec("复核人").Value & "'") & "," & _
                    IIf(IsNull(lobjRec("复核日期").Value), "NULL", "'" & lobjRec("复核日期").Value & "'") & ",'" & _
                    lobjRec("复核错误信息").Value & "','" & _
                    lobjRec("化验单状态").Value & "')"
            
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先SQLserver取出需要导入的数据
        Set lobjRec = dafuncGetData(lstrSQL)
        
        '把数据依次添加到MDB数据库
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_化验单基本信息表(化验单编号,试管编号范围,化验项目,化验日期,录入日期," & _
                    "化验人,复核人,复核日期,复核错误信息,化验单状态) " & _
                    "values('" & lobjRec("化验单编号").Value & "','" & lobjRec("试管编号范围").Value & "','" & _
                    lobjRec("化验项目").Value & "'," & _
                    IIf(IsNull(lobjRec("化验日期").Value), "NULL", "#" & lobjRec("化验日期").Value & "#") & "," & _
                    IIf(IsNull(lobjRec("录入日期").Value), "NULL", "#" & lobjRec("录入日期").Value & "#") & "," & _
                    IIf(IsNull(lobjRec("化验人").Value), "null", "'" & lobjRec("化验人").Value & "'") & "," & _
                    IIf(IsNull(lobjRec("复核人").Value), "null", "'" & lobjRec("复核人").Value & "'") & "," & _
                    IIf(IsNull(lobjRec("复核日期").Value), "NULL", "#" & lobjRec("复核日期").Value & "#") & ",'" & _
                    lobjRec("复核错误信息").Value & "','" & _
                    lobjRec("化验单状态").Value & "')"
                    
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输化验单基本信息表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_化验单不正常结果信息表" 的内容
'输入：paraRange as collection数据导入或导出的过滤条件。[数据范围名，数据范围值]
'作者：刘浩。
Public Sub sub传输化验单不正常结果信息表(paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    lstrSQL = "select distinct 体检管理_化验单不正常结果信息表.* from 体检管理_化验单不正常结果信息表,体检管理_化验单试管编号数据库,体检管理_体检基本数据库 where " & _
            "体检管理_化验单不正常结果信息表.化验单编号 = 体检管理_化验单试管编号数据库.化验单编号 and " & _
            "体检管理_化验单试管编号数据库.试管编号 = 体检管理_体检基本数据库.试管编号 and " & _
            "体检管理_体检基本数据库.体检日期 = 体检管理_化验单试管编号数据库.化验日期 "
    lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
    If mblnInputOrOutput Then
        '导入数据前.  先MDB取出需要导入的数据
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '只导入不存在的记录。
            lstrSQL = "if not exists(select * from 体检管理_化验单不正常结果信息表 where 化验单编号 = '" & lobjRec("化验单编号").Value & "' and " & _
                        " 序号 = " & lobjRec("序号").Value & ") "
            lstrSQL = lstrSQL & "insert into 体检管理_化验单不正常结果信息表(化验单编号,序号,化验项目,化验结果,试管编号范围) " & _
                        "values('" & lobjRec("化验单编号").Value & "'," & lobjRec("序号").Value & ",'" & _
                        lobjRec("化验项目").Value & "','" & lobjRec("化验结果").Value & "','" & _
                        lobjRec("试管编号范围").Value & "')"
            dafuncGetData lstrSQL
            
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.先从SQLserver取出需要导入的数据。
        Set lobjRec = dafuncGetData(lstrSQL)
        
        '把数据依次添加到MDB数据库。
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_化验单不正常结果信息表(化验单编号,序号,化验项目,化验结果,试管编号范围) " & _
                    "values('" & lobjRec("化验单编号").Value & "'," & lobjRec("序号").Value & ",'" & _
                    lobjRec("化验项目").Value & "','" & lobjRec("化验结果").Value & "','" & _
                    lobjRec("试管编号范围").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输化验单不正常结果信息表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_体检表模板体检项目表" 的内容。
'作者：刘浩
Public Sub sub传输体检表模板体检项目表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检表模板体检项目表 "
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete  from 体检管理_体检表模板体检项目表 where 体检表名称 = '" & _
                lobjRec("体检表名称").Value & "' and 体检项目 = '" & lobjRec("体检项目").Value & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_体检表模板体检项目表(体检表名称,体检项目) " & _
                    "values('" & lobjRec("体检表名称").Value & "','" & lobjRec("体检项目").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检表模板体检项目表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        '把数据依次添加到MDB数据库
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检表模板体检项目表(体检表名称,体检项目) " & _
                    "values('" & lobjRec("体检表名称").Value & "','" & lobjRec("体检项目").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "ssub传输体检表模板体检项目表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_体检表模板体检结论表" 的内容
'作者：刘浩。
Public Sub sub体检表模板体检结论表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检表模板体检结论表 "
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_体检表模板体检结论表 where 体检表名称 = '" & _
                lobjRec("体检表名称").Value & "' and  体检结论 = " & lobjRec("体检结论").Value
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_体检表模板体检结论表(体检表名称,体检结论) " & _
                    "values('" & lobjRec("体检表名称").Value & "'," & lobjRec("体检结论").Value & ")"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检表模板体检结论表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        '把数据依次添加到MDB数据库
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检表模板体检结论表(体检表名称,体检结论) " & _
                    "values('" & lobjRec("体检表名称").Value & "'," & lobjRec("体检结论").Value & ")"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "ssub传输体检表模板体检项目表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_体检项目设置表" 的内容。
'作者：刘浩
Public Sub sub传输体检项目设置表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检项目设置表 "
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果需要导入的数据已存在,修改原有记录.否则添加一条新记录(本表中的记录有被其它表引用的情况)
            lstrSQL = "select * from 体检管理_体检项目设置表 where 编码 = '" & lobjRec("编码").Value & "'"
            Set lobjRec1 = dafuncGetData(lstrSQL)
            If lobjRec1.RecordCount >= 1 Then
                lstrSQL = "update 体检管理_体检项目设置表 set " & _
                    "名称 = '" & lobjRec("名称").Value & _
                    "',缺省值 = '" & lobjRec("缺省值").Value & _
                    "',枚举来源 = '" & lobjRec("枚举来源").Value & _
                    "',属性 = '" & lobjRec("属性").Value & _
                    "',体检大类 = '" & lobjRec("体检大类").Value & "' where 编码 = '" & lobjRec("编码").Value & "'"
            Else
                lstrSQL = "insert into 体检管理_体检项目设置表(编码,名称,缺省值,枚举来源,属性," & _
                    "体检大类) " & _
                    "values('" & lobjRec("编码").Value & "','" & lobjRec("名称").Value & "','" & lobjRec("缺省值").Value & "','" & _
                    lobjRec("枚举来源").Value & "','" & lobjRec("属性").Value & "','" & _
                    lobjRec("体检大类").Value & "')"
            End If
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检项目设置表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检项目设置表(编码,名称,缺省值,枚举来源,属性," & _
                    "体检大类) " & _
                    "values('" & lobjRec("编码").Value & "','" & lobjRec("名称").Value & "','" & lobjRec("缺省值").Value & "','" & _
                    lobjRec("枚举来源").Value & "','" & lobjRec("属性").Value & "','" & _
                    lobjRec("体检大类").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输体检项目设置表", Err.Number, Err.Description, True
End Sub


'功能：导入或导出"体检管理_体检医师项目设置表" 的内容。
'作者：刘浩。
Public Sub sub传输体检医师项目设置表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检医师项目设置表 "
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_体检医师项目设置表 where 医师编号 = '" & _
                lobjRec("医师编号").Value & "' and 体检项目 = '" & lobjRec("体检项目").Value & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_体检医师项目设置表(医师编号,体检项目) " & _
                    "values('" & lobjRec("医师编号").Value & "','" & lobjRec("体检项目").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检医师项目设置表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检医师项目设置表(医师编号,体检项目) " & _
                    "values('" & lobjRec("医师编号").Value & "','" & lobjRec("体检项目").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输体检医师项目设置表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_体检结论判断条件表" 的内容。
'作者：刘浩
Public Sub sub传输体检结论判断条件表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检结论判断条件表 "
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_体检结论判断条件表 where 编号 = '" & _
                lobjRec("编号").Value & "' and 体检结论 = '" & lobjRec("体检结论").Value & "' and 序号 = '" & _
                lobjRec("序号").Value & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_体检结论判断条件表(编号,体检结论,描述,序号,体检项目,判断条件,标准值) " & _
                    "values('" & lobjRec("编号").Value & "','" & _
                    lobjRec("体检结论").Value & "','" & lobjRec("描述").Value & "','" & _
                    lobjRec("序号").Value & "','" & lobjRec("体检项目").Value & "','" & _
                    lobjRec("判断条件").Value & "','" & lobjRec("标准值").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检结论判断条件表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检结论判断条件表(编号,体检结论,描述,序号,体检项目,判断条件,标准值) " & _
                    "values('" & lobjRec("编号").Value & "','" & _
                    lobjRec("体检结论").Value & "','" & lobjRec("描述").Value & "','" & _
                    lobjRec("序号").Value & "','" & lobjRec("体检项目").Value & "','" & _
                    lobjRec("判断条件").Value & "','" & lobjRec("标准值").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输体检医师项目设置表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_试管编号最大流水号表" 的内容。
'作者：刘浩
Public Sub sub传输试管编号最大流水号表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检项目设置表 "
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
             '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_试管编号最大流水号表 where 字母 = '" & lobjRec("字母").Value & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_试管编号最大流水号表(字母,当前最大流水号,日期) " & _
                    "values('" & lobjRec("字母").Value & "','" & lobjRec("当前最大流水号").Value & "'," & _
                    IIf(IsNull(lobjRec("日期").Value), "NULL", Format(lobjRec("日期").Value, "yyyy-mm-dd")) & ")"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_试管编号最大流水号表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_试管编号最大流水号表(字母,当前最大流水号,日期) " & _
                    "values('" & lobjRec("字母").Value & "','" & lobjRec("当前最大流水号").Value & "'," & _
                    IIf(IsNull(lobjRec("日期").Value), "NULL", "#" & Format(lobjRec("日期").Value, "yyyy-mm-dd") & "#") & ")"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输试管编号最大流水号表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_业务设置信息表" 的内容。
'作者：刘浩。
Public Sub sub传输业务设置信息表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_业务设置信息表"
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_业务设置信息表 where 设置项目 = '" & lobjRec("设置项目").Value & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_业务设置信息表(设置项目,设置值,枚举来源,说明) " & _
                    "values('" & lobjRec("设置项目").Value & "','" & lobjRec("设置值").Value & "','" & lobjRec("枚举来源").Value & "','" & _
                    lobjRec("说明").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_业务设置信息表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_业务设置信息表(设置项目,设置值,枚举来源,说明) " & _
                    "values('" & lobjRec("设置项目").Value & "','" & lobjRec("设置值").Value & "','" & _
                    lobjRec("枚举来源").Value & "','" & _
                    lobjRec("说明").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输业务设置信息表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_试管编号字母表" 的内容。
'作者：刘浩。
Public Sub sub传输试管编号字母表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_试管编号字母表"
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_试管编号字母表 where 体检表名称 = '" & lobjRec("体检表名称").Value & "'"
            dafuncGetData lstrSQL
             lstrSQL = "insert  into 体检管理_试管编号字母表(体检表名称,字母,日期,是否固定) " & _
                    "values('" & lobjRec("体检表名称").Value & "','" & lobjRec("字母").Value & "'," & _
                    IIf(IsNull(lobjRec("日期").Value), "NULL", Format(lobjRec("日期").Value, "yyyy-mm-dd")) & "," & _
                    lobjRec("是否固定").Value & ")"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_试管编号字母表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert  into 体检管理_试管编号字母表(体检表名称,字母,日期,是否固定) " & _
                    "values('" & lobjRec("体检表名称").Value & "','" & lobjRec("字母").Value & "'," & _
                    IIf(IsNull(lobjRec("日期").Value), "NULL", "#" & Format(lobjRec("日期").Value, "yyyy-mm-dd") & "#") & "," & _
                    lobjRec("是否固定").Value & ")"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输试管编号字母表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_试管编号字母表" 的内容
'作者：刘浩。
Public Sub sub传输编号生成规则表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_编号生成规则表"
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_编号生成规则表 where 编号名称 = '" & _
                lobjRec("编号名称").Value & "' and 分段号 = " & lobjRec("分段号").Value
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_编号生成规则表(编号名称,分段号,组成,该段是否允许用户修改,分段长度) " & _
                    "values('" & lobjRec("编号名称").Value & "','" & lobjRec("分段号").Value & "','" & _
                    lobjRec("组成").Value & "','" & _
                    lobjRec("该段是否允许用户修改").Value & "','" & _
                    lobjRec("分段长度").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导入的数据
        lstrSQL = "select * from 体检管理_编号生成规则表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_编号生成规则表(编号名称,分段号,组成,该段是否允许用户修改,分段长度) " & _
                    "values('" & lobjRec("编号名称").Value & "','" & lobjRec("分段号").Value & "','" & _
                    lobjRec("组成").Value & "','" & _
                    lobjRec("该段是否允许用户修改").Value & "','" & _
                    lobjRec("分段长度").Value & "')"
            subExecuteSql lstrSQL
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输编号生成规则表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_编号最大流水号表" 的内容。
'作者：刘浩。
Public Sub sub传输编号最大流水号表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前，  先从MDB取出需要导入的数据。
        lstrSQL = "select * from 体检管理_编号最大流水号表"
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在，先删除原有记录。(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from 体检管理_编号最大流水号表 where 编号名称 = '" & _
                lobjRec("编号名称").Value & "' and datepart(yy,日期) = '" & _
                DatePart("yyyy", lobjRec("日期").Value) & "' and datepart(mm,日期) = '" & _
                DatePart("m", lobjRec("日期").Value) & "' and datepart(dd,日期) = '" & _
                DatePart("d", lobjRec("日期").Value) & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert into 体检管理_编号最大流水号表(编号名称,最大流水号,日期) " & _
                    "values('" & lobjRec("编号名称").Value & "','" & _
                    lobjRec("最大流水号").Value & "','" & _
                    lobjRec("日期").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前，  先从SQLserver取出需要导出的数据。
        lstrSQL = "select * from 体检管理_编号最大流水号表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_编号最大流水号表(编号名称,最大流水号,日期) " & _
                    "values('" & lobjRec("编号名称").Value & "','" & _
                    lobjRec("最大流水号").Value & "',#" & _
                    lobjRec("日期").Value & "#)"
            subExecuteSql lstrSQL

            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输编号最大流水号表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_体检人员附加项目设置表" 的内容。
'作者：刘浩。
Public Sub sub传输体检人员附加项目设置表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select * from 体检管理_体检人员附加项目设置表"
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,修改原有记录,否则添加新记录(本表的记录存在被其它表引用的情况)
            lstrSQL = "select *  from 体检管理_体检人员附加项目设置表 where 附加项目 = '" & lobjRec("附加项目").Value & "'"
            Set lobjRec1 = dafuncGetData(lstrSQL)
            If lobjRec1.RecordCount >= 1 Then
'                lstrSQL = "update 体检管理_体检人员附加项目设置表 set " & _
                    "录入标题 = '" & lobjRec("录入标题").Value & _
                    "',数据类型 = '" & lobjRec("数据类型").Value & _
                    "',数据长度 = '" & lobjRec("数据长度").Value & _
                    "',枚举值 = '" & lobjRec("枚举值").Value & "' where 附加项目 = '" & lobjRec("附加项目").Value & "'"
            Else
                lstrSQL = "insert into 体检管理_体检人员附加项目设置表(附加项目,录入标题,数据类型,数据长度,枚举值) " & _
                    "values('" & lobjRec("附加项目").Value & "','" & lobjRec("录入标题").Value & "','" & _
                    lobjRec("数据类型").Value & "','" & _
                    lobjRec("数据长度").Value & "','" & lobjRec("枚举值") & "')"
                
                dafuncGetData lstrSQL
            End If
            
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导出的数据
        lstrSQL = "select * from 体检管理_体检人员附加项目设置表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检人员附加项目设置表(附加项目,录入标题,数据类型,数据长度,枚举值) " & _
                    "values('" & lobjRec("附加项目").Value & "','" & lobjRec("录入标题").Value & "','" & _
                    lobjRec("数据类型").Value & "','" & _
                    lobjRec("数据长度").Value & "','" & lobjRec("枚举值") & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输体检人员附加项目设置表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_体检访问标志表" 的内容。
'输入：paraRange as collection[数据范围名，数据范围值]。
'作者：刘浩。
Public Sub sub传输体检访问标志表(paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    Dim lobjRec1 As Object
    On Error GoTo errHandler
    
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        lstrSQL = "select 体检管理_体检访问标志表.* from 体检管理_体检访问标志表,体检管理_体检基本数据库 " & _
                " where 体检管理_体检访问标志表.系统编号 = 体检管理_体检基本数据库.系统编号 "
        lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "if not exists(select *  from 体检管理_体检访问标志表 where 系统编号 = '" & lobjRec("系统编号").Value & "') "
            lstrSQL = lstrSQL & "insert 体检管理_体检访问标志表(系统编号,健康证,劳动卫生,放射卫生,学校卫生) " & _
                "values('" & lobjRec("系统编号").Value & "','" & lobjRec("健康证").Value & "','" & _
                lobjRec("劳动卫生").Value & "','" & _
                lobjRec("放射卫生").Value & "','" & lobjRec("学校卫生") & "')"
        
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导出的数据
        lstrSQL = "select * from 体检管理_体检访问标志表 "
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_体检访问标志表(系统编号,健康证,劳动卫生,放射卫生,学校卫生) " & _
                    "values('" & lobjRec("系统编号").Value & "','" & lobjRec("健康证").Value & "','" & _
                    lobjRec("劳动卫生").Value & "','" & _
                    lobjRec("放射卫生").Value & "','" & lobjRec("学校卫生") & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输体检访问标志表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_化验单试管编号表" 的内容
'输入：paraRange as collection[数据范围名,数据范围值]
'作者：刘浩。
Public Sub sub传输化验单试管编号表(ByVal paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    lstrSQL = "select distinct 体检管理_化验单试管编号表.* from 体检管理_化验单试管编号表,体检管理_体检基本数据库,体检管理_化验单试管编号数据库 where " & _
        "体检管理_化验单试管编号数据库.试管编号 = 体检管理_体检基本数据库.试管编号 and " & _
        "体检管理_化验单试管编号数据库.化验日期 = 体检管理_体检基本数据库.体检日期 and " & _
        "体检管理_化验单试管编号数据库.化验单编号 = 体检管理_化验单试管编号表.化验单编号 "
    lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
            '如果要导出的数据已经存在,先删除原有记录.(本表的记录不存在被其它表引用的情况)
            lstrSQL = "delete from  体检管理_化验单试管编号表 where 化验单编号 = '" & _
                lobjRec("化验单编号").Value & "' and 试管编号 = '" & lobjRec("试管编号").Value & "'"
            dafuncGetData lstrSQL
            lstrSQL = "insert 体检管理_化验单试管编号表(化验单编号,试管编号) " & _
                    "values('" & lobjRec("化验单编号").Value & "','" & lobjRec("试管编号").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导出的数据
        Set lobjRec = dafuncGetData(lstrSQL)
        
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_化验单试管编号表(化验单编号,试管编号) " & _
                    "values('" & lobjRec("化验单编号").Value & "','" & lobjRec("试管编号").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输化验单试管编号表", Err.Number, Err.Description, True
    Exit Sub
End Sub

'功能：导入或导出"体检管理_化验单不正常结果试管编号表" 的内容
'输入：paraRange as collection[数据范围名,数据范围值]
'作者：刘浩。
Public Sub sub传输化验单不正常结果试管编号表(ByVal paraRange As Collection)
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    On Error GoTo errHandler
    
    lstrSQL = "select distinct 体检管理_化验单不正常结果试管编号表.* " & _
        "from 体检管理_化验单不正常结果试管编号表,体检管理_体检基本数据库,体检管理_化验单试管编号数据库 where " & _
        "体检管理_化验单试管编号数据库.化验单编号 = 体检管理_化验单不正常结果试管编号表.化验单编号 and " & _
        "体检管理_化验单不正常结果试管编号表.试管编号 = 体检管理_体检基本数据库.试管编号 and " & _
        "体检管理_化验单试管编号数据库.化验日期 = 体检管理_体检基本数据库.体检日期 "
    lstrSQL = lstrSQL & funcFilter(paraRange, mblnInputOrOutput)
    If mblnInputOrOutput Then
        '导入数据前.  先从MDB取出需要导入的数据
        Set lobjRec = funcExecuteSql(lstrSQL)
        While Not lobjRec.EOF
            '先删除原有记录
            lstrSQL = "delete from 体检管理_化验单不正常结果试管编号表 where " & _
                "化验单编号 = '" & lobjRec("化验单编号").Value & _
                "' and 序号 = '" & lobjRec("序号").Value & _
                "' and 试管编号 = '" & lobjRec("试管编号").Value & "'"
            dafuncGetData lstrSQL
        
            lstrSQL = "insert into 体检管理_化验单不正常结果试管编号表(化验单编号,序号,试管编号)" & "values('" & _
                lobjRec("化验单编号").Value & "','" & _
                lobjRec("序号").Value & "','" & _
                lobjRec("试管编号").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
    Else
        '导出数据前.  先从SQLserver取出需要导出的数据
        Set lobjRec = dafuncGetData(lstrSQL)
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_化验单不正常结果试管编号表(化验单编号,序号,试管编号)" & "values('" & _
                lobjRec("化验单编号").Value & "','" & _
                lobjRec("序号").Value & "','" & _
                lobjRec("试管编号").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输化验单试管编号表", Err.Number, Err.Description, True
End Sub

'功能：导入或导出"体检管理_数据分类信息表" 的内容
'输入：paraRange as collection[数据范围名,数据范围值]
'作者：刘浩。
Public Sub sub传输数据分类信息表()
    Dim lobjRec As Object       '从源数据库中读出的数据
    Dim lstrSQL As String       '拼出的SQL语句
    
    On Error GoTo errHandler
    
    lstrSQL = "select * from 体检管理_数据分类信息表"
    If mblnInputOrOutput Then
        '先删除原有记录。
        dafuncGetData "delete from 体检管理_数据分类信息表"
        
        On Error Resume Next
        dafuncGetData "set IDENTITY_INSERT 体检管理_数据分类信息表 ON"
        On Error GoTo errHandler
        
        '导入数据前.  先从MDB取出需要导入的数据
        Set lobjRec = funcExecuteSql(lstrSQL)
        
        While Not lobjRec.EOF
        
            lstrSQL = "insert into 体检管理_数据分类信息表(编号,数据分类名,基础表名,服务器类型)" & "values('" & _
                lobjRec("编号").Value & "','" & _
                lobjRec("数据分类名").Value & "','" & _
                lobjRec("基础表名").Value & "','" & _
                lobjRec("服务器类型").Value & "')"
            dafuncGetData lstrSQL
            lobjRec.MoveNext
        Wend
        On Error Resume Next
        dafuncGetData "set IDENTITY_INSERT 体检管理_数据分类信息表 OFF"
        
    Else
        '导出数据前.  先从SQLserver取出需要导出的数据
        Set lobjRec = dafuncGetData(lstrSQL)
        While Not lobjRec.EOF
            lstrSQL = "insert into 体检管理_数据分类信息表(编号,数据分类名,基础表名,服务器类型)" & "values('" & _
                 lobjRec("编号").Value & "','" & _
                lobjRec("数据分类名").Value & "','" & _
                lobjRec("基础表名").Value & "','" & _
                lobjRec("服务器类型").Value & "')"
            subExecuteSql lstrSQL
            
            lobjRec.MoveNext
        Wend
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub传输数据分类信息表", Err.Number, Err.Description, True
    
End Sub


'功能：使用dao把当前连接的sqlserver库上上指定表内容导入当前mdb文件中。
'输入：paraTableName  表名。
'作者：杨春。
Public Sub sub导出sql表到mdb文件(ByVal paraTableName As String, ByVal paraFilter As String)
    Dim ldb As Database
    Dim lstrSQL As String
    On Error GoTo errHandler
    Set ldb = Workspaces(0).OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DATABASE=" & mstrData & ";SERVER=" & mstrServer & ";UID=user26;PWD=welcome;DSN=WSFY2001")

    '导出数据前.  先从SQLserver取出需要导入的数据
    lstrSQL = "insert into [" & mstrMdbFileName & "]." & paraTableName & " select * from " & paraTableName & IIf(paraFilter = "", "", " where " & paraFilter)
    ldb.Execute (lstrSQL)
    
    ldb.Close
    Exit Sub
errHandler:
    If Not ldb Is Nothing Then
        ldb.Close
    End If
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "sub导出sql表到mdb文件", Err.Number, Err.Description, True
End Sub

'功能：使用DAO把一张ACCESS表的数据导入到SQL SERVER的表中
'输入：表名,条件。
'作者：杨春
Public Sub sub从mdb导入表内容到sql(ByVal para表名 As String, ByVal paraFilter As String)
    Dim ldb As Database
    Dim lobjTemp As Object
    Dim lstrSQL As String
    Dim lstrTemp As String
    Dim i As Integer
    
    On Error GoTo errHandle
    Set ldb = Workspaces(0).OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DATABASE=" & mstrData & ";SERVER=" & mstrServer & ";UID=user26;PWD=welcome;DSN=WSFY2001")
    Set lobjTemp = ldb.OpenRecordset("select * from [" & mstrMdbFileName & "]." & para表名 & "")
    If lobjTemp.RecordCount > 0 Then
        '有记录
        lstrTemp = ""
        lstrSQL = "select b.name from sysindexkeys a,syscolumns b Where a.id = b.id and a.id=object_id(N'" & para表名 & "') and a.colid=b.colid"
        Set lobjTemp = dafuncGetData(lstrSQL)
        If lobjTemp.RecordCount > 0 Then
            '该表有主键
            lstrSQL = "insert into " & para表名 & " select * from [" & mstrMdbFileName & "]." & para表名 & " a where " & IIf(paraFilter = "", "", paraFilter & " and ") & "not exists(select * from " & para表名 & " where  "
            For i = 1 To lobjTemp.RecordCount
                lstrTemp = lstrTemp & "" & para表名 & "." & lobjTemp.Fields("name") & "= a." & lobjTemp.Fields("name") & " and "
                lobjTemp.MoveNext
            Next i
            lstrTemp = Left(lstrTemp, Len(lstrTemp) - 4) & ")"
            
        Else  '没有主键
            Err.Raise 6666, , "表“" & para表名 & "”没有主键，不能导入。"
        End If
        
        lstrSQL = lstrSQL & lstrTemp
        ldb.Execute lstrSQL
        
    End If
    ldb.Close
    Exit Sub
    
errHandle:
    ldb.Close
    Call sfsub错误处理("体检对外接口", "clsMdbFile", "sub从mdb导入表内容到sql", Err.Number, Err.Description, True)
End Sub

