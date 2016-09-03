VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsManageTransmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'属性变量。
Private mcol业务设置 As Collection              ' [设置项目，设置值，枚举来源，说明]key:设置项目。
Private mobj工作站配置 As ClsLocalConfigure     '工作站配置信息：工作模式。

Private mobjInMdbFile As clsMdbFile  '导入mdb文件。
Private mobjOutMdbFile As clsMdbFile '导出mdb文件。

Private mstrInMDBPath As String      '导入文件名（带路经）。
Private mstrOutMDBPath As String     '导出文件名（带路经）。
Private mstrOutTmpMdbFile As String  '临时导出文件名（带路经）。

'功能： 初始化类。
'作者： 刘浩。
Private Sub Class_Initialize()
    Dim lobjRec As Object
    Dim lcolItem As Collection
    Dim lstrMDBPath  As String
    
    On Error GoTo errHandler
    
    '从表“体检管理_业务设置信息表”获取业务设置：设置项目，设置值，枚举来源，说明。
    Set mcol业务设置 = New Collection
    Set lobjRec = dafuncGetData("select * from 体检管理_业务设置信息表")
    Do While Not lobjRec.EOF
        Set lcolItem = New Collection
        lcolItem.Add lobjRec("设置项目").Value, "设置项目"
        lcolItem.Add IIf(IsNull(lobjRec("设置值")), "", lobjRec("设置值").Value), "设置值"
        lcolItem.Add IIf(IsNull(lobjRec("枚举来源")), "", lobjRec("枚举来源").Value), "枚举来源"
        lcolItem.Add IIf(IsNull(lobjRec("说明")), "", lobjRec("说明").Value), "说明"
        
        mcol业务设置.Add lcolItem, lcolItem("设置项目")
        lobjRec.MoveNext
    Loop
    lobjRec.Close
    
    '创建工作站配置对象。
    Set mobj工作站配置 = New ClsLocalConfigure
    
    '创建导入导出mdb文件对象。
    Set mobjInMdbFile = New clsMdbFile
    mobjInMdbFile.InputOrOutput = True
    
    Set mobjOutMdbFile = New clsMdbFile
    mobjOutMdbFile.InputOrOutput = False
    
    '设置临时导出文件名（带路经）。
    mstrOutTmpMdbFile = IIf(Right(App.Path, 1) = "\", App.Path & "数据传输导出数据库.mdb", App.Path & "\数据传输导出数据库.mdb")
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageMedicalExam", "Class_Initialize", Err.Number, Err.Description, False
End Sub

'功能： 释放类引用的变量。
'作者： 刘浩。
Private Sub Class_Terminate()
    On Error Resume Next
    
    Set mobj工作站配置 = Nothing
    
    '关闭连接。
    Set mobjInMdbFile = Nothing
    Set mobjOutMdbFile = Nothing
End Sub

'功能：获取指定设置项目的内容。
'返回：设置值 as string
'作者： 刘浩。
Public Property Get 业务设置(ByVal para设置项目 As String) As String
    On Error GoTo errHandler
    业务设置 = mcol业务设置(para设置项目)("设置值")
    Exit Property
errHandler:
    业务设置 = ""
End Property

Public Property Get 工作站配置() As ClsLocalConfigure
    On Error GoTo errHandler
    Set 工作站配置 = mobj工作站配置
errHandler:
    '
End Property

'功能：启动本部件所包括的一级操作界面，
'      按照总体结构的要求，必须无模式启动界面，并指定其parent指向主导航界面，
'       并把界面移动到指定位置。
'输入：para操作名称      要启动的界面对应的操作名称。
'返回：操作对应的界面句柄。
'作者： 刘浩。
Public Function funcStart(ByVal para操作名称 As String) As Long
    Dim lfrmForm As Form
    On Error GoTo errHandler
    Select Case para操作名称
        Case "外单位数据导入", "体检管理_外单位数据导入"
            Set lfrmForm = frmInputDataFromExcel
        Case "内部数据导入", "体检管理_内部数据导入"
            Set lfrmForm = frmInputDataFromMdb
        Case "内部数据导出", "体检管理_内部数据导出"
            Set lfrmForm = frmOutputData
        Case Else
            funcStart = -1
            Err.Raise 6666, , "当前体检管理版本不支持“" & para操作名称 & "”操作。"
    End Select
    '如果找到窗体则打开窗体并调整窗体位置，返回窗体句柄。
    If funcStart <> -1 Then
        If lfrmForm.pblnInUse Then
            funcStart = -2
        Else
            funcStart = lfrmForm.hWnd
        End If
        
        lfrmForm.Move 700, 350
'        lfrmForm.ZOrder 0
        lfrmForm.Show
        
        Set lfrmForm = Nothing
    End If
    Exit Function
errHandler:
    Set lfrmForm = Nothing
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "funcStart", Err.numer, Err.Description, True
End Function

' 功能：    将操作名所对应的窗体关闭，所有界面模板必须提供此方法。
' 输入：    para操作名
' 返回：    是否成功关闭窗体。
' 作者：    刘浩。
Public Function funcClose(para操作名 As String) As Boolean
Attribute funcClose.VB_Description = "'功能: 从paraRange[数据范围名，数据范围值] 集合的值分析出导入或导出的条件,并拼成一个SQL的条件语句.\r\n    '输入: paraRange As Collection [数据范围名，数据范围值]，数据范围名：起始日期，开始日期，单位名称集，系统编号范围。\r\n    '返回值 : 以and 开始的string."
    On Error GoTo errHandler
    Dim lfrmForm As Form '待关闭的窗体
    
    funcClose = False
    '根据操作名设置不同的待关闭窗体。
    Select Case para操作名
        Case "外单位数据导入", "体检管理_外单位数据导入"
            Set lfrmForm = frmInputDataFromExcel
        Case "内部数据导入", "体检管理_内部数据导入"
            Set lfrmForm = frmInputDataFromMdb
        Case "内部数据导出", "体检管理_内部数据导出"
            Set lfrmForm = frmOutputData
        Case Else
            funcClose = True
    End Select
    '如果找到窗体则关闭。
    If funcClose = False Then
        Unload lfrmForm
        '如果窗体未确认退出则返回False。
        funcClose = Not lfrmForm.pblnInUse
    End If
    Set lfrmForm = Nothing
    Exit Function
    
errHandler:
    Set lfrmForm = Nothing
    sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "funcClose", Err.Number, Err.Description, True
End Function

'功能：为主导航界面获取主推信息。
'输入：需要获取主推信息的操作名称。
'作者： 刘浩。
Public Function Func获取主推信息(ByVal para操作名称 As String) As String
    Func获取主推信息 = ""
End Function

'功能：获取满足输入参数并使用指定体检表的体检人员的已体检完毕（无需复查）的并且满足指定访问标志值的体检信息。
'输入：查询的过滤条件.
'返回：collection[系统编号，姓名，性别，年龄，健康档案编号，单位申请编号，单位名称，附加项目[项目名，项目值]，体检日期，
'        体检结论id，体检结论名称，诊断和处理意见，体检结果信息[体检项目编码，体检项目名称，体检结果]]
'作者：刘浩。
'修改：2001-12-29（体检表名称可以是多个，以英文逗号隔开）。
Public Function Func获取已完毕的体检信息(Optional ByVal para从系统编号 As String = "", _
                                        Optional ByVal para到系统编号 As String = "", _
                                        Optional ByVal para单位名称 As String = "", _
                                        Optional ByVal para姓名 As String = "", _
                                        Optional ByVal para身份证号 As String = "", _
                                        Optional ByVal para体检表名称 As String = "", _
                                        Optional ByVal para标志类型 As String = "", _
                                        Optional ByVal para标志值 As String = "") As Collection
    
    Dim lobjMain As Object  '体检管理_体检访问数据库中的满足参数的体检记录。
    Dim lobjRec As Object
    Dim lcolItem As Collection
    Dim lcolItem1 As Collection
    Dim lcolItem11 As Collection
    Dim lcolInfo As Collection
    Dim lstrFilter As String   '根据参数拼成的sql语句的条件。
    Dim lstrSQL As String
    Dim lstrList As String
    Dim lstrItem As String
    
    On Error GoTo errHandler
    '根据参数拼条件。
    lstrFilter = "体检状态=3 and isnull(复查体检表名,'')=''"
    If para从系统编号 <> "" Then lstrFilter = lstrFilter & " and 系统编号>='" & para从系统编号 & "'"
    If para到系统编号 <> "" Then lstrFilter = lstrFilter & " and 系统编号<='" & para到系统编号 & "'"
    If para单位名称 <> "" Then lstrFilter = lstrFilter & " and 单位名称 like '" & para单位名称 & "%'"
    If para姓名 <> "" Then lstrFilter = lstrFilter & " and 姓名 like '" & para姓名 & "%'"
    If para身份证号 <> "" Then lstrFilter = lstrFilter & " and 身份证号='" & para身份证号 & "'"
    
    If Trim(para体检表名称) <> "" Then
        lstrList = ""
        para体检表名称 = Trim(para体检表名称)
        If Right(para体检表名称, 1) <> "," Then
            para体检表名称 = para体检表名称 & ","
        End If
        lstrItem = Left(para体检表名称, InStr(para体检表名称, ",") - 1)
        Do While lstrItem <> ""
            lstrList = lstrList & "'" & lstrItem & "',"
            para体检表名称 = Right(para体检表名称, Len(para体检表名称) - InStr(para体检表名称, ","))
            If para体检表名称 <> "" Then
                lstrItem = Left(para体检表名称, InStr(para体检表名称, ",") - 1)
            Else
                lstrItem = ""
            End If
        Loop
        lstrList = Left(lstrList, Len(lstrList) - 1)
        lstrFilter = lstrFilter & " and 体检表名称 in (" & lstrList & ")"
    End If
    If para标志类型 <> "" Then
        lstrFilter = lstrFilter & " and isnull(" & para标志类型 & ",'')='" & para标志值 & "'"
    End If
        
    '获取满足条件的体检记录。
    lstrSQL = "select * from 体检管理_体检访问数据库 where " & lstrFilter
    Set lobjMain = dafuncGetData(lstrSQL)
    Set lcolInfo = New Collection
    While Not lobjMain.EOF
        '返回值的一级子集合
        Set lcolItem = New Collection
        With lcolItem
            .Add lobjMain("系统编号").Value, "系统编号"
            .Add lobjMain("健康档案编号").Value, "健康档案编号"
            .Add lobjMain("姓名").Value, "姓名"
            .Add lobjMain("性别").Value, "性别"
            If IsDate(IIf(IsNull(lobjMain("出生日期").Value), "", lobjMain("出生日期").Value)) Then
                .Add DateDiff("yyyy", lobjMain("出生日期").Value, Date), "年龄"
            Else
                .Add 0, "年龄"
            End If
            .Add IIf(IsNull(lobjMain("单位申请编号")), "", lobjMain("单位申请编号").Value), "单位申请编号"
            .Add IIf(IsNull(lobjMain("单位名称")), "", lobjMain("单位名称").Value), "单位名称"
            .Add lobjMain("体检日期").Value, "体检日期"
            .Add IIf(IsNull(lobjMain("体检结论")), "", lobjMain("体检结论").Value), "体检结论"
            .Add IIf(IsNull(lobjMain("诊断和处理意见")), "", lobjMain("诊断和处理意见").Value), "诊断和处理意见"
            .Add IIf(IsNull(lobjMain("体检单号")), "", lobjMain("体检单号").Value), "体检单号"
        End With
        
        '二级子集合：一个体检人员的附加项目名称和附加项目值
        lstrSQL = "select * from 体检管理_体检附加信息表 where 系统编号 = '" & lcolItem("系统编号") & "'"
        Set lobjRec = dafuncGetData(lstrSQL)
        Set lcolItem1 = New Collection
        While Not lobjRec.EOF
            Set lcolItem11 = New Collection
            lcolItem11.Add lobjRec("附加项目").Value, "项目名"
            lcolItem11.Add IIf(IsNull(lobjRec("项目值")), "", lobjRec("项目值").Value), "项目值"
            
            lcolItem1.Add lcolItem11, lcolItem11("项目名")
            lobjRec.MoveNext
        Wend
        
        lcolItem.Add lcolItem1, "附加项目"
        
        '二级子集合. 包含一个体检人员的体检结果信息
        lstrSQL = "select * from 体检管理_体检结果数据库 where 系统编号 = '" & lcolItem("系统编号") & "'"
        Set lobjRec = dafuncGetData(lstrSQL)
        Set lcolItem1 = New Collection
        While Not lobjRec.EOF
            Set lcolItem11 = New Collection
            lcolItem11.Add lobjRec("体检项目").Value, "体检项目编码"
            lcolItem11.Add lobjRec("体检项目名称").Value, "体检项目名称"
            lcolItem11.Add IIf(IsNull(lobjRec("体检结果")), "", lobjRec("体检结果").Value), "体检结果"
            
            lcolItem1.Add lcolItem11, lcolItem11("体检项目编码")
            lobjRec.MoveNext
        Wend
        
        lcolItem.Add lcolItem1, "体检结果信息"
        
        '把一级子集合加入到要返回的集合中.
        lcolInfo.Add lcolItem, lcolItem("系统编号")
        lobjMain.MoveNext
    Wend
    
    lobjMain.Close
    If lobjRec Is Nothing Then
    Else
        lobjRec.Close
    End If
    Set Func获取已完毕的体检信息 = lcolInfo
    Exit Function
    
errHandler:
    sfsub错误处理 "体检对外接口", "clsManageTransmission", "Func获取已完毕的体检信息", Err.Number, Err.Description, True
    Set Func获取已完毕的体检信息 = New Collection
    Exit Function
    Resume
End Function
'功能：获取满足输入参数并使用指定体检表的体检人员的已体检完毕（无需复查）的并且满足指定访问标志值的体检信息。
'输入：查询的过滤条件.
'返回：collection[系统编号，姓名，性别，年龄，健康档案编号，单位申请编号，单位名称，附加项目[项目名，项目值]，体检日期，
'        体检结论id，体检结论名称，诊断和处理意见，体检结果信息[体检项目编码，体检项目名称，体检结果]]
'作者：刘浩。
Public Function Func获取已完毕的包括复查的体检信息(para体检表名称 As String, _
                                        para标志类型 As String, _
                                        para标志值 As String) As Collection
    
    Dim lobjMain As Object  '体检管理_体检访问数据库中的满足参数的体检记录。
    Dim lobjRec As Object
    Dim lcolItem As Collection
    Dim lcolItem1 As Collection
    Dim lcolItem11 As Collection
    Dim lcolInfo As Collection
    Dim lstrFilter As String   '根据参数拼成的sql语句的条件。
    Dim lstrSQL As String
    
    On Error GoTo errHandler
        
    '获取满足条件的体检记录。
    lstrSQL = "select a.* from 体检管理_体检访问数据库 where 体检状态=3 and isnull(复查体检表名,'')='' and isnull(" & para标志类型 & ",'')='" & para标志值 & "'"
    lstrSQL = lstrSQL & " and (体检表名称='" & para体检表名称 & "' or (系统编号 in (select 复查系统编号 from 体检管理_体检访问数据库 where 体检状态=3 and isnull(" & para标志类型 & ",'')='" & para标志值 & "' and 体检表名称='" & para体检表名称 & "')))"
    
    Set lobjMain = dafuncGetData(lstrSQL)
    Set lcolInfo = New Collection
    While Not lobjMain.EOF
        '返回值的一级子集合
        Set lcolItem = New Collection
        With lcolItem
            .Add lobjMain("系统编号").Value, "系统编号"
            .Add lobjMain("健康档案编号").Value, "健康档案编号"
            .Add lobjMain("姓名").Value, "姓名"
            .Add lobjMain("性别").Value, "性别"
            If IsDate(IIf(IsNull(lobjMain("出生日期").Value), "", lobjMain("出生日期").Value)) Then
                .Add DateDiff("yyyy", lobjMain("出生日期").Value, Date), "年龄"
            Else
                .Add 0, "年龄"
            End If
            .Add IIf(IsNull(lobjMain("单位申请编号")), "", lobjMain("单位申请编号").Value), "单位申请编号"
            .Add IIf(IsNull(lobjMain("单位名称")), "", lobjMain("单位名称").Value), "单位名称"
            .Add lobjMain("体检日期").Value, "体检日期"
            .Add IIf(IsNull(lobjMain("体检结论")), "", lobjMain("体检结论").Value), "体检结论"
            .Add IIf(IsNull(lobjMain("诊断和处理意见")), "", lobjMain("诊断和处理意见").Value), "诊断和处理意见"
            .Add IIf(IsNull(lobjMain("体检单号")), "", lobjMain("体检单号").Value), "体检单号"
        End With
        
        '二级子集合：一个体检人员的附加项目名称和附加项目值
        lstrSQL = "select * from 体检管理_体检附加信息表 where 系统编号 = '" & lcolItem("系统编号") & "'"
        Set lobjRec = dafuncGetData(lstrSQL)
        Set lcolItem1 = New Collection
        While Not lobjRec.EOF
            Set lcolItem11 = New Collection
            lcolItem11.Add lobjRec("附加项目").Value, "项目名"
            lcolItem11.Add IIf(IsNull(lobjRec("项目值")), "", lobjRec("项目值").Value), "项目值"
            
            lcolItem1.Add lcolItem11, lcolItem11("项目名")
            lobjRec.MoveNext
        Wend
        
        lcolItem.Add lcolItem1, "附加项目"
        
        '二级子集合. 包含一个体检人员的体检结果信息
        lstrSQL = "select * from 体检管理_体检结果数据库 where 系统编号 = '" & lcolItem("系统编号") & "'"
        Set lobjRec = dafuncGetData(lstrSQL)
        Set lcolItem1 = New Collection
        While Not lobjRec.EOF
            Set lcolItem11 = New Collection
            lcolItem11.Add lobjRec("体检项目").Value, "体检项目编码"
            lcolItem11.Add lobjRec("体检项目名称").Value, "体检项目名称"
            lcolItem11.Add IIf(IsNull(lobjRec("体检结果")), "", lobjRec("体检结果").Value), "体检结果"
            
            lcolItem1.Add lcolItem11, lcolItem11("体检项目编码")
            lobjRec.MoveNext
        Wend
        
        lcolItem.Add lcolItem1, "体检结果信息"
        
        '把一级子集合加入到要返回的集合中.
        lcolInfo.Add lcolItem, lcolItem("系统编号")
        lobjMain.MoveNext
    Wend
    
    lobjMain.Close
    If lobjRec Is Nothing Then
    Else
        lobjRec.Close
    End If
    Set Func获取已完毕的包括复查的体检信息 = lcolInfo
    Exit Function
    
errHandler:
    sfsub错误处理 "体检对外接口", "clsManageTransmission", "Func获取已完毕的包括复查的体检信息", Err.Number, Err.Description, True
    Set Func获取已完毕的包括复查的体检信息 = New Collection
    Exit Function
    Resume
End Function
'功能：设置指定系统编号的体检记录指定标志类型的标志。
'输入：paraCode  系统编号。
'      paraType  标志类型：健康证，劳动卫生，学校卫生，放射卫生。
'      paraValue 要设置的标志值。
'返回：无
'作者：刘浩。
Public Sub Sub设置体检访问标志(paraCode As String, paraType As String, paraValue As String)
Attribute Sub设置体检访问标志.VB_Description = "'功能:设置指定系统编号的体检记录指定标志类型的标志。\r\n'输入:paraCode  系统编号。\r\n'     paraType  标志类型：健康证，劳动卫生，学校卫生，放射卫生。\r\n'     paraValue 要设置的标志值。\r\n'返回:无"
    Dim lstrSQL As String
    
    On Error GoTo errHandler
    lstrSQL = "update 体检管理_体检访问标志表 set " & paraType & " = '" & paraValue & "' where 系统编号 = '" & paraCode & "'"
    dafuncGetData lstrSQL
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Sub设置体检访问标志", 6666, "标志类别“" & paraType & "”不能识别！", True
End Sub

'功能：重新从本地工作站配置文件中获取工作模式设置信息，并重新获取业务设置。
'输入：无
'作者：刘浩。
Public Sub Sub更新工作模式()
Attribute Sub更新工作模式.VB_Description = "'功能:  重新从本地工作站配置文件中获取工作模式设置信息，并重新获取业务设置。\r\n"
    On Error GoTo errHandler
    mobj工作站配置.subRefresh
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Sub设置体检访问标志", Err.Number, Err.Description, True
End Sub

'功能：从外单位报盘excel文件中导入单位的体检人员名单。文件中包含数据：单位名称，姓名，性别，年龄。
'      并且把文件名（带路径）记入本地配置文件中。
'输入：文件路径(带文件名)
'返回：Collection[单位名称，姓名，性别，年龄]key: 索引
'作者： 刘浩。
Public Function Func读报盘文件中体检人员信息(paraInputFile As String) As Collection
Attribute Func读报盘文件中体检人员信息.VB_Description = "'功能:  从外单位报盘excel文件中导入单位的体检人员名单。文件中包含数据：单位名称，姓名，性别，年龄。\r\n    '       并且把文件名（带路径）记入本地配置文件中。\r\n    '输入:  文件路径(带文件名)\r\n    '返回值: Collection[单位名称，姓名，性别，年龄]key: 索引\r\n"
    Dim lobjCn As ADODB.Connection
    Dim lobjcm As ADODB.Command
    Dim lobjRec As Object
    Dim lcolInfo As Collection
    Dim lcolItem As Collection
    Dim i As Long
    
    On Error GoTo errHandler
    
    '把文件名写入工作站配置文件。
    mobj工作站配置.Excel文件 = paraInputFile
    
    '建立指向Excel文件的odbc数据源。
    subRegisterDataSource paraInputFile
    
    '建立与excel文件的连接。
    Set lobjCn = New ADODB.Connection
    lobjCn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dyFyExcel;" & _
            "Initial Catalog=" & paraInputFile
    lobjCn.Open
        
    '获取Excel文件中的内容。
    Set lobjcm = New ADODB.Command
    lobjcm.ActiveConnection = lobjCn
    lobjcm.CommandText = "select * from [Sheet1$]"
    Set lobjRec = lobjcm.Execute
    
    '把获取的内容返回。
    Set lcolInfo = New Collection
    i = 1
    While Not lobjRec.EOF
        Set lcolItem = New Collection
        lcolItem.Add IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称").Value), "单位名称"
        lcolItem.Add IIf(IsNull(lobjRec("姓名")), "", lobjRec("姓名").Value), "姓名"
        lcolItem.Add IIf(IsNull(lobjRec("性别")), "", lobjRec("性别").Value), "性别"
        lcolItem.Add IIf(IsNull(lobjRec("年龄")), "", lobjRec("年龄").Value), "年龄"
        
        If lcolItem(1) = "" And lcolItem(2) = "" And lcolItem(3) = "" Or lcolItem(4) = "" Then
        Else
            lcolInfo.Add lcolItem, CStr(i)
        End If
        i = i + 1
        lobjRec.MoveNext
    Wend
    
    '关闭与Excel文件中的连接。
    lobjCn.Close
    Set lobjCn = Nothing
    
    Set Func读报盘文件中体检人员信息 = lcolInfo
    Exit Function
    
errHandler:
    Select Case Err.Number
        Case 76
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "Func读报盘文件中体检人员信息", 6666, "系统无法完成导入，因为：你输入的文件路径不存在。", True
        Case Else
            sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Func读报盘文件中体检人员信息", 6666, "报盘文件格式非法。请检查报盘文件是否是Excel文件，并且该工作簿中是否有Sheet1，Sheet1上是否有“姓名、性别、年龄、单位名称”项。", True
    End Select
End Function

'功能：建立从文件中导入的体检人员的健康档案，并进行体检登记处理。
'输入：paraTemplateName 体检表模板名称。
'      paraProgress     ProgressBar显示处理进度。
'      paraList as collection[体检表名称,显示导入进度控件名称,导入名单]。
'作者：刘浩。
Public Sub Sub导入体检人员登记(paraTemplateName As String, paraProgress As Object, paraList As Collection)
    Dim lobjTemplate As Object '体检表模板对象clsMedicalExamTemplate。
    Dim lobjMedical As Object  '体检对象clsMedicalExam。
    Dim lcolItem As Collection
    Dim lstrError As String
    
    On Error GoTo errHandle
    '创建体检表模板对象。
    Set lobjTemplate = CreateObject("体检对象部件.clsMedicalExamTemplate")
    lobjTemplate.体检表名 = paraTemplateName
    
    '创建体检对象。
    Set lobjMedical = CreateObject("体检对象部件.clsMedicalExam")
    
    '设置进度条的最大值。
    paraProgress.Max = paraList.Count
    paraProgress.Value = 0
    
    '开始事务。
    dasubBeginTran
    
    '修改：2001-11-5（根据单位名称，查找单位档案，获取卫生种类、行业类别）。
    Dim lstrTemp As String
    Dim lobjBase As Object
    If paraList.Count > 0 Then
        lstrTemp = paraList(1)("单位名称")
        Set lobjBase = dafuncGetData("select 卫生种类,行业类别,卫生种类BH,行业类别BH=b.编号 from dbo.单位档案_单位定位查询视图 a,dbo.系统管理_行业类别字典视图 b where a.单位名称='" & lstrTemp & "' and a.行业类别 = b.名称")
    End If
    
    For Each lcolItem In paraList
        With lobjMedical
            .系统编号 = .Func分配系统编号
            
            .体检类别 = 0 '初检登记。
            
            .体检表.体检表名 = paraTemplateName
            If Len(lobjMedical.体检表.试管编号字母) = 0 Then
                lobjMedical.体检表.试管编号字母 = Left(lobjTemplate.试管字母编号, 1)
            End If
            
            .体检人员.姓名 = lcolItem("姓名")
            If Not IsNumeric(lcolItem("年龄")) Then
                Err.Raise 6666, , "年龄非法！"
            ElseIf lcolItem("年龄") > 150 Then
                Err.Raise 6666, , "年龄超过150岁，这是系统不允许的！"
            End If
            .体检人员.出生日期 = CStr(DateAdd("yyyy", -CInt(lcolItem("年龄")), Date))
            If Trim(lcolItem("性别")) <> "男" And Trim(lcolItem("性别")) <> "女" Then
                Err.Raise 6666, , "导入数据中“性别”非法，请在网格中检查并更正所有性别数据（只能是“男”或“女”）！"
            End If
            .体检人员.性别 = lcolItem("性别")
            .体检人员.单位名称 = lcolItem("单位名称")
            .体检日期 = Format(Date, "yyyy-mm-dd")
            
            '修改：2001-11-5（根据单位名称，查找单位档案，获取卫生种类、行业类别）。
            On Error Resume Next
            If lobjBase.RecordCount > 0 Then
                .体检人员.卫生种类 = lobjBase("卫生种类BH").Value
                .体检人员.行业类别 = lobjBase("行业类别BH").Value
                
                .体检表.Sub填附加信息值 "卫生种类", lobjBase("卫生种类BH").Value & " " & lobjBase("卫生种类").Value
                .体检表.Sub填附加信息值 "行业类别", lobjBase("行业类别BH").Value & " " & lobjBase("行业类别").Value
            End If
            On Error GoTo errHandle
            
            .Sub保存体检登记信息
        End With
        
        '显示进度。
        paraProgress.Value = paraProgress.Value + 1
    Next
    
    '提交事务。
    dasubCommitTran
    
    '释放对象。
    Set lobjTemplate = Nothing
    Set lobjMedical = Nothing
    Exit Sub
    
errHandle:
    Select Case Err.Number
        Case 76
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "Sub导入体检人员登记", 6666, "系统无法完成导入，因为：你输入的文件路径不存在。", True
        Case Else
            sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Sub导入体检人员登记", 6666, Err.Description, True
    End Select
End Sub

'功能：更改已被导入过的文件的文件名,目地是防止同一文件被导入多次。
'输入：paraFileName   文件名(带路径)。
'作者：刘浩。
Public Sub sub更改文件名(paraFileName As String)
    Dim lobjFso As Object
    Dim lobjFile As Object
    
    On Error GoTo errHandler
    Set lobjFso = CreateObject("Scripting.FileSystemObject")
    Set lobjFile = lobjFso.GetFile(paraFileName)
    If Mid(lobjFile.Name, Len(lobjFile.Name) - 3, 1) = "." Then
        lobjFile.Name = Left(lobjFile.Name, Len(lobjFile.Name) - 3) & "rnd"
    Else
        lobjFile.Name = lobjFile.Name & "rnd"
    End If
    
    Set lobjFile = Nothing
    Set lobjFso = Nothing
    
    Exit Sub
errHandler:
    If Err.Number = 7 Then
        sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "sub更改文件名", 6666, "用户没有更改文件名的权限或此文件正在被使用！", True
    Else
        sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "sub更改文件名", Err.Number, Err.Description, True
    End If
    Exit Sub
    Resume
End Sub

'功能：导入数据前,先指定文件复制到当前程序执行的路径下, 并更名为"数据传输导入数据库.mdb"
'作者：刘浩。
Public Sub sub导入准备(ByVal paraFileName As String)
    Dim lobjFso As Object
    Dim lobjFile As Object
    Dim lstrDriveName As String
    
    On Error GoTo errHandler
    
    Set lobjFso = CreateObject("Scripting.FileSystemObject")
    Set lobjFile = lobjFso.GetFile(paraFileName)
    
    lstrDriveName = lobjFile.Drive
    If UCase(Right(lobjFile.Name, 3)) <> "MDB" Then
        Err.Raise 6666, , "指定文件不是一个MDB文件。"
    End If
    
    mstrInMDBPath = paraFileName
    
    '拷贝文件前，先断开连接。
    mobjInMdbFile.FileName = ""
    
    '重新建立连接。
    mobjInMdbFile.FileName = mstrInMDBPath
    
    '把文件名写入工作站配置。
    mobj工作站配置.内部导入文件 = paraFileName
        
    '释放对象。
    Set lobjFile = Nothing
    Set lobjFso = Nothing
    
    Exit Sub
errHandler:
    Select Case Err.Number
        Case 70
            '文件正在被使用。
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导入准备", 6666, "文件正在被其它人使用，暂时不能被导入。", True
        Case 76
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出准备", 6666, "系统无法完成导入，因为：你输入的文件路径不存在。", True
        Case -2147024784
            '磁盘空间不够。
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导入准备", 6666, "磁盘空间不够，请整理“" & lstrDriveName & "”上的磁盘空间！", True
        Case Else
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导入准备", Err.Number, Err.Description, True
    End Select
    Exit Sub
    Resume
End Sub

'功能：导出数据前,先指定文件复制到用户指定的路径下, 并更名为用户指定的文件名
'作者： 刘浩。
'修改：2001-4-24（杨春：避免导出mdb文件越来越大，把拷贝文件放在导出之前）
Public Sub sub导出准备(ByVal paraFileName As String)
    Dim lobjFso As Object
    Dim lobjFile As Object
    Dim lstrDriveName As String
    
    On Error GoTo errHandler
    
    '获取输入文件所在的驱动器号。
    lstrDriveName = Left(paraFileName, 2)
    
    Set lobjFso = CreateObject("Scripting.FileSystemObject")
    
    '检查系统默认导出文件是否存在。
    If Dir(mstrOutTmpMdbFile) = "" Then
        Err.Raise 6666, , "部件文件“体检对外接口部件.dll”所在路径下找不到系统文件“数据传输导出数据库.mdb”，请先找系统管理员拷贝一个同名文件到系统路径下。"
    End If
    
    If paraFileName = mstrOutTmpMdbFile Then
        Err.Raise 6666, , "你选择了系统使用的文件，请重新输入导出数据库所在的路径和文件名！"
    End If
    
    '设置目的文件的属性为可读可写。
    If Dir(paraFileName) <> "" Then
        Set lobjFile = lobjFso.GetFile(paraFileName)
        lobjFile.Attributes = 0
    End If
    
    '拷贝文件前，先断开连接。
    mobjOutMdbFile.FileName = ""
    
    '拷贝文件。
    lobjFso.copyfile mstrOutTmpMdbFile, paraFileName, True
        
    '设置导出文件名。
    mstrOutMDBPath = paraFileName
    
    '再建立连接。
    mobjOutMdbFile.FileName = paraFileName
    
    '把文件名写入工作站配置。
    If mobj工作站配置.内部导出文件 <> paraFileName Then
        mobj工作站配置.内部导出文件 = paraFileName
    End If
    
    '释放对象。
    Set lobjFso = Nothing
    Exit Sub
errHandler:
    Select Case Err.Number
        Case 70
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出准备", 6666, "文件正在被其它人使用，暂时不能被导出。", True
        Case 76
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出准备", 6666, "系统无法完成导出，因为：你输入的文件路径不存在。", True
        Case -2147024784
            '磁盘空间不够。
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出准备", 6666, "磁盘空间不够，请整理“" & lstrDriveName & "”上的磁盘空间！", True
        Case Else
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出准备", 6666, "系统无法完成导出，因为：" & Err.Description, True
    End Select
End Sub

'功能：导出数据完成后,先指定文件复制到用户指定的路径下, 并更名为用户指定的文件名
'说明：不再用该接口（避免导出mdb文件越来越大，把拷贝文件放在导出之前）。
'作者：刘浩。
Public Sub sub导出结束(ByVal paraFileName As String)
    Dim lobjFso As Object
    Dim lstrDriveName As String
    On Error GoTo errHandler
    
    Set lobjFso = CreateObject("Scripting.FileSystemObject")
    
    If paraFileName = IIf(Right(App.Path, 1) = "\", App.Path & "数据传输导出数据库.mdb", App.Path & "\数据传输导出数据库.mdb") Then
        Err.Raise 6666, , "导出失败，原因是你选择了系统使用的文件，请重新输入导出数据库！"
    End If
    
    '拷贝文件前，先断开连接。
    mobjOutMdbFile.FileName = ""
    
    lstrDriveName = Left(paraFileName, 2)
    '拷贝文件。
    lobjFso.copyfile IIf(Right(App.Path, 1) = "\", App.Path & "数据传输导出数据库.mdb", App.Path & "\数据传输导出数据库.mdb"), paraFileName, True
    
    
    '把文件名写入工作站配置。
    If mobj工作站配置.内部导出文件 <> paraFileName Then
        mobj工作站配置.内部导出文件 = paraFileName
    End If
    
    '释放对象。
    Set lobjFso = Nothing
    Exit Sub
errHandler:
    Select Case Err.Number
        Case 70
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出结束", 6666, "文件正在被其它人使用，暂时不能被导出。", True
        Case 76
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出结束", 6666, "系统无法完成导出，因为：你输入的文件路径不存在。", True
        Case -2147024784
            '磁盘空间不够。
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出结束", 6666, "磁盘空间不够，请整理“" & lstrDriveName & "”上的磁盘空间！", True
        Case Else
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出结束", Err.Number, Err.Description, True
    End Select
End Sub

'功能：获取本接口导出的mdb数据库文件中的"体检管理_数据传输类型表"中所记录的所有数据分类
'返回：Recordset[数据分类名]。
'作者：刘浩。
Public Property Get 所有数据分类清单() As Object
    On Error GoTo errHandler
    Set 所有数据分类清单 = dafuncGetData("select distinct 数据分类名 from 体检管理_数据分类信息表")

    Exit Property
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "Property Get 所有数据分类清单", Err.Number, Err.Description, True
End Property

'功能：获取本接口导出的mdb数据库文件中的"体检管理_数据传输类型表"中所记录的所有数据分类
'输入：数据服务器类型 as string
'返回：recordset[数据分类名]
'作者：刘浩。
Public Function Func获取mdb文件中的数据分类清单(ByVal paraServerType As String) As Object
    Dim lstrSQL As String
        
    On Error GoTo errHandler
    
    '建立与mdb数据库的连接。
    If mobjInMdbFile.FileName = "" Then
        mobjInMdbFile.FileName = mstrInMDBPath
    End If
    
    '获取mdb库中所有类型。
    Set Func获取mdb文件中的数据分类清单 = mobjInMdbFile.funcExecuteSql("select 数据分类名 from 体检管理_数据传输类型表 where 数据分类名 in (select 数据分类名 from 体检管理_数据分类信息表 where 服务器类型 = '" & paraServerType & "')")
    
    Exit Function
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Func获取mdb文件中的数据分类清单", Err.Number, Err.Description, True
    Exit Function
    Resume
End Function

'功能:  获取本接口导出的mdb数据库文件中的"体检管理_数据传输信息表"中所记录的数据范围值
'返回值:   recordset[数据范围名，数据范围值]
'作者： 刘浩。
Public Function Func获取mdb数据交换文件中的数据范围() As Object
    Dim lstrMDBPath As String
    
    On Error GoTo errHandler
    
    '建立与mdb数据库的连接。
    If mobjInMdbFile.FileName = "" Then
        mobjInMdbFile.FileName = mstrInMDBPath
    End If
    
    Set Func获取mdb数据交换文件中的数据范围 = mobjInMdbFile.funcExecuteSql("select * from 体检管理_数据传输范围表")
    
    
    Exit Function
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Func获取mdb数据交换文件中的数据范围", Err.Number, Err.Description, True
End Function

'功能：清空除"体检管理_数据分类信息表"外本业务所有的表内容。
'作者：刘浩。
Public Sub Sub清空mdb数据交换文件()
Attribute Sub清空mdb数据交换文件.VB_Description = "'功能:清空除""体检管理_数据分类信息表""外本业务所有的表内容。"
    On Error GoTo errHandler
    
    mobjOutMdbFile.Sub清空文件
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Sub清空mdb数据交换文件", Err.Number, Err.Description, True
End Sub

'功能：在数据导入或导出前，取出要导入或导出的数据（只能是体检人员基本信息）。
'输入：Para数据范围 as collection[数据范围名，数据范围值]，Para类型 as integer(0导入/1 导出)
'      (数据范围名包含以下内容: 开始日期，结束日期，单位编号集，系统编号范围)
'      单位编号集，系统编号范围 都是以逗号分隔的字符串
'      paraType    0导入/1 导出
'      paraInputFile 导入文件的路径
'返回：体检基本信息 as recordset[系统编号，姓名，性别，年龄，单位名称]
'作者：刘浩。
Public Function Func查看数据(ByVal paraRange As Collection, ByVal paraType As Integer) As Object
Attribute Func查看数据.VB_Description = "'功能: 在数据导入或导出前，取出要导入或导出的数据（只能是体检人员基本信息）。\r\n    '输入:Para数据范围 as collection[数据范围名，数据范围值]，Para类型 as integer(0导入/1 导出)\r\n    '      (数据范围名包含以下内容: 开始日期，结束日期，单位编号集，系统编号范"
    Dim lobjRec As Object
    Dim lstrSQL As String
    
    On Error GoTo errHandler
    
    '获取并拼上查询条件。
    lstrSQL = "select * from 体检管理_体检基本数据库 where 1 = 1 "
    lstrSQL = lstrSQL & funcFilter(paraRange, IIf(paraType = 0, True, False))
    
    If paraType = 0 Then
        '导入，连接导入mdb文件。
        If mobjInMdbFile.FileName = "" Then
            mobjInMdbFile.FileName = mstrInMDBPath
        End If
        '执行sql语句。
        Set lobjRec = mobjInMdbFile.funcExecuteSql(lstrSQL)
    Else
        '导出，从sql server库中获取数据。
        Set lobjRec = dafuncGetData(lstrSQL)
    End If
            
    Set Func查看数据 = lobjRec
    Exit Function
    
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "Func查看数据", Err.Number, Err.Description, True
    Exit Function
    Resume
End Function

'功能:把站外体检所需要的基础数据，或站外体检后的满足指定数据范围和数据分类的结果信息导出到mdb文件
'输入:Para数据范围 as collection[数据范围名，数据范围值],paraobj导出进度 as ProgressBar,para数据分类 as collection
'      (数据范围名包含以下内容: 开始日期，结束日期，单位编号集，系统编号范围)
'单位编号集，系统编号范围 都是以逗号分隔的字符串
'作者： 刘浩。
Public Sub Sub数据导出(ByVal para数据范围 As Collection, _
                       ByVal para数据分类 As Collection, _
                       ByVal para进度条 As Object)
    Dim lobjTransSystem As clsTransSystem
    Dim lstrSQL As String
    Dim lstrFilter As String
    Dim lstrDriveName As String  '数据传输导出数据库所在磁盘驱动器。
    Dim i As Integer
    
    On Error GoTo errHandler
    
    lstrDriveName = Left(App.Path, 2)
    If para数据分类.Count = 0 Then
        Err.Raise 6666, , "没有被选中的数据分类！请选择要导出的数据分类。"
    End If
    
    '判断化验单和体检人员登记信息是否一起被导出。
    If sffunc判断集合键值是否存在(para数据分类, "化验单信息") And Not sffunc判断集合键值是否存在(para数据分类, "体检人员登记信息") Then
        sffuncMsg "要导出化验单信息，必需导出体检人员登记信息。" & vbCrLf & "体检人员登记信息将被一起导出。"
        para数据分类.Add "体检人员登记信息", "体检人员登记信息"
    End If
    
    If sffunc判断集合键值是否存在(para数据分类, "站外体检基础数据") And Not sffunc判断集合键值是否存在(para数据分类, "系统基础数据") Then
        If Not sffuncMsg("你确认导出站外体检基础数据，而不导出系统基础数据吗？ 若站外服务器上没有系统基础数据，你可能复无法导入体检基础数据。", sf询问) Then
            Err.Raise 6666, , "你选择终止导出操作。"
        End If
    End If
    
    '连接导出mdb文件。
    If mobjOutMdbFile.FileName = "" Then
        mobjOutMdbFile.FileName = mstrOutMDBPath
    End If

    '先清空文件。
    mobjOutMdbFile.Sub清空文件
    
    '设置进度条最大值。
    para进度条.Max = 2
    If sffunc判断集合键值是否存在(para数据分类, "体检人员登记信息") Or sffunc判断集合键值是否存在(para数据分类, "体检登记及结论信息") Then
        para进度条.Max = para进度条.Max + 5
    End If
    If sffunc判断集合键值是否存在(para数据分类, "站外体检基础数据") Then
        para进度条.Max = para进度条.Max + 10
    End If
    
    If sffunc判断集合键值是否存在(para数据分类, "化验单信息") Then
        para进度条.Max = para进度条.Max + 4
    End If
    para进度条.Value = 0
    
    '导出分类名为"系统基础数据"的数据
    If sffunc判断集合键值是否存在(para数据分类, "系统基础数据") Then
        
        '调用“通用对象.cls用户管理”的方法导出系统管理的所有表。
        umsub数据导出 mstrOutMDBPath, para进度条
    End If
    
    '导出分类名为"站外体检基础数据"的数据
    If sffunc判断集合键值是否存在(para数据分类, "站外体检基础数据") Then
        
        mobjOutMdbFile.sub传输体检医师项目设置表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检结论判断条件表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检项目设置表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检人员附加项目设置表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输编号生成规则表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检表模板基本信息表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub体检表模板体检结论表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检表模板体检项目表
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检表模板附加项目信息表
        para进度条.Value = para进度条.Value + 1
        
    End If
    
    '必须导出数据分类信息表。
    mobjOutMdbFile.sub传输数据分类信息表
    para进度条.Value = para进度条.Value + 1
    
    '导出分类名为"体检人员登记信息"或"体检登记及结论信息"的数据。
    If sffunc判断集合键值是否存在(para数据分类, "体检人员登记信息") Or sffunc判断集合键值是否存在(para数据分类, "体检登记及结论信息") Then
        mobjOutMdbFile.sub传输体检人员基本信息表 para数据范围
        
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检基本信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检结果信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检附加信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输体检访问标志表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        '使用DAO导出图片表。
        lstrFilter = funcFilter(para数据范围, True)
        mobjOutMdbFile.sub导出sql表到mdb文件 "系统管理_系统图片管理表", "子系统名='体检管理' and 图片编号 in (select 健康档案编号 from 体检管理_体检基本数据库 where 1=1 " & lstrFilter & ")"
    End If
    
    
    '导出分类名为"化验单信息"的数据
    If sffunc判断集合键值是否存在(para数据分类, "化验单信息") Then
        mobjOutMdbFile.sub传输化验单基本信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输化验单不正常结果信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输化验单试管编号表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjOutMdbFile.sub传输化验单不正常结果试管编号表 para数据范围
        para进度条.Value = para进度条.Value + 1
    End If
        
    '把数据分类信息写入"体检管理_数据传输类型表"、数据范围信息写入"体检管理_数据传输范围表"。
    For i = 1 To para数据范围.Count
        lstrSQL = "insert into 体检管理_数据传输范围表(范围名,范围值) values('" & _
            para数据范围.Item(i)("数据范围名") & "','" & para数据范围.Item(i)("数据范围值") & "')"
        mobjOutMdbFile.subExecuteSql lstrSQL
    Next i
    
    '在mdb库中登记数据分类信息。
    For i = 1 To para数据分类.Count
        lstrSQL = "insert into 体检管理_数据传输类型表(数据分类名) values('" & para数据分类.Item(i) & "')"
        mobjOutMdbFile.subExecuteSql lstrSQL
    Next i
    
    '关闭连接。
    mobjOutMdbFile.FileName = ""
    
    para进度条.Value = para进度条.Max
    
    Exit Sub
errHandler:
    Select Case Err.Number
        Case -2147024784
            '磁盘空间不够。
            sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub数据导出", 6666, "磁盘空间不够，请整理“" & lstrDriveName & "”上的磁盘空间！", True
        Case 76
            sfsub错误处理 "体检对外接口部件", "ClsManageTransmission", "sub导出准备", 6666, "系统无法完成导出，因为：你输入的文件路径不存在。", True
        Case Else
            sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub数据导出", 6666, "系统无法完成导出，因为：" & Err.Description, True
    End Select
End Sub

'功能：把站外体检所需要的基础数据，或站外体检后的满足指定数据范围和数据分类的结果信息从mdb文件 导入到SQLserver
'输入：Para数据范围 as collection[数据范围名，数据范围值],para进度条 as ProgressBar,para数据分类 as collection
'      (数据范围名包含以下内容: 开始日期，结束日期，单位编号集，系统编号范围)
'       单位编号集，系统编号范围 都是以逗号分隔的字符串
'作者： 刘浩。
Public Sub Sub数据导入(ByVal para数据范围 As Collection, _
                        ByVal para数据分类 As Collection, _
                        ByVal para进度条 As Object)
Attribute Sub数据导入.VB_Description = "'功能:把站外体检所需要的基础数据，或站外体检后的满足指定数据范围和数据分类的结果信息从mdb文件 导入到SQLserver\r\n    '输入:Para数据范围 as collection[数据范围名，数据范围值],para进度条 as ProgressBar,para数据分类 as collection\r\n    '      (数据范围名包含以"
    
    Dim lobjTransSystem As clsTransSystem
    Dim lstrDriveName As String
    On Error GoTo errHandler
    
    lstrDriveName = Left(App.Path, 2)
    
    '连接导入mdb文件。
    If mobjInMdbFile.FileName = "" Then
        mobjInMdbFile.FileName = mstrInMDBPath
    End If
    
    If para数据分类 Is Nothing Or para数据分类.Count = 0 Then
        Err.Raise 6666, , "没有被选中的数据分类！"
    End If
    
    Dim lstrServer  As String
    Dim lstrData As String
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    
    '给用户提示。
    If Not sffuncMsg("你确认要导入数据到服务器“" & lstrServer & "”上的数据库“" & lstrData & "”中吗？你若选择是，则该服务器上库中相应的原有内容将被覆盖，而且不可恢复。" & Chr(13) & Chr(10) & "切记！切记！切记！", sf询问) Then
        Err.Raise 6666, , "你选择了取消操作。"
    End If
    
    If sffunc判断集合键值是否存在(para数据分类, "站外体检基础数据") And Not sffunc判断集合键值是否存在(para数据分类, "系统基础数据") Then
        If Not sffuncMsg("你确认导入站外体检基础数据，而不导入系统基础数据吗？ 若站外服务器上没有系统基础数据，你可能复无法正常导入体检基础数据。", sf询问) Then
            Err.Raise 6666, , "你选择了取消操作。"
        End If
    End If
    
    '设置返回数据导入的进度的最大值
    para进度条.Max = 1
    If sffunc判断集合键值是否存在(para数据分类, "体检人员登记信息") Or sffunc判断集合键值是否存在(para数据分类, "体检登记及结论信息") Then
        para进度条.Max = para进度条.Max + 5
    End If
    If sffunc判断集合键值是否存在(para数据分类, "站外体检基础数据") Then
        para进度条.Max = para进度条.Max + 11
    End If
    If sffunc判断集合键值是否存在(para数据分类, "化验单信息") Then
        para进度条.Max = para进度条.Max + 5
    End If
    para进度条.Value = 0
    
    '设置SQL查询超时时间。
    dasubSetQueryTimeout 600
    
    '若要导入"系统基础数据"，先清空整个库。
    If sffunc判断集合键值是否存在(para数据分类, "系统基础数据") Then
        para进度条.Max = para进度条.Max + 1
        
        '清所有表。
        dafuncGetData "exec 系统管理_清空所有数据"
        
    ElseIf sffunc判断集合键值是否存在(para数据分类, "站外体检基础数据") Then
        '若要导入"站外体检基础数据"，先清库。
        para进度条.Max = para进度条.Max + 1
        
        '先清空业务表。
        sub清空业务表
        
    End If
    
    dasubBeginTran
    
    '导入系统基础数据。
    If sffunc判断集合键值是否存在(para数据分类, "系统基础数据") Then

        '调用“通用对象.用户管理”的接口导入系统管理的表。
        umsub数据导入 mstrInMDBPath, False, para进度条
        
        para进度条.Value = para进度条.Value + 1
    End If
    
    '导入"站外体检基础数据"的数据。
    If sffunc判断集合键值是否存在(para数据分类, "站外体检基础数据") Then

        mobjInMdbFile.sub传输体检项目设置表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检人员附加项目设置表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检医师项目设置表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检结论判断条件表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输编号生成规则表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检表模板基本信息表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub体检表模板体检结论表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检表模板体检项目表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检表模板附加项目信息表
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输数据分类信息表
        para进度条.Value = para进度条.Value + 1
        
    End If
    
    '导入分类名为"体检人员登记信息"的数据。
    If sffunc判断集合键值是否存在(para数据分类, "体检人员登记信息") Or sffunc判断集合键值是否存在(para数据分类, "体检登记及结论信息") Then
        mobjInMdbFile.sub传输体检人员基本信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检基本信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检附加信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检结果信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输体检访问标志表 para数据范围
        para进度条.Value = para进度条.Value + 1
    
        '使用DAO导出图片表。
        mobjInMdbFile.sub从mdb导入表内容到sql "系统管理_系统图片管理表", "a.子系统名='体检管理'"
        '删除图片编号非法的体检人员的图片。
        dafuncGetData "delete 系统管理_系统图片管理表 where 子系统名='体检管理' and 图片编号 not in (select 健康档案编号 from 体检管理_体检人员基本信息表)"
    End If
    
    '导入分类名为"化验单信息"的数据。
    If sffunc判断集合键值是否存在(para数据分类, "化验单信息") Then
        mobjInMdbFile.sub传输化验单基本信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输化验单不正常结果信息表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输化验单试管编号表 para数据范围
        para进度条.Value = para进度条.Value + 1
        
        mobjInMdbFile.sub传输化验单不正常结果试管编号表 para数据范围
        para进度条.Value = para进度条.Value + 1
    End If
        
    para进度条.Value = para进度条.Max
    
    dasubCommitTran
    
    Exit Sub
errHandler:
    Select Case Err.Number
        Case -2147024784
            '磁盘空间不够。
            sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub数据导入", 6666, "磁盘空间不够，请整理“" & lstrDriveName & "”上的磁盘空间！", True
        Case Else
            sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "Sub数据导入", Err.Number, Err.Description, True
    End Select
End Sub

'功能： 建立一个新的DSN(用于读取excel文件)数据源,如果注册表中已有同名数据源,用新的数据源设置覆盖已有设置.
'输入： excel文件名(带路径)
'作者： 刘浩。
Private Sub subRegisterDataSource(ByVal paraExcelName As String)
    On Error GoTo errHandler
    Dim lstrAttribs As String
    Dim lstrPath As String
    Dim lobjFso As Object
    
    Set lobjFso = CreateObject("Scripting.FileSystemObject")
    lstrPath = Left(paraExcelName, Len(paraExcelName) - Len(lobjFso.GetFileName(paraExcelName)) - 1)
    '建造关键字字符串。
    lstrAttribs = "Description=" _
            & "registerDataSource for excel" _
            & Chr$(13) & "DBQ=" & paraExcelName _
            & Chr$(13) & "defaultdir=" & lstrPath _
            & Chr$(13) & "FIL=excel 8.0 "
  
  
    '建立新的注册的 DSN。
    rdoEngine.rdoRegisterDataSource "dyFyExcel", "Microsoft Excel Driver (*.xls)", True, lstrAttribs
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口", "clsManageTransmission", "subRegisterDataSource", Err.Number, Err.Description, True
End Sub

'功能：清空站外体检的业务表内容。
'作者：杨春。
Private Sub sub清空业务表()
    On Error GoTo errHandler
    
    dafuncGetData "delete 体检管理_化验单基本信息表"
    dafuncGetData "delete 体检管理_体检基本信息表"
    dafuncGetData "delete 体检管理_体检人员基本信息表"
    dafuncGetData "delete 体检管理_体检医师项目设置表"
    dafuncGetData "delete 体检管理_体检结论判断条件表"
    dafuncGetData "delete 体检管理_试管编号最大流水号表"

    dafuncGetData "delete 体检管理_试管编号字母表"
    dafuncGetData "delete 体检管理_编号最大流水号表"
    dafuncGetData "delete 系统管理_操作日志记录表"
    
    dafuncGetData "delete 体检管理_体检表模板基本信息表"
    dafuncGetData "delete 体检管理_体检表模板体检结论表"
    
    dafuncGetData "delete 体检管理_数据分类信息表"
    
    On Error Resume Next
    dafuncGetData "set IDENTITY_INSERT 体检管理_数据分类信息表 OFF"
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口", "clsManageTransmission", "sub清空业务表", Err.Number, Err.Description, True
End Sub

'功能：给健康证提供修改修改基本信息的接口
'输入： 系统编号,姓名，性别，年龄，单位名称，卫生种类，行业类别,身份证号，职业"。
'返回：boolean。
'作者： 周雪松。
Public Function func修改从业基本信息(paraSysNo As String, paraName As String, paraSex As String, _
                                    paraAge As String, paraAppNum As String, paraUnit As String, paraWs As String, _
                                    paraBusiness As String, paraid As String, paraJob As String) As Boolean
    On Error GoTo errHandler
    Dim lstrSQL As String
    Dim lobjTmp As Object
    Dim lobjRec As Object
    func修改从业基本信息 = False
    '判断申请编号为空
    If IsNull(paraAppNum) Then
        paraAppNum = ""
    End If
    '判断该体检编号是否在体检系统存在
    lstrSQL = "select 健康档案编号 from 体检管理_体检基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjTmp = dafuncGetData(lstrSQL)
    If lobjTmp.RecordCount > 0 Then
        lstrSQL = "exec 体检管理_保存人员修改信息  '" & paraSysNo & "','" & paraName & "', '" & paraSex & "'," _
        & "'" & paraAge & "','" & paraAppNum & "','" & paraUnit & "','" & paraWs & "','" & paraBusiness & "','" & paraid & "','" & paraJob & "' "
          dafuncGetData lstrSQL
        func修改从业基本信息 = True
    Else
        func修改从业基本信息 = False
    End If
    Exit Function
errHandler:
    sfsub错误处理 "体检对外接口部件", "clsManageTransmission", "func修改从业基本信息", Err.numer, Err.Description, True
End Function
