VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManageMedicalExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'属性变量。
Private mobj工作站配置 As ClsLocalConfigure     '工作站配置信息：工作模式。
Private mobj当日工作记忆簿  As clsLocalWorkBook '[单位名称]。

'私有变量。
Private mobj收费对外接口 As Object              '用来发送收费请求，获取收费项目。
Private mobj单位对外接口 As Object              '用来进行单位定位。

'功能：对象初始化（创建属性对象）。
'作者：杨春
Private Sub Class_Initialize()
    Dim lobjRec As Object
    Dim lcolItem As Collection
    
    On Error GoTo errHandler
    
    '创建工作站配置对象。
    Set mobj工作站配置 = New ClsLocalConfigure
    
    '创建“当日工作记忆簿”对象，中获取体检登记当天的已录入的单位名称。
    Set mobj当日工作记忆簿 = New clsLocalWorkBook
    
    On Error Resume Next
    
    '创建收费对外接口。
    Set mobj收费对外接口 = CreateObject("收费接口.cls对外接口")
    
    '创建单位对外接口。
    Set mobj单位对外接口 = CreateObject("单位档案业务对象.ClsUnitInterface")
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Class_Initialize", 6666, lstrError, False
End Sub

'功能：获取指定设置项目的内容。
'输入：para设置项目 as string   设置项目名称。
'返回：设置值 as string
'作者：杨春
Public Property Get 业务设置(ByVal para设置项目 As String) As String
Attribute 业务设置.VB_Description = "'功能：获取指定设置项目的内容。\r\n    '返回：设置值 as string"
    Dim lobjRec As Object
    On Error GoTo errHandler
    业务设置 = ""
    Set lobjRec = dafuncGetData("select * from 体检管理_业务设置信息表 where 设置项目='" & para设置项目 & "'")
    If lobjRec.recordcount > 0 Then
        业务设置 = IIf(IsNull(lobjRec("设置值")), "", lobjRec("设置值"))
    End If
    
    Exit Property
errHandler:
    
End Property

'功能：获取所有业务设置项目的信息。
'返回：[[设置项目，设置值，枚举来源，说明]key:设置项目]
'作者：杨春
Public Property Get 所有业务设置() As Collection
    Dim lobjRec  As Object
    Dim lcolInfo As Collection
    Dim lcolItem As Variant
    
    On Error GoTo errHandler
    '从表“体检管理_业务设置信息表”获取业务设置：设置项目，设置值，枚举来源，说明。
    Set lcolInfo = New Collection
    Set lobjRec = dafuncGetData("select * from 体检管理_业务设置信息表")
    Do While Not lobjRec.EOF
        Set lcolItem = New Collection
        lcolItem.Add lobjRec("设置项目").Value, "设置项目"
        lcolItem.Add IIf(IsNull(lobjRec("设置值")), "", lobjRec("设置值").Value), "设置值"
        lcolItem.Add IIf(IsNull(lobjRec("枚举来源")), "", lobjRec("枚举来源").Value), "枚举来源"
        lcolItem.Add IIf(IsNull(lobjRec("说明")), "", lobjRec("说明").Value), "说明"
        
        lcolInfo.Add lcolItem, lcolItem("设置项目")
        lobjRec.MoveNext
    Loop
    lobjRec.Close
    
    Set 所有业务设置 = lcolInfo
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Property Get 所有业务设置", 6666, lstrError, True
End Property

'功能：获取所有可用的体检附加项目。
'返回：recordset[附加项目，录入标题，数据类型，数据长度，枚举值]。
'作者：杨春
Public Property Get 所有体检附加项目() As Object
Attribute 所有体检附加项目.VB_Description = "'功能：获取所有可用的体检附加项目。\r\n    '返回：recordset[附加项目，录入标题，数据类型，数据长度，枚举值]。"
    On Error GoTo errHandler

    Set 所有体检附加项目 = dafuncGetData("select * from 体检管理_体检人员附加项目设置表 order by 附加项目")
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Property Get 所有体检附加项目", 6666, lstrError, True
End Property

'功能：获取所有可用的体检收费标准。
'返回：collection[名称，总额]。
'作者：杨春
Public Property Get 所有体检收费标准() As Object
Attribute 所有体检收费标准.VB_Description = "'功能：获取所有可用的体检收费项目。\r\n    '返回：recordset[编码，名称，单价]。"
    Dim lobjRec As Object
    Dim lcolItems As Collection '返回结果。
    Dim lcolItem As Collection  '返回结果中的一项[名称，总额]。
    Dim lstrName As String      '收费标准名称。
    Dim ldblTotal As Double     '收费总额。
    
    On Error GoTo errHandler
    Set lcolItems = New Collection
    If mobj收费对外接口 Is Nothing Then
        Err.Raise 6666, , "使用收费管理提供的对外接口失败。请退出系统，重新注册“收费对外接口.dll”。"
    Else
        '根据收费接口对象获取收费标准：收费标准名称，收费项目编号，单价，数量
        Set lobjRec = mobj收费对外接口.func查询收费标准("子系统名称='" & P_SUBSYSNAME & "',")
        If lobjRec Is Nothing Then
            Err.Raise 6666, , "获取收费标准失败。请找系统管理员配置收费系统的基础数据（与体检系统相关的）。"
        End If
        '求出各收费标准的收费总额。
        Do While Not lobjRec.EOF
            '开始新标准。
            lstrName = lobjRec!收费标准名称
            ldblTotal = 0
            Do While lstrName = lobjRec!收费标准名称
                '累计总额。
                ldblTotal = ldblTotal + IIf(IsNull(lobjRec!单价), 0, lobjRec!单价) * IIf(IsNull(lobjRec!数量), 1, lobjRec!数量)
                lobjRec.MoveNext
                If lobjRec.EOF Then Exit Do
            Loop
            Set lcolItem = New Collection
            lcolItem.Add lstrName, "名称"
            lcolItem.Add ldblTotal, "总额"
            
            lcolItems.Add lcolItem, lstrName
        Loop
        
        Set 所有体检收费标准 = lcolItems
    End If
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Property Get 所有体检收费标准", 6666, lstrError, True
    Set 所有体检收费标准 = New Collection
    Exit Property
    Resume
End Property

'功能：获取所有体检结论分组情况。
'返回：recordset[编号,体检结论ID,体检结论名称,描述]。
'作者：杨春
Public Property Get 所有体检结论条件() As Object
Attribute 所有体检结论条件.VB_Description = "'功能：获取所有体检结论分组情况。\r\n    '返回：recordset[编号,体检结论id,体检结论名称,描述]。"
    On Error GoTo errHandler
    Set 所有体检结论条件 = dafuncGetData("select distinct 编号,体检结论ID,体检结论名称,描述 from 体检管理_体检结论判断条件数据库 order by 体检结论ID,编号")

    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Property Get 所有体检结论条件", 6666, lstrError, True
    Exit Property
    Resume
End Property

'功能：获取工作站配置对象。
'返回：clsLocalConfigure。
'作者：杨春
Public Property Get 工作站配置() As Object 'clsLocalConfigure
Attribute 工作站配置.VB_Description = "'功能：获取工作站配置对象。\r\n    '返回：clsLocalConfigure。"
    On Error GoTo errHandler
    Set 工作站配置 = mobj工作站配置

    Exit Property
errHandler:
    'sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Property Get 工作站配置", Err.Number, Err.Description, True
End Property

'功能：获取工作记忆簿对象。
'返回：clsLocalWorkBook。
'作者：杨春
Public Property Get 当日工作记忆簿() As clsLocalWorkBook
Attribute 当日工作记忆簿.VB_Description = "'功能：获取工作记忆簿对象。\r\n    '返回：clsLocalWorkBook。\r\n"
    On Error GoTo errHandler
    Set 当日工作记忆簿 = mobj当日工作记忆簿

    Exit Property
errHandler:
    'sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Property Get 当日工作记忆簿", Err.Number, Err.Description, True
End Property

'功能：为复查登记界面获取当前需要复查的体检记录。
'返回：recordset[系统编号，姓名，性别，单位名称，体检表名，诊断和处理意见]。
'作者：杨春
Public Function Func获取需要复查的体检记录() As Object
Attribute Func获取需要复查的体检记录.VB_Description = "'功能：为复查登记界面获取当前需要复查的体检记录。\r\n    '返回：recordset[系统编号，姓名，性别，单位名称，体检表名，诊断和处理意见]。\r\n"
    Dim lobj体检集 As clsMedicalExamSet
    Dim llng复查最大期限 As Long
    
    On Error GoTo errHandler
    '获取业务设置的"复查最大期限"。
    llng复查最大期限 = Val(业务设置("复查周期"))
        
    '创建"体检集"对象。
    Set lobj体检集 = New clsMedicalExamSet
    
    '根据参数设置过滤条件（其中"体检状态='3"，"复查标志=1"，"复查系统编号=''"，"体检日期>=now-复查周期"），
    With lobj体检集
        .体检状态 = "3"
        .复查标志 = 1
        .复查系统编号 = ""
        .从体检日期 = Format(DateAdd("d", -llng复查最大期限, Now), "yyyy-mm-dd") & " 00:00:00"
        .到体检日期 = Format(Now, "yyyy-mm-dd") & " 23:59:59"
    End With
    
    '根据"体检集.元素集"返回满足过滤条件的体检记录
    Set Func获取需要复查的体检记录 = lobj体检集.元素集("系统编号,姓名,性别,单位名称,体检表名称,诊断和处理意见")
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Func获取需要复查的体检记录", 6666, lstrError, True
    Exit Function
    Resume
End Function

'功能：体检登记业务的保存处理。
'输入：para体检对象 as clsMedicalExam    要保存的体检对象，
'      para旧系统编号 as string          可选参数：当该参数不为空，表示复查体检登记时所对应的旧体检编号。
'作者：杨春
'修改：2001-12-29（杨春）增加参数“收费项目集”,只有复查体检登记才有用。
Public Sub Sub体检登记(para体检对象 As clsMedicalExam, _
                    Optional para旧系统编号 As String = "", _
                    Optional para是否打印体检单 As Boolean = True, _
                    Optional para收费项目集 As Variant)
    Dim lobjRec As Object
    Dim lobj体检表模板 As clsMedicalExamTemplate   '用来获取收费标准名称。
    
    Dim lcol收费项目集 As Collection
    
    Dim lstr系统编号 As String       '体检对象的“系统编号”属性。
    Dim lstr是否收费 As String       '业务设置。
    Dim lstr是否快速录入 As String   '业务设置。
    Dim lstr是否打印体检单 As String '业务设置。
    
    Dim lstr收费批号 As String       '收费接口返回的收费编号。
    Dim lstr收费标准 As String
    Dim lbln是否已存在 As Boolean
    
    
    On Error GoTo errHandler
    
    '获取业务设置：是否快速登记，是否收费。
    lstr是否收费 = 业务设置("是否收费")
    lstr是否快速录入 = 业务设置("是否快速登记")
    
    '若不是快速录入，并且不是复查，必须录入姓名、单位名称。
    '修改说明:由于实际业务需要,从业人员在登记体检信息时,单位信息不做为必录入项目
    '详细说明:在判断条件中将单位名称的条件去掉,并修改相应的提示信息
    '修改人:徐冀川    时间:2002/11/19
    
    If Not lstr是否快速录入 = "是" And para体检对象.体检类别 <> P_EXAM_AGAIN Then
        With para体检对象.体检人员
            'If Trim(.姓名) = "" Or Trim(.单位名称) = "" Or Trim(.性别) = "" Or .年龄 = 0 Then
            If Trim(.姓名) = "" Or Trim(.性别) = "" Or .年龄 = 0 Then
                Err.Raise 6666, , "不是快速录入，必须录入姓名、性别、年龄。"
            End If
        End With
    End If
    lbln是否已存在 = para体检对象.是否已存在
    
    '开始事务。
    dasubBeginTran
    
    lstr系统编号 = para体检对象.系统编号
    
    '保存体检登记信息。
    para体检对象.Sub保存体检登记信息
    
    '若是复查，在旧系统编号体检记录上登记复查系统编号。
    If para体检对象.体检类别 = P_EXAM_AGAIN And para旧系统编号 <> "" Then
        '判断是否已复查登记过。
        Set lobjRec = dafuncGetData("select * from 体检管理_体检基本信息表 where 系统编号='" & para旧系统编号 & "' and isnull(复查系统编号,'')<>''")
        If lobjRec.recordcount > 0 Then
            Err.Raise 6666, , "该复查人员已经过了复查登记，不允许重复登记！"
        End If
        dafuncGetData "update 体检管理_体检基本信息表 set 复查系统编号='" & lstr系统编号 & "' where 系统编号='" & para旧系统编号 & "'"
    End If
    
    If IIf(lstr是否收费 = "", "否", lstr是否收费) = "是" And IIf(lstr是否快速录入 = "", "否", lstr是否快速录入) = "否" Then
        '不是快速登记，要收费。
        '判断收费接口是否可用。
        If mobj收费对外接口 Is Nothing Then
            Err.Raise 6666, , "收费对外接口无法使用。请退出系统并重新注册“收费对外接口部件.dll”，或在业务设置处配置“是否收费”=“否”。"
        End If
        
        '创建体检表模板对象。
        Set lobj体检表模板 = CreateObject("体检对象部件.clsMedicalExamTemplate")
        lobj体检表模板.体检表名 = para体检对象.体检表.体检表名
        
        '修改：2001-12-29（增加了可选参数）。
        If IsMissing(para收费项目集) Then
            '获取收费标准名称。
            lstr收费标准 = lobj体检表模板.收费标准
            
            '判断收费标准是否存在。
            If lstr收费标准 = "" Then
                Err.Raise 6666, , "没有设置体检表“" & para体检对象.体检表.体检表名 & "”的收费标准，不能发送收费信息。" & Chr(13) & Chr(10) & "请重新进入“体检表设置”操作界面设置该体检表的收费标准，或在业务设置处配置“是否收费”=“否”。"
            End If
            
            '按标准收费。
            lstr收费批号 = func按收费标准收费(para体检对象, lstr收费标准)
            
        Else
            Set lcol收费项目集 = para收费项目集
            If lcol收费项目集 Is Nothing Then
                Set lcol收费项目集 = New Collection
            End If
            If lcol收费项目集.Count > 0 Then
                '按具体的收费项目收费。
                lstr收费批号 = func按收费项目收费(para体检对象, lcol收费项目集)
            End If
        End If
        
        '修改收费记录的收费编号。
        If lstr收费批号 <> "" Then
            dafuncGetData "update 体检管理_体检基本信息表 set 收费批号='" & lstr收费批号 & "' where 系统编号='" & lstr系统编号 & "'"
            para体检对象.收费批号 = lstr收费批号
        End If
        
        '释放对象。
        Set lobj体检表模板 = Nothing
    End If
    
    '结束事务。
    dasubCommitTran
    
    '若是新增，并且业务设置("是否打印体检单")="是"，并且参数“para是否打印体检单=true”，
    '则打印体检单。
    '修改：2003-4-16（杨春）增加业务设置“是否打印体检表”。
    Dim lcolNo As Collection '系统编号集合。
    Dim lstrError As String
    Dim lstr是否打印体检表 As String
    Dim lstr文书名称 As String
    
    lstr是否打印体检单 = 业务设置("是否打印体检单")
    lstr是否打印体检表 = 业务设置("是否打印体检表")
    If IIf(lstr是否打印体检单 = "", "否", lstr是否打印体检单) = "是" Then
        lstr文书名称 = "体检单"
    ElseIf IIf(lstr是否打印体检表 = "", "否", lstr是否打印体检表) = "是" Then
        lstr文书名称 = "体检表"
    End If
    If Not lbln是否已存在 And para是否打印体检单 And lstr文书名称 <> "" Then
        Set lcolNo = New Collection
        lcolNo.Add lstr系统编号
        
        On Error Resume Next
        Sub打印文书 lstr文书名称, lcolNo, False
        
        If Err <> 0 Then
            '记录错误信息。
            lstrError = Err.Description
            sffuncMsg lstrError & "但你的保存是成功了的。你可以到“文书打印”操作界面上去补打" & lstr文书名称 & "。若你不想在体检登记时打印" & lstr文书名称 & "，请进入“业务设置”设置。", sf警告
        End If
    End If

    Exit Sub
errHandler:
    lstrError = func错误处理(Err.Number, Err.Description)
'    para体检对象.是否已存在 = False
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub体检登记", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：按收费标准进行体检收费。
'返回：收费批号。
'创建：2001-12-29（从sub体检登记中分离出来）。
'修改：2002-6-23（调用收费接口时，传递参数“业务份类”）。
Private Function func按收费标准收费(ByVal para体检对象 As clsMedicalExam, _
                                    ByVal para收费标准 As String) As String
    
    Dim lobjRec As Object            'Recordset。
    Dim lcol附加信息 As Collection   '需要传递给收费接口的附加信息：交费人、交费单位编号、经手人、收费人、主管科室。
    Dim lstr收费批号 As String       '生成的收费记录的收费批号。
    
    On Error GoTo errHandler
    
    '形成收费接口需要的附加信息集合。
    Set lcol附加信息 = New Collection
    With lcol附加信息
        .Add para体检对象.体检人员.姓名, "交费人"
        .Add para体检对象.体检人员.单位申请编号, "交费单位编号"
        .Add para体检对象.体检人员.单位名称, "交费单位名称"
        .Add um用户编号, "主管科室经手人"
        .Add um用户编号, "收费人"
        .Add um用户所属科室编号, "主管科室编号"
    End With
    
    lstr收费批号 = ""
    
    '判断是否是修改收费。
    If para体检对象.收费批号 = "" Then
        '添加收费记录。
        '输入：收费标准，附加信息[交费人、交费单位编号、交费单位名称、经手人、收费人、主管科室]
        '返回：总额；收费批号。
        '修改：2002-6-23（调用收费接口时，传递参数“业务份类”）。
        lstr收费批号 = mobj收费对外接口.func划价_收费标准(para收费标准, lcol附加信息, "", "体检收费")
        lstr收费批号 = Right(lstr收费批号, Len(lstr收费批号) - InStr(lstr收费批号, ";"))
    Else
        '修改收费前，判断收费状态。
        Set lobjRec = mobj收费对外接口.func查询费用信息("收费编号='" & para体检对象.收费批号 & "'")
        If lobjRec.recordcount > 0 Then
            If lobjRec("收费状态") = 0 Then
                '需要修改收费信息。
                '修改：2002-6-23（调用收费接口时，传递参数“业务份类”）。
                lstr收费批号 = mobj收费对外接口.func划价_收费标准(para收费标准, lcol附加信息, para体检对象.收费批号, "体检收费")
                lstr收费批号 = Right(lstr收费批号, Len(lstr收费批号) - InStr(lstr收费批号, ";"))
            ElseIf lobjRec("收费状态") = 2 Then
                '原有收费记录已退费，重新发送收费信息。
                '修改：2002-6-23（调用收费接口时，传递参数“业务份类”）。
                lstr收费批号 = mobj收费对外接口.func划价_收费标准(para收费标准, lcol附加信息, , "体检收费")
                lstr收费批号 = Right(lstr收费批号, Len(lstr收费批号) - InStr(lstr收费批号, ";"))
            Else
                lstr收费批号 = para体检对象.收费批号
                '修改：2002-10-9（杨春）若没有修改姓名和单位名称，允许修改体检登记信息。
                If para体检对象.体检人员.姓名 = IIf(IsNull(lobjRec("交费人")), "", lobjRec("交费人")) _
                    And para体检对象.体检人员.单位名称 = IIf(IsNull(lobjRec("交费单位名称")), "", lobjRec("交费单位名称")) Then
                    '没有修改姓名和单位名称，允许修改体检登记信息。直接返回原来的收费编号。
                    
                    '修改费用记录的交费单位编号。（允许补单位定位）
'                    If para体检对象.体检人员.单位申请编号 <> IIf(IsNull(lobjRec("交费单位编号")), "", lobjRec("交费单位编号")) Then
'                        dafuncGetData "update 收费管理_费用信息表 set 交费单位编号='" & para体检对象.体检人员.单位申请编号 & "' where 收费编号='" & lstr收费批号 & "'"
'                    End If
                Else
                    '原有收费记录已收费，不允许修改体检登记信息中的姓名和单位名称。
                    Err.Raise 6666, , "该体检人员已缴费，不能修改其姓名和单位名称。若确实需要修改，请叫收费科室先对编号为“" & lstr收费批号 & "”的记录作退费处理。"
                End If
            End If
        Else
            '查询收费状态失败。
            '修改：2002-9-17（杨春）已增加了报废功能，所以允许该情况出现。重新发送收费信息。
            lstr收费批号 = mobj收费对外接口.func划价_收费标准(para收费标准, lcol附加信息, , "体检收费")
            lstr收费批号 = Right(lstr收费批号, Len(lstr收费批号) - InStr(lstr收费批号, ";"))
            'Err.Raise 6666, , "查询费用信息失败！无法继续保存。请到收费科室检查收费编号为“" & para体检对象.收费批号 & "”的收费记录是否正常。"
        End If
    End If
    
    func按收费标准收费 = lstr收费批号
    
    Exit Function
errHandler:
    Err.Raise Err, , Error
    Exit Function
    Resume
    
End Function

'功能：按具体的收费项目进行体检收费。
'输入：para体检对象 as clsMedicalExam，
'      para收费项目集 as collection 'Item:[收费项目编号，单价],key:收费项目编号。
Private Function func按收费项目收费(ByVal para体检对象 As clsMedicalExam, _
                                    ByVal para收费项目集 As Collection) As String
    Dim lobjRec As Object            'Recordset。
    Dim lcol费用信息 As New Collection
    Dim lcolItem As Collection
    Dim lstr收费批号 As String       '生成的收费记录的收费批号。
    Dim i As Long
    
    On Error GoTo errHandler
    
    lstr收费批号 = ""
    
    '形成收费接口需要的费用信息集合：
    '交费人、交费单位编号、交费单位名称、主管科室编号、主管科室经手人、收费项目编号、单价、数量。
    Set lcol费用信息 = New Collection
    For i = 1 To para收费项目集.Count
        Set lcolItem = New Collection
        With lcolItem
            .Add para体检对象.体检人员.姓名, "交费人"
            .Add para体检对象.体检人员.单位申请编号, "交费单位编号"
            .Add para体检对象.体检人员.单位名称, "交费单位名称"
            .Add um用户编号, "主管科室经手人"
            .Add um用户所属科室编号, "主管科室编号"
            .Add para收费项目集(i)("收费项目编号"), "收费项目编号"
            .Add para收费项目集(i)("单价"), "单价"
            .Add 1, "数量"
        End With
        
        lcol费用信息.Add lcolItem, lcolItem("收费项目编号")
    Next
    
    '判断是否是修改收费。
    If para体检对象.收费批号 = "" Then
        '添加收费记录。
        '返回：总额；收费批号。
        lstr收费批号 = mobj收费对外接口.func划价_数据集合(lcol费用信息, , False)
    Else
        '修改收费前，判断收费状态。
        Set lobjRec = mobj收费对外接口.func查询费用信息("收费编号='" & para体检对象.收费批号 & "'")
        If lobjRec.recordcount > 0 Then
            If lobjRec("收费状态") = 0 Then
                '需要修改收费信息。
                lstr收费批号 = mobj收费对外接口.func划价_数据集合(lcol费用信息, para体检对象.收费批号, False)
            ElseIf lobjRec("收费状态") = 2 Then
                '原有收费记录已退费，重新发送收费信息。
                lstr收费批号 = mobj收费对外接口.func划价_数据集合(lcol费用信息, , False)
            Else
                '原有收费记录已收费，不允许修改体检登记信息。
                Err.Raise 6666, , "该体检人员已缴费，不能修改其体检登记信息。若确实需要修改，请叫体检人员先到收费科室退费。"
            End If
        Else
            '查询收费状态失败。
            '修改：2002-9-17（杨春）已增加了报废功能，所以允许该情况出现。
            'Err.Raise 6666, , "查询费用信息失败！无法继续保存。请到收费科室检查收费编号为“" & para体检对象.收费批号 & "”的收费记录是否正常。"
        End If
    End If
    
    func按收费项目收费 = Right(lstr收费批号, Len(lstr收费批号) - InStr(lstr收费批号, ";"))
    Exit Function
errHandler:
    Err.Raise Err, , Error
    Exit Function
    Resume
End Function

'功能：为体检结果录入界面获取可以登记体检结果的体检记录。
'输入：para从系统编号,para到系统编号     系统编号起止范围；
'      para体检日期                      要获取的体检记录所在的日期；
'      para编号类型                      0系统编号/1 试管编号。
'      para体检表名称
'返回：recordset [系统编号，姓名，性别，单位名称，体检表名，体检日期]。
'作者：杨春
'修改：2001-11-2（杨春）增加参数"para体检表名称"。
Public Function Func获取可修改的体检记录(ByVal para从系统编号 As String, _
                                    ByVal para到系统编号 As String, _
                                    ByVal para体检日期 As String, _
                                    Optional para编号类型 As Integer = 0, _
                                    Optional para体检表名称 As String = "") As Object
    
    Dim lobj体检集 As clsMedicalExamSet
 
    On Error GoTo errHandler
    ' 创建"体检集"对象。
    Set lobj体检集 = New clsMedicalExamSet
    
    ' 设置过滤条件：体检集.体检状态='0,1，2'，体检集.从系统编号=para从系统编号，体检集.到系统编号=para到系统编号，体检集.从体检日期=para体检日期，体检集.到体检日期=para体检日期。
    With lobj体检集
        .体检状态 = "0,1,2"
        If para编号类型 = 0 Then
            .从系统编号 = para从系统编号
            .到系统编号 = para到系统编号
        Else
            .从试管编号 = para从系统编号
            .到试管编号 = para到系统编号
        End If
        .从体检日期 = para体检日期
        If para体检日期 <> "" Then
            .到体检日期 = para体检日期 & " 23:59:59"
        Else
            .到体检日期 = ""
        End If
        
        '修改：2001-11-2（杨春）增加参数"para体检表名称"。
        .体检表名称 = para体检表名称
    End With
    
    ' 返回"体检集.元素集"。
    '修改：2002-10-14（若编号类型为试管编号，则多返回试管编号）。
    'Set Func获取可修改的体检记录 = lobj体检集.元素集("系统编号,姓名,性别,出生日期,单位名称,体检表名称,体检日期")
    Set Func获取可修改的体检记录 = lobj体检集.元素集("系统编号,试管编号,姓名,性别,出生日期,单位名称,体检表名称,体检日期")
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Func获取可修改的体检记录", 6666, lstrError, True
End Function

'功能：为"体检常规结果录入"和"检验项目结果录入"提供保存处理，即保存一个或一批人员的指定项目体检结果。
'输入：para系统编号范围 as collection[系统编号]；
'      para体检结果 as collection[体检项目，体检结果]；
'      para体检医师 as string       体检医师编号；
'      para填写日期 as string       体检结果填写日期。
'返回：无。
'作者：杨春
Public Sub Sub填写体检结果(ByVal para系统编号范围 As Collection, _
                        ByVal para体检结果 As Collection, _
                        ByVal para体检医师 As String, _
                        ByVal para填写日期 As String)
    Dim lobjProgress As clsProgress '进度显示。
    Dim lobj体检 As clsMedicalExam
    Dim lcolItem As Collection
    Dim i As Long
    
    On Error GoTo errHandler
    '创建进度显示对象。
    If para系统编号范围.Count > 1 Then
        Set lobjProgress = New clsProgress
        With lobjProgress
            .MaxValue = para系统编号范围.Count + 1
            .Message = "开始保存体检结果..."
            .Value = 1
            .subShow
        End With
    End If
    
    '创建"体检"对象。
    Set lobj体检 = New clsMedicalExam
    
    
    '开始事务。
    dasubBeginTran
    
    '依次保存该批系统编号所对应的体检记录的体检结果。
    For i = 1 To para系统编号范围.Count
        If Not lobjProgress Is Nothing Then
            lobjProgress.Message = "正在保存第" & i & "个(共" & para系统编号范围.Count & "个)体检人员的体检结果..."
        End If
        
        '体检对象实例化。
        lobj体检.系统编号 = para系统编号范围(i)
        
        '根据体检状态判断是否已下结论，若已下结论，不允许修改体检结果。
        If lobj体检.体检状态 = P_ENDED_STATUS Then
            Err.Raise 6666, , "系统编号为“" & lobj体检.系统编号 & "”的体检记录已下过结论，不允许修改体检结果。"
        End If
        
        '设置体检表的各项目体检结果。
        For Each lcolItem In para体检结果
            lobj体检.体检表.Sub填体检结果 lcolItem("体检项目"), lcolItem("体检结果"), para体检医师, para填写日期
        Next
        
        '保存体检结果和体检结论到数据库中。
        lobj体检.Sub保存体检结果
        
        '显示进度。
        If Not lobjProgress Is Nothing Then
            lobjProgress.Value = i
        End If
    Next
    
    '结束事务。
    dasubCommitTran
    
    If Not lobjProgress Is Nothing Then
        Set lobjProgress = Nothing
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub填写体检结果", 6666, lstrError, True
    
    Exit Sub
    Resume
End Sub
'功能：为"体检常规结果录入"和"检验项目结果录入"提供保存处理，即保存一个或一批人员的指定项目体检结果。
'输入：para系统编号范围 as collection[系统编号]；
'      para体检结果 as collection[体检项目，体检结果]；
'      para体检医师 as string       体检医师编号；
'      para填写日期 as string       体检结果填写日期。
'返回：无。
'其它：优化的地方：保存体检结果，直接写sql语句，而不再通过对象；下体检结论也用存储过程实现。
'作者：杨春
'修改：2001-4-20。
Public Sub Sub优化的填写体检结果(ByVal para系统编号范围 As Collection, _
                        ByVal para体检结果 As Collection, _
                        ByVal para体检医师 As String, _
                        ByVal para填写日期 As String)
                        
    Dim lobjProgress As clsProgress '进度显示。
    Dim lcolItem As Collection
    Dim lobjRec As Object           'Recordset。
    Dim lstrListNo As String        '系统编号串。
    Dim i As Long
    
    On Error GoTo errHandler
    
    '创建进度显示对象。
    If para系统编号范围.Count > 1 Then
        Set lobjProgress = New clsProgress
        With lobjProgress
            .MaxValue = para系统编号范围.Count + 1
            .Message = "开始保存体检结果..."
            .Value = 1
            .subShow
        End With
    End If
    
    '开始事务。
    dasubBeginTran
    
    '批量保存体检表的各项目体检结果。
    For i = 1 To para系统编号范围.Count
    
        '显示进度条提示信息。
        If Not lobjProgress Is Nothing Then
            lobjProgress.Message = "正在保存第" & i & "个(共" & para系统编号范围.Count & "个)体检人员的体检结果..."
        End If
        
        '依次保存该批系统编号所对应的体检记录的体检结果。
        For Each lcolItem In para体检结果
            dafuncGetData "update 体检管理_体检结果信息表 set 体检结果='" & lcolItem("体检结果") & "'" _
                        & ",体检医师=" & IIf(para体检医师 = "", "null", "'" & para体检医师 & "'") & ",填写日期='" & para填写日期 & "'" _
                        & " where 系统编号='" & para系统编号范围(i) & "' and 体检项目='" & lcolItem("体检项目") & "'"
        Next
        
        '下体检结论。
        On Error Resume Next
        dafuncGetData "exec 体检管理_下体检结论 '" & para系统编号范围(i) & "'"
        If Err <> 0 Then
            On Error GoTo errHandler
            Err.Raise 6666, , "系统保存体检结果失败，具体原因可能为：" & Chr(13) & Chr(10) & "(1) 该人员的体检记录已下过结论，此时系统不允许修改体检结果。" & Chr(13) & Chr(10) & "(2) 你输入的体检结果(包括常规和化验结果)的数据类型与体检结论判断条件中要求的数据类型不一致，比如：判断条件要求数值型，而你输入的体检结果是字符型。导致系统无法自动下体检结论。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 先进入“体检结论录入”取消该体检人员的体检结论。" & Chr(13) & Chr(10) & "(2) 核对该体检人员的所有常规和化验项目的结果的数据类型，以及体检结论判断条件设置中的类型，检查两者是否一致。并修改常规或化验项目结果，或修改体检结论判断条件设置。"
        Else
            On Error GoTo errHandler
        End If
        
        '显示进度。
        If Not lobjProgress Is Nothing Then
            lobjProgress.Value = i
        End If
    Next
    
    '结束事务。
    dasubCommitTran
    
    If Not lobjProgress Is Nothing Then
        Set lobjProgress = Nothing
    End If
        
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub优化的填写体检结果", 6666, lstrError, True
End Sub

'功能：为体检结论录入界面获取已下体检结论但还未确定的体检记录。
'输入：查询可选条件。只要任意一个参数为空，表示忽略该条件。
'返回：Recordset [系统编号，姓名，性别，单位名称，体检表名，体检日期，体检结论]。
'作者：杨春
'修改：2003-4-16（杨春）返回结果增加“卫生种类”，方便下结论时判断是否已录入卫生种类，否则信息无法流入健康证系统。
Public Function Func获取已下结论但未确定的体检记录(ByVal para体检日期 As String, _
                                            ByVal para单位名称 As String, _
                                            ByVal para从系统编号 As String, _
                                            ByVal para到系统编号 As String) As Object
    Dim lobj体检集 As clsMedicalExamSet
    
    On Error GoTo errHandler
    

    '创建"体检集"对象。
    Set lobj体检集 = New clsMedicalExamSet
    
    '除了参数中规定的条件外，另外设置：体检集.体检状态=2，体检集.下结论医师=当前用户。
    With lobj体检集
        .从体检日期 = para体检日期
        If para体检日期 <> "" Then
            .到体检日期 = para体检日期 & " 23:59:59"
        Else
            .到体检日期 = ""
        End If
        .单位名称 = para单位名称
        .从系统编号 = para从系统编号
        .到系统编号 = para到系统编号
        
        .体检状态 = "2"
    End With
    
    '返回"体检集.元素集"。
    Set Func获取已下结论但未确定的体检记录 = lobj体检集.元素集("系统编号,姓名,性别,年龄=datediff(year,出生日期,getdate()),单位名称,卫生种类,体检结论,诊断和处理意见,体检表名称,体检日期,复查体检表名")
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Func获取已下结论但未确定的体检记录", 6666, lstrError, True
    Exit Function
    Resume
End Function

'功能：为"体检结论录入"业务操作提供保存操作（包括登记操作日志）。
'输入： para系统编号 as string
'       para体检结论,para下结论日期,para诊断和处理意见,para复查体检表名    下结论医师所辖的体检结论；
'       para文书名称,para是否预览文书 。                     是否打印体检结果单（或体检通知单），打印时是否预览。
'作者：杨春
Public Sub Sub确定体检结论(ByVal para系统编号 As String, _
                            ByVal para体检结论 As String, _
                            ByVal para下结论日期 As String, _
                            ByVal para诊断和处理意见 As String, _
                            ByVal para复查体检表名 As String, _
                            ByVal para文书名称 As String, _
                            ByVal para是否预览文书 As Boolean)
    
    Dim lobj体检 As clsMedicalExam
    Dim lstrLog As String           '操作日志描述信息。
    Dim lstrError As String
    
    On Error GoTo errHandler
    
    '判断参数的合法性。
    If Len(Trim(para体检结论)) > 250 Then
        lstrError = lstrError & Chr(13) & Chr(10) & "请限制体检结论的长度在250个字符内。"
    End If
    If Len(Trim(para体检结论)) = 0 Then
        lstrError = lstrError & Chr(13) & Chr(10) & "必须录入体检结论。"
    End If
    If Len(para诊断和处理意见) > 250 Then
        lstrError = lstrError & Chr(13) & Chr(10) & "请限制诊断和处理意见的长度在250个字符内。"
    End If
    If lstrError <> "" Then
        Err.Raise 6666, , "录入数据不合法，无法保存：" & lstrError
    End If
    
    '创建"体检"对象。
    Set lobj体检 = New clsMedicalExam
    
    '体检对象实例化。
    lobj体检.系统编号 = para系统编号
    
    '检查是否已确定结论，若是，不允许再下结论。
    If lobj体检.体检状态 = P_ENDED_STATUS Then
        Err.Raise 6666, , "该体检记录已下过结论，不允许再下结论。若确实需要重新下结论，请先取消下结论。"
    End If
    
    '开始事务。
    dasubBeginTran
    
    '保存当前体检记录的体检结论。
    lobj体检.Sub保存体检结论 para体检结论, para诊断和处理意见, um用户编号, para复查体检表名, para下结论日期
    
    '修改：2001-9-21（恢复体检访问标志）。
    dafuncGetData "update 体检管理_体检访问标志表 set 健康证='' where 系统编号='" & para系统编号 & "'"
    
    '结束事务。
    dasubCommitTran
    
    '保存操作日志。
    On Error Resume Next
    lstrLog = "系统编号：" & para系统编号 & ",姓名：" & lobj体检.体检人员.姓名 & ",体检结论：" & para体检结论 & ",诊断：" & para诊断和处理意见 & IIf(para复查体检表名 = "", "", ",复查：" & para复查体检表名) & "。"
    oesubSave lstrLog, "体检结论录入"
    
    '打印文书。
    Dim lcolInfo As Collection
    If para文书名称 <> "" Then
        Set lcolInfo = New Collection
        lcolInfo.Add para系统编号
        Sub打印文书 para文书名称, lcolInfo, para是否预览文书
    End If
    
    Exit Sub
errHandler:
    
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub确定体检结论", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：取消体检结论，登记操作日志。
'输入：Para系统编号 as string  '体检的系统编号。
'作者：杨春
Public Sub Sub取消体检结论(ByVal para系统编号 As String)
Attribute Sub取消体检结论.VB_Description = "'功能：取消体检结论，登记操作日志。\r\n    '输入：Para体检对象 as clsMedicalExam。\r\n    '输出：Para体检对象 as clsMedicalExam。\r\n"
    Dim lobj体检 As clsMedicalExam '体检对象。
    Dim lobj复查体检 As clsMedicalExam
    Dim lobjRec As Object
    Dim lstrLog As String          '操作日志描述信息。
    Dim lstrInfo As String         '提示信息。
    
    On Error GoTo errHandler
    '创建体检对象。
    Set lobj体检 = New clsMedicalExam
    
    lobj体检.系统编号 = para系统编号
    
    '判断是否可以做取消操作。
    If lobj体检.体检状态 <> P_ENDED_STATUS Then
        Err.Raise 6666, , "该体检记录还未下过结论，不能作取消结论操作。"
    End If
        
    '判断是否已复查登记。
    If lobj体检.复查系统编号 <> "" Then
        '已复查登记。
        lstrInfo = "该体检人员已进行复查登记了，复查系统编号：" & lobj体检.复查系统编号 & "。"
        '询问。
        If Not sffuncMsg(lstrInfo & "你确认要继续取消该体检人员的体检结论吗？" & Chr(13) & Chr(10) & "    若你取消了该体检人员的体检结论，则原来的复查登记记录会被作废。", sf询问) Then
            '用户选择取消。
            Err.Raise 6666, , "用户选择终止取消结论操作。"
        End If
        
        Set lobj复查体检 = New clsMedicalExam
        lobj复查体检.系统编号 = lobj体检.复查系统编号
        
        If lobj复查体检.是否已存在 Then
            If lobj复查体检.体检状态 = P_ENDED_STATUS Then
                '复查登记已下结论，不允许取消。
                lstrInfo = lstrInfo & "但复查体检也已结论，不允许取消当前体检结论。若确实需要，请先取消复查登记的体检结论。"
                Err.Raise 6666, , lstrInfo
            End If
            
            '判断复查登记的体检状态和收费状态。
            If lobj复查体检.收费批号 <> "" Then
                Set lobjRec = mobj收费对外接口.func查询费用信息("收费编号='" & lobj复查体检.收费批号 & "'")
                If lobjRec.recordcount > 0 Then
                    If IIf(IsNull(lobjRec("收费状态")), 0, lobjRec("收费状态")) = 1 Then
                        '原有收费记录已收费。不允许取消结论。
                        lstrInfo = lstrInfo & "并且体检人员已缴费，收费编号：" & lobj复查体检.收费批号 & "。请先叫体检人员到收费科室退费，再作取消结论处理。"
                        Err.Raise 6666, , lstrInfo
                    Else
                        '未收费，或已退费。允许作取消结论。
                    End If
                End If
            End If
            
        End If
    End If
    
    '开始事务。
    dasubBeginTran
    
    '先删除旧的复查登记记录。
    If lobj体检.复查系统编号 <> "" Then
        dafuncGetData "delete 体检管理_体检基本信息表 where 系统编号='" & lobj体检.复查系统编号 & "'"
    End If
    
    '取消体检结论审核"。
    lobj体检.Sub取消体检结论审核
    
    '结束事务。
    dasubCommitTran
        
    '保存操作日志。
    On Error Resume Next
    lstrLog = "系统编号：" & lobj体检.系统编号 & ",姓名：" & lobj体检.体检人员.姓名 & ",所属单位：" & lobj体检.体检人员.单位名称 & "。"
    oesubSave lstrLog, "取消体检结论"
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub取消体检结论", 6666, lstrError, True
End Sub

'功能：为体检结论录入界面获取可以取消体检结论的体检记录。
'输入：查询条件。若任意一个参数为空，表示忽略该参数所代表的子条件。
'返回：Recordset[系统编号，姓名，性别，单位名称，体检表名，体检日期，体检日期，体检结论，下结论医师]。
'作者：杨春
'修改：2003-4-16（杨春）返回结果增加“卫生种类”，方便下结论时判断是否已录入卫生种类，否则信息无法流入健康证系统。
Public Function Func获取体检结论已确定的体检记录(ByVal para从系统编号 As String, _
                                                ByVal para到系统编号 As String, _
                                                ByVal para体检日期 As String) As Object
    Dim lobj体检集 As clsMedicalExamSet

    On Error GoTo errHandler
    '创建"体检集"对象。
    Set lobj体检集 = New clsMedicalExamSet
    
    '按照参数设置其过滤条件，并附加条件：体检状态=3，下结论医师=当前用户。
    With lobj体检集
        .从系统编号 = para从系统编号
        .到系统编号 = para到系统编号
        .从体检日期 = para体检日期
        If para体检日期 = "" Then
            .到体检日期 = ""
        Else
            .到体检日期 = para体检日期 & " 23:59:59"
        End If
        .下结论医师 = um用户编号
        .体检状态 = "3"
    End With
    
    '返回"体检集.元素集"。
    Set Func获取体检结论已确定的体检记录 = lobj体检集.元素集("系统编号,姓名,性别,年龄=datediff(year,出生日期,getdate()),单位名称,卫生种类,体检结论,诊断和处理意见,下结论医师,下结论医师姓名,体检日期,下结论日期,复查体检表名,体检表名称")
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Func获取体检结论已确定的体检记录", 6666, lstrError, True
    Exit Function
End Function

'功能：根据健康证条码号，调用健康证管理系统提供的对外接口获取最后一次体检的系统编号。
'输入：para健康证条码号   健康证上的条码号。
'返回：该健康证对应的最后次体检的系统编号。
'作者：杨春
Public Function Func根据健康证条码号获取体检系统编号(ByVal para健康证条码号 As String) As String
Attribute Func根据健康证条码号获取体检系统编号.VB_Description = "'功能：根据健康证条码号，调用健康证管理系统提供的对外接口获取最后一次体检的系统编号。\r\n    '输入：para健康证条码号   健康证上的条码号。\r\n    '返回：该健康证对应的最后次体检的系统编号。"
    Dim lobj健康证对外接口 As Object
    Dim lobjRec As Object
    
    '创建健康证对外接口。
    On Error Resume Next
    Set lobj健康证对外接口 = CreateObject("EMCI.ClsEmdicalCertificateInterface")
    If Err <> 0 Then
        Err.Clear
        On Error GoTo errHandler
        Err.Raise 6666, , "创建健康证对外接口部件的对象失败。请退出系统并重新注册“EMCI.dll”。"
    End If
    
    On Error GoTo errHandler
    '调用其方法：GetCertificateSN。
    Func根据健康证条码号获取体检系统编号 = lobj健康证对外接口.GetCertificateSN(para健康证条码号)
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Func根据健康证条码号获取体检系统编号", 6666, lstrError, True
End Function

'功能：使用通用文书管理工具打印指定文书（体检单、体检结果表、体检结果通知书），
'      若打印的是体检结果单，判断体检记录是否需要复查（或则时复查记录），
'      若需要，必须把复查记录和原记录的结果打在一起。
'      若创建通用文书对象失败，使用体检系统自定的文书管理部件打印指定文书。
'输入：para文书名称             要打印的体检文书名称；
'      para编号 as collection   指体检的系统编号的集合；
'      para是否预览 as boolean  打印时是否预览文书。
'      para是否只预览 as boolean 是否只预览，不打印。
'作者：杨春
Public Sub Sub打印文书(ByVal para文书名称 As String, _
                            ByVal para编号 As Collection, _
                            ByVal para是否预览 As Boolean, _
                            Optional para是否只预览 As Boolean = False)
    
    Dim lobjProgress As clsProgress '进度显示。
    Dim lobj体检 As clsMedicalExam
    Dim lobj文书 As Object          '文书管理提供的文书对象，用于打印文书。
    Dim lobjRec As Object           '调用获取文书数据的存储过程所返回的文书数据。
    Dim lstr文书名称 As String      '实际应打印的文书名称。
    Dim lstr系统编号 As String      '当前打印的体检系统编号。
    Dim lstr复查系统编号 As String  '当前系统编号对应的体检的复查系统编号。
    Dim lstrLog As String           '操作日志描述信息。
    Dim i As Long
    
    '创建"文书"对象。
    On Error Resume Next
    Set lobj文书 = CreateObject("文书管理.cls文书")
    If Err <> 0 Then
        '无法创建通用的“文书管理”的对象，按本子系统自定的文书打印程序打印文书。
        Err.Clear
        Set lobj文书 = CreateObject("体检文书管理.cls文书")
        If Err <> 0 Then
            On Error GoTo errHandler
            Err.Raise 6666, , "无法创建“体检文书管理”的对象。请重新注册“体检文书管理.dll”。"
        End If
    End If
    
    On Error GoTo errHandler
    '创建进度显示对象。
    If para编号.Count > 1 Then
        Set lobjProgress = New clsProgress
        With lobjProgress
            .FormTitle = "正在打印，请稍候..."
            .MaxValue = para编号.Count
            .Message = "开始打印..."
            .Value = 0
            .subShow
        End With
    End If
    
    '创建体检对象。
    Set lobj体检 = New clsMedicalExam
    
    '依次打印编号集合中的各编号对应的体检记录的文书。
    Do While para编号.Count > 0
        If Not lobjProgress Is Nothing Then
            lobjProgress.Message = "正在打印第" & lobjProgress.Value + 1 & "张（共" & lobjProgress.MaxValue & "张）..."
        End If
        
        lstr系统编号 = para编号(1)
        lobj体检.系统编号 = lstr系统编号
        
        '对"体检单"文书，创建"体检表"对象，获取对应的体检单名称作为真正的文书名称。
        If para文书名称 = "体检单" Then
            lstr文书名称 = lobj体检.体检表.体检单名称
        Else
            lstr文书名称 = para文书名称
        End If
        
        '调用文书对应的存储过程（存储过程名称："体检管理_获取"+文书名称+"内容"）获取文书数据。
        '    （注意：获取体检结果单的文书中，还要判断体检记录是否需要复查（或则时复查记录），
        '      若需要，必须把复查记录和原记录的结果一起返回）。
        If para文书名称 = "体检结果单" Then
            If lobj体检.复查系统编号 <> "" Then
                '从编号集合中去掉复查系统编号，以免重复打印。
                On Error Resume Next
                para编号.Remove lobj体检.复查系统编号
                On Error GoTo errHandler
            End If
        End If
        
        '修改：2003-4-16（杨春）若找不到存储过程，从视图中获取信息。
        On Error Resume Next
        Set lobjRec = dafuncGetData("exec 体检管理_获取" & para文书名称 & "内容 '" & lstr系统编号 & "'")
        If Err.Number <> 0 Then
            Set lobjRec = dafuncGetData("select * from 体检管理_体检基本数据库 where 系统编号='" & lstr系统编号 & "'")
        End If
        Err.Clear
        
        On Error GoTo errHandler
        
        '调用"文书.sub打印文书"。(体检单上打印条形码和照片)
        If lobjRec.recordcount = 0 Then
            '没有文书内容。
            Err.Raise 6666, , "获取系统编号为" & lstr系统编号 & "的体检文书内容为空，不能继续打印文书。"
        Else
            lobj文书.Sub打印文书 lstr文书名称, lobjRec, para是否预览, para是否只预览
            
            '保存操作日志。
            lstrLog = "文书名称：" & lstr文书名称 & ",系统编号：" & lstr系统编号 & ",姓名：" & lobj体检.体检人员.姓名 & ",所属单位：" & lobj体检.体检人员.单位名称 & "。"
            oesubSave lstrLog, "文书打印"
        End If
        
        '从编号集合中去掉当前系统编号。
        para编号.Remove 1
        
        '显示进度。
        If Not lobjProgress Is Nothing Then
            lobjProgress.Value = lobjProgress.Value + 1
        End If
    Loop
    
    If Not lobjProgress Is Nothing Then
        Set lobjProgress = Nothing
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub打印文书", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：获取本体检系统支持的体检单类型。
'作者：杨春
Public Function func获取所有体检单类型() As Collection
    Dim lobj文书 As Object
    
    On Error Resume Next
    '创建文书对象。
    Set lobj文书 = CreateObject("体检文书管理.cls文书")
    Set func获取所有体检单类型 = lobj文书.func获取所有体检单类型()
    
    If Err <> 0 Then
        Err.Clear
        On Error GoTo errHandler
        Set func获取所有体检单类型 = New Collection
        With func获取所有体检单类型
            .Add "尿常规体检单", "尿常规体检单"
            .Add "尿常规大便培养体检单", "尿常规大便培养体检单"
            .Add "尿常规大便培养验血体检单", "尿常规大便培养验血体检单"
        End With
    End If
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "func获取所有体检单类型", 6666, lstrError, True
    Exit Function
    Resume
End Function

'功能：修改指定业务配置参数值。
'输入：para配置项目   要修改的业务配置项目名称；
'      para项目值     修改后的业务配置项目值。
'作者：杨春
Public Sub Sub修改业务配置(ByVal para配置项目 As String, ByVal para项目值 As String)
Attribute Sub修改业务配置.VB_Description = "'功能：修改指定业务配置参数值。\r\n    '输入：para配置项目   要修改的业务配置项目名称；\r\n    '      para项目值     修改后的业务配置项目值。\r\n"
    Dim lobjRec As Object
    Dim lcolItem As Collection
    
    On Error GoTo errHandler
    
    '判断配置项目是否存在。
    Set lobjRec = dafuncGetData("select * from 体检管理_业务设置信息表 where 设置项目='" & para配置项目 & "'")
    If lobjRec.recordcount > 0 Then
        '已存在。
        '修改数据库中表"体检管理_业务设置表"中对应配置项目的配置值。
        dafuncGetData "update 体检管理_业务设置信息表 set 设置值='" & para项目值 & "' where 设置项目='" & para配置项目 & "'"
    Else
        '插入表"体检管理_业务设置表"中对应配置项目的配置值。
        dafuncGetData "insert into 体检管理_业务设置信息表(设置项目,设置值) values('" & para配置项目 & "','" & para项目值 & "')"
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub修改业务配置", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：添加、修改、删除体检附加项目。
'输入：para操作类型               1添加/2 修改/3 删除；
'      para项目信息 as collection[项目名称,录入标题,数据类型,数据长度,枚举值]；
'      para旧项目名称             修改操作，必须传递的参数。
'作者：杨春
Public Sub Sub设置体检附加项目(ByVal para操作类型 As Integer, _
                                ByVal para项目信息 As Collection, _
                                Optional para旧项目名称 As String = "")
    Dim lstrSql As String
    
    On Error GoTo errHandler
    
    '拼sql语句。
    Select Case para操作类型
    Case 1 '添加。
        lstrSql = "insert into 体检管理_体检人员附加项目设置表(附加项目,录入标题,数据类型,数据长度,枚举值) values("
        lstrSql = lstrSql & "'" & para项目信息("项目名称") & "','" & para项目信息("录入标题") & "'," & para项目信息("数据类型") & "," & para项目信息("数据长度") & ",'" & para项目信息("枚举值") & "')"
    Case 2 '修改。
        lstrSql = "update 体检管理_体检人员附加项目设置表 set "
        lstrSql = lstrSql & "附加项目='" & para项目信息("项目名称") & "',录入标题='" & para项目信息("录入标题") & "',数据类型=" & para项目信息("数据类型") & ",数据长度=" & para项目信息("数据长度") & ",枚举值='" & para项目信息("枚举值") & "'"
        lstrSql = lstrSql & " where 附加项目='" & para旧项目名称 & "'"
    Case 3 '删除。
        lstrSql = "delete 体检管理_体检人员附加项目设置表 where 附加项目='" & para旧项目名称 & "'"
    Case Else
        Err.Raise 6666, , "参数“操作类型”值非法。合法值只有：1，2，3。"
    End Select
    
    '执行sql语句。
    dafuncGetData lstrSql
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = Err.Description
    If Err.Number = P_ERRORCODE_VIOLATE Then
        If para操作类型 = 3 Then
            lstrError = "你删除的体检附加项目已被使用，不允许删除。"
        Else
            lstrError = "你设置的体检附加项目已存在。附加项目不允许重复。"
        End If
    Else
        lstrError = func错误处理(Err.Number, Err.Description)
    End If
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "Sub设置体检附加项目", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：调用单位档案对外接口提供的单位定位方法，启动单位定位界面。
'输入：para单位申请编号  可选参数。
'返回：Recordset[单位名称，地址，片区、卫生种类、行业类别]
'作者：杨春
Public Function func单位定位(Optional para单位申请编号 As String = "") As Object
    On Error GoTo errHandler
    
    If mobj单位对外接口 Is Nothing Then
        Err.Raise 6666, , "使用单位档案提供的单位定位接口失败。请退出系统，重新注册“单位档案对外接口.dll”。"
    End If
    
    Set func单位定位 = mobj单位对外接口.func单位简单定位(2330, 800)

    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsManageMedicalExam", "func单位定位", 6666, lstrError, True
End Function

'功能：根据单位申请编号获取单位属性。
'输入：para申请编号 as string
'返回：卫生种类,卫生种类名称,片区，片区名称，行业类别，行业类别名称，经济性质。
Public Function func获取单位属性(ByVal para申请编号 As String) As Collection
    Dim lobjRec As Object
    Dim lobjDict  As Object
    Dim lcolInfo As New Collection
    Dim lstrTemp As String
    
    On Error Resume Next
    
    '若没有给卫生种类等属性复制，根据单位申请编号获取这些属性。
    If para申请编号 <> "" Then
        Set lobjRec = dafuncGetData("select * from 单位档案_单位定位查询视图 where 申请编号='" & para申请编号 & "'")
        If lobjRec.recordcount > 0 Then
            '卫生种类。
            lstrTemp = ""
            lstrTemp = IIf(IsNull(lobjRec!卫生种类BH), "", lobjRec("卫生种类BH"))
            lcolInfo.Add lstrTemp, "卫生种类"
            lstrTemp = ""
            lstrTemp = IIf(IsNull(lobjRec!卫生种类), "", lobjRec!卫生种类)
            lcolInfo.Add lstrTemp, "卫生种类名称"
            
            '片区。
            lstrTemp = ""
            lstrTemp = IIf(IsNull(lobjRec!片区BH), "", lobjRec!片区BH)
            lcolInfo.Add lstrTemp, "片区"
            lstrTemp = ""
            lstrTemp = IIf(IsNull(lobjRec!片区), "", lobjRec!片区)
            lcolInfo.Add lstrTemp, "片区名称"
            
            lstrTemp = ""
            lstrTemp = IIf(IsNull(lobjRec!经济性质), "", lobjRec!经济性质)
            lcolInfo.Add lstrTemp, "经济性质名称"
            
            lstrTemp = ""
            lstrTemp = IIf(IsNull(lobjRec!行业类别), "", lobjRec!行业类别)
            lcolInfo.Add lstrTemp, "行业类别名称"
            
            '获取单位地址。
            lstrTemp = ""
            lstrTemp = IIf(IsNull(lobjRec!地址), "", lobjRec!地址)
            lcolInfo.Add lstrTemp, "单位地址"
            
            '创建字典对象。
            Set lobjDict = CreateObject("字典管理.clsDictionary")
            
            '获取行业类别编号。
            If lcolInfo("行业类别名称") <> "" Then
                lstrTemp = ""
                Set lobjRec = lobjDict.Fetch("行业类别字典视图", "名称='" & lcolInfo("行业类别名称") & "'")
                If Not lobjRec Is Nothing Then
                    If lobjRec.recordcount > 0 Then
                        lstrTemp = IIf(IsNull(lobjRec("编号")), "", lobjRec("编号"))
                    End If
                End If
                lcolInfo.Add lstrTemp, "行业类别"
                
            Else
                lcolInfo.Add "", "行业类别"
            End If
        
            '获取经济性质编号。
            If lcolInfo("经济性质名称") <> "" Then
                lstrTemp = ""
                Set lobjRec = lobjDict.Fetch("经济性质字典", "名称='" & lcolInfo("经济性质名称") & "'")
                If Not lobjRec Is Nothing Then
                    If lobjRec.recordcount > 0 Then
                        lstrTemp = IIf(IsNull(lobjRec("编号")), "", lobjRec("编号"))
                    End If
                End If
                lcolInfo.Add lstrTemp, "经济性质"
                
            Else
                lcolInfo.Add "", "经济性质"
            End If
        End If
                
    End If
    Set func获取单位属性 = lcolInfo
End Function
'功能：释放对象。
'作者：杨春
Private Sub Class_Terminate()
    On Error Resume Next
    Set mobj单位对外接口 = Nothing
    Set mobj当日工作记忆簿 = Nothing
    Set mobj工作站配置 = Nothing
    Set mobj收费对外接口 = Nothing
    
End Sub

'功能：获取指定收费标准的收费项目。
'返回：收费项目编号、收费项目名称、单价。
'创建：2001-12-29。
Public Property Get 收费标准的项目(ByVal para收费标准名 As String) As Object
    On Error GoTo errHandler
    Set 收费标准的项目 = dafuncGetData("select b.收费项目编号,b.收费项目名称,b.单价 from 收费管理_收费标准信息表 a,收费管理_收费项目字典表 b where a.收费项目编号=b.收费项目编号 and a.收费标准名称='" & para收费标准名 & "'")
    
    Exit Property
errHandler:
    Err.Raise 6666, , "无法获取指定收费标准的收费项目！" & Chr(13) & Chr(10) & "收费系统可能没有正确安装，请重新安装系统，或进入体检管理的业务设置，设置不要收费。"
End Property

'功能：删除指定系统编号的体检记录，若该体检记录对应人员已无其他体检记录，同时删除健康档案。
'创建：2003-6-26（杨春）。
Public Sub sub删除体检登记(ByVal para系统编号 As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    dasubBeginTran
    Set lobjRec = dafuncGetData("exec 体检管理_删除体检记录 '" + para系统编号 + "'")
    If lobjRec(0) = 1 Then
        Err.Raise 6666, , "该体检记录已下体检结论，不允许删除！"
    End If
    dasubCommitTran
    Exit Sub
errHandler:
    dasubRollBack
    Err.Raise 6666, , "删除体检登记失败！！" & Chr(13) & Chr(10) & Err.Description
End Sub
