VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCareerHstRegt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
'名称：职业病史(受检者个人信息)录入类模块
'函数：SubSaveSymptom()，SubSavePastMedcHst(),subSaveWorkHst()
'      sub修改体检状态()
'功能：职业病史(受检者个人信息)录入上信息保存至数据库；修改体检状态
'作者：Yunle Liu
'时间：2012.03
'********************************************************************

Option Explicit
Private mstr系统编号 As String
'职业史信息栏

Public mobjWorkHst As New Collection
Public mobjPastHst As New Collection
Public mobjSymptom As New Collection
Public col体检项目 As Collection

Public Property Let 系统编号(ByVal vNewValue As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    mstr系统编号 = vNewValue
    'If Len(mstr系统编号) > 5 Then
    '    Set lobjRec = dafuncGetData("select * from 职业病体检_职业史表 where 系统编号='" & mstr系统编号 & "'")
    '    If lobjRec.RecordCount > 0 Then
    '        MsgBox "已有该编号,你可以修改受检者个人信息信息！"
        'Else
        '    MsgBox "没找到该条码对应的信息，请检查是否已登记！"
    '        Exit Property
    '    End If
    '    lobjRec.Close
    'End If
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clscareerhstregt", "Property Let 系统编号", Err.Number, Err.Description, False
End Property

'保存 自觉症状
Public Sub SubSaveSymptom()
    Dim i As Integer
    Dim mbln是否已存在 As Boolean
    Dim lstrsql As String
    Dim lobjRec As Object
    Dim lobjdetail As clsSymptomDetl
    On Error GoTo errHandler
    mbln是否已存在 = False
    
    For i = 1 To mobjSymptom.Count
        Set lobjdetail = mobjSymptom(i)
        lstrsql = "insert into 职业病体检_自觉症状表(系统编号,编号,症状,程度,出现时间) values(" _
                & "'" & mstr系统编号 & "','" & lobjdetail.mstr编号 & "','" & lobjdetail.mstr症状 & "','" _
                & lobjdetail.mstr程度 & "','" & lobjdetail.mstr出现时间 & "')"
        dafuncGetData lstrsql
    Next
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub subsavesymptom", Err.Number, Err.Description, False
    Exit Sub
End Sub

'保存 既往病史
Public Sub SubSavePastMedcHst()
    Dim i As Integer
    Dim mbln是否已存在 As Boolean
    Dim lstrsql As String
    Dim lobjRec As Object
    Dim lobjdetail As clsPastMedcHstDetl
    On Error GoTo errHandler
    mbln是否已存在 = False
    
    For i = 1 To mobjPastHst.Count
        Set lobjdetail = mobjPastHst(i)
        lstrsql = "insert into 职业病体检_既往病史表(系统编号,编号,疾病名称,诊断日期,诊断单位,治疗经过,转归) values(" _
                & "'" & mstr系统编号 & "','" & lobjdetail.mstr编号 & "','" & lobjdetail.mstr疾病名称 & "','" _
                & lobjdetail.mstr诊断日期 & "','" & lobjdetail.mstr诊断单位 & "','" & lobjdetail.mstr治疗经过 & "','" _
                & lobjdetail.mstr转归 & "')"
        dafuncGetData lstrsql
    Next
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub subsavepastmedchst", Err.Number, Err.Description, False
End Sub

'保存 职业史
Public Sub subSaveWorkHst()
    Dim i As Integer
    Dim mbln是否已存在 As Boolean
    Dim lstrsql As String
    Dim lobjRec As Object
    Dim lobjdetail As ClsCareerHstDetl
    On Error GoTo errHandler
    mbln是否已存在 = False
    'If Len(Trim(mstr系统编号)) > 5 Then
    '    Set lobjrec = dafuncGetData("select * from 职业病体检_个人生活史表 where 系统编号='" & mstr系统编号 & "'")
    '    If lobjrec.RecordCount > 0 Then
    '        mbln是否已存在 = True
    '    End If
    'Else
    '    MsgBox "系统编号错误，请检查！", vbInformation, "系统提示"
    '    Exit Sub
    'End If
     '保存体检基本信息。
     For i = 1 To mobjWorkHst.Count
    'If mbln是否已存在 Then
        '已存在，修改记录。
        'lstrsql = "update 职业病体检_个人生活史表 set 试管编号='" & mstr试管编号 & "',体检表编号='" & mobj体检表.体检表名 & "'"
        '        & ",体检日期='" & mstr体检日期 & " ',收费批号='" & mstr收费批号 & "',体检类别='" & mlng体检人类别 & "',体检类型='" & mlng体检人类型 & "'" _
         '       & " where 系统编号='" & mstr系统编号 & "'"
    'Else
        '新增，插入记录。
        Set lobjdetail = mobjWorkHst(i)
        lstrsql = "insert into 职业病体检_职业史表(系统编号,编号,工作单位,部门,工种,危害种类,接触时间,防护措施,备注,放射种类,每日工作量,累积照射量,过量照射史,起始时间,结束时间,是否放射性) values(" _
                & "'" & mstr系统编号 & "','" & lobjdetail.mstr编号 & "','" & lobjdetail.mstr单位 & "','" & lobjdetail.mstr部门 & "','" & lobjdetail.mstr工种 & "','" & lobjdetail.mstr危害种类 & "','" _
                & lobjdetail.mstr接触时间 & "','" & lobjdetail.mstr措施 & "','" & lobjdetail.mstr备注 & "','" & lobjdetail.mstr放射种类 & "','" & lobjdetail.mstr工作量 & "','" _
                & lobjdetail.mstr照射量 & "','" & lobjdetail.mstr过量照射史 & "','" & lobjdetail.mstr起始时间 & "','" & lobjdetail.mstr结束时间 & "','" & lobjdetail.mstr是否放射性 & "')"
    'End If
        dafuncGetData lstrsql
    Next
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub subsaveworkhst", Err.Number, Err.Description, False
End Sub

'修改体检状态
Public Sub sub修改体检状态()
    Dim lstrsql As String
    On Error GoTo errHandler
    lstrsql = "update 职业病体检_体检基本信息表 set 体检状态=1 where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrsql
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub sub修改体检状态", Err.Number, Err.Description, False
End Sub

'删除  职业史
Public Sub subDelWorkHst()
    Dim lstrsql As String
    On Error GoTo errHandler
    lstrsql = "delete from 职业病体检_职业史表 where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrsql
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub subdelworkhst", Err.Number, Err.Description, False
End Sub

'删除  既往病史
Public Sub subDelPastMedcHst()
    Dim lstrsql As String
    On Error GoTo errHandler
    lstrsql = "delete from 职业病体检_既往病史表 where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrsql
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub subdelpastmedchst", Err.Number, Err.Description, False
End Sub

'删除  自觉症状
Public Sub subDelSymptom()
    Dim lstrsql As String
    On Error GoTo errHandler
    lstrsql = "delete from 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrsql
    Exit Sub
errHandler:
   sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub subdelsymptom", Err.Number, Err.Description, False
End Sub

'功能：保存职业病史里选择的体检项目
'作者：翁乔
'时间：2012-04-14
'说明：首先要查看数据库里面是否有相同的体检项目，然后再进行增加或者修改

Public Sub save优化的体检项目()
    Dim lstrsql As String
    Dim MedicProjt As String
    Dim rs As Object
    Dim i As Integer
    On Error GoTo errHandler
    
    Set rs = dafuncGetData("select 名称 from 系统管理_字典_字典内容表 where ID = (select id from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 名称 like '%科'")
    
    For i = 1 To rs.RecordCount
        
        lstrsql = "delete 职业病体检_结果信息_" & rs("名称") & " where 系统编号='" & mstr系统编号 & "'"
        dafuncGetData lstrsql
        rs.MoveNext
    Next i
    
    For i = 1 To col体检项目.Count
        MedicProjt = Left(Trim(col体检项目(i)("编码")), 2)
        
        lstrsql = "select 名称 from 系统管理_字典_字典内容表 where ID = (select id from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 编号= '" & MedicProjt & "'"
        Set rs = dafuncGetData(lstrsql)
        
        lstrsql = "insert into 职业病体检_结果信息_" & rs("名称") & "(系统编号,体检项目) values(" _
            & "'" & mstr系统编号 & "','" & col体检项目(i)("编码") & "')"
        dafuncGetData lstrsql
    Next i
    
    
    Exit Sub
errHandler:
   sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub save体检项目", Err.Number, Err.Description, False
End Sub

'保存体检项目  职业病史录入后
Public Sub save体检项目()
    Dim lstrsql As String
    Dim MedicProjt As String
    Dim SqlSentence As String
    Dim i As Integer
    On Error GoTo errHandler
    lstrsql = "delete 职业病体检_体检结果信息表 where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrsql
    For i = 1 To col体检项目.Count
        MedicProjt = Left(Trim(col体检项目(i)("编码")), 2)
        Select Case MedicProjt
        Case "01"
            SqlSentence = "结果信息_五官科"
        Case "02"
            SqlSentence = "结果信息_外科"
        Case "03"
            SqlSentence = "结果信息_内科"
        Case Else
            SqlSentence = "体检结果信息表"
        End Select
        lstrsql = "insert into 职业病体检_" & SqlSentence & "(系统编号,体检项目) values(" _
            & "'" & mstr系统编号 & "','" & col体检项目(i)("编码") & "')"
        dafuncGetData lstrsql
    Next i
    Exit Sub
errHandler:
   sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub save体检项目", Err.Number, Err.Description, False
End Sub

'功能：对指定体检记录进行划价，若划价成功，生成对应的费用信息，并传递给收费系统。
'输入：para监测系统编号 as String
'
'输出：para收费批号 as string
'返回：错误信息（为空表示成功,Cancel表示没有保存）。
'创建：2012-03  liuyunle
Public Function func收费(para收费批号 As String) As String
    'Dim lobj业务设置 As cls卫生监测设置
    Dim lobj收费管理 As Object
    Dim lobj收费对外接口 As Object
    Dim lobjRec As Object

    Dim lstr收费状态 As String
    Dim lstr收费批号 As String
    Dim lstr送样人 As String
    Dim lstrLog As String
    Dim lsngTotal As Single '总金额。

    Dim lcol划价结果 As Collection
    Dim lcol费用信息 As Collection
    Dim lcolItem As Collection

    Dim i As Long

    On Error GoTo errHandler
    func收费 = ""
    para收费批号 = ""

    '创建业务设置对象。
    'Set lobj业务设置 = New cls卫生监测设置

    '判断业务设置"使用道源收费系统"，若设置值不为"是"，
    '上举错误"业务设置了不使用道源收费系统，所以无法进行收费。若要收费，请…"。
    'If lobj业务设置.业务设置("使用收费系统") <> "是" Then
    '    Err.Raise 6666, , "业务设置了不使用收费系统，所以无法进行收费。若要收费，请进入“检验管理_业务设置”选中“使用收费系统”，并保存设置后再试。"
    'End If

    '创建"检验收费管理.cls检验收费管理"对象，和"收费接口.cls对外接口" 。
   
    Set lobj收费管理 = CreateObject("职业病对象.clstestitemset")
    
    Set lobj收费对外接口 = CreateObject("收费接口对象.cls对外接口")
    
    '判断是否已有收费记录存在。
    lstr收费批号 = ""
    Set lobjRec = dafuncGetData("select 收费批号 from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
    If lobjRec.RecordCount = 0 Then
        MsgBox "没找到该体检人员收费记录，退出后检查！"
        Exit Function
    ElseIf IIf(IsNull(lobjRec("收费批号")), "", lobjRec("收费批号")) = "" Then
        '该次检验还没有收费信息，调用lobj收费管理. Func检验划价，若划价结果不为空，按收费接口要求组装收费信息集，并调用lobj收费对外接口. func划价_数据集合；
        lstr收费状态 = ""
    Else
        lstr收费批号 = lobjRec("收费批号")
        'lstr送样人 = IIf(IsNull(lobjRec("采样人")), "", lobjRec("采样人"))

        '已收费，获取收费状态。
        Set lobjRec = lobj收费对外接口.func查询费用信息("收费编号='" & lstr收费批号 & "'")
        lstr收费状态 = ""
        If lobjRec.RecordCount = 0 Then
'            lstr收费状态 = "0"
            '已报废，应该重新计费。
            lstr收费批号 = ""
        Else
            lstr收费状态 = IIf(IsNull(lobjRec("收费状态")), 0, lobjRec("收费状态"))
        End If
    End If

    
    If lstr收费状态 = "1" Then
        '已收费，返回错误信息。
        Err.Raise 6666, , "该次检验已收费完毕，不能重新收费。解决办法：" & Chr(13) & Chr(10) & "请财务科操作员对该次检验的收费记录进行退费处理后才可以重新划价。" & Chr(13) & Chr(10) & "其收费编号为：" & lstr收费批号
    Else
        '启动划价界面。
       
        Set lcol划价结果 = lobj收费管理.fun划价(mstr系统编号)
        
        '若划价结果不为空，处理收费信息。
        If lcol划价结果.Count > 0 Then
            lsngTotal = 0

            '按收费接口要求组装收费信息集。
            '交费人、交费单位编号、交费单位名称、主管科室编号、主管科室经手人、收费项目编号、单价、数量。
            Set lcol费用信息 = New Collection
            For i = 1 To lcol划价结果.Count
                Set lcolItem = New Collection
                With lcolItem
                    .Add "", "交费人" 'lstr送样人
                    .Add lcol划价结果(i)("单位编号"), "交费单位编号"
                    .Add lcol划价结果(i)("单位名称"), "交费单位名称"
                    .Add um用户编号, "主管科室经手人"
                    .Add um用户所属科室编号, "主管科室编号"
                    .Add lcol划价结果(i)("体检项目"), "收费项目编号"
                    .Add lcol划价结果(i)("单价"), "单价"
                    .Add lcol划价结果(i)("数量"), "数量"
                    .Add lcol划价结果(i)("体检类型"), "备注1"
                    .Add lcol划价结果(i)("体检类别"), "备注2"
                End With

                lcol费用信息.Add lcolItem, lcolItem("收费项目编号")

                '累计总额。
                lsngTotal = Format(lsngTotal + lcol划价结果(i)("单价") * lcol划价结果(i)("数量"), "0.00")
            Next


            If lstr收费状态 = "" Or lstr收费状态 = "2" Then
                '开始事务。
                dasubBeginTran

                '该次检验还没有收费信息，或则已退费。新增收费记录。
                '增加参数“业务分类”。
                lstr收费批号 = lobj收费对外接口.func划价_数据集合(lcol费用信息, , False, lcol划价结果(1)("单位编号"), "职业病体检收费")

                '输出收费批号。
                para收费批号 = Right(lstr收费批号, Len(lstr收费批号) - InStr(lstr收费批号, ";"))

                '修改本检验记录的收费编号。
                dafuncGetData "update 职业病体检_体检基本信息表 set 收费批号='" & para收费批号 & "',收费金额=" & lsngTotal & " where 系统编号='" & mstr系统编号 & "'"

                '提交事务。
                dasubCommitTran

                '拼操作日志。
                lstrLog = "收费：监测系统编号=" & mstr系统编号 & "，交费单位：" & lcol划价结果(1)("单位名称") & "，收费批号=" & para收费批号 & "，总金额=" & lsngTotal & "。"

            ElseIf lstr收费状态 = "0" Then
                '还未收费，修改收费记录。
                '增加参数“业务分类”。
                Call lobj收费对外接口.func划价_数据集合(lcol费用信息, lstr收费批号, False, lcol划价结果(1)("单位编号"), "职业病体检收费")

                '输出收费批号。
                para收费批号 = lstr收费批号

                '拼操作日志。
                lstrLog = "修改收费记录：职业病体检系统编号=" & mstr系统编号 & "，交费单位：" & lcol划价结果(1)("单位名称") & "，收费批号=" & lstr收费批号 & "，新总金额=" & lsngTotal & "。"

            End If

            On Error Resume Next
            '登记操作日志。
            oesubSave lstrLog, "职业病体检收费"
        Else
            func收费 = "Cancel"
        End If
    End If


    Exit Function

errHandler:
    func收费 = Err.Description
    dasubRollBack
    Exit Function
    Resume
End Function

Public Function func获取体检人员的体检项目(ByVal para系统编号 As String) As Collection

    Dim lobjRec As Object
    Dim lcolInfo As Collection
    Dim lcolItem As Collection
    Dim lstrItem As String
    Dim lstrsql As String
    
    On Error GoTo errHandler
    '获取选择的项目。
    
    lstrsql = "select b.编码,b.名称 from 职业病体检_体检结果视图 a, 职业病体检_体检项目设置表 b where a.体检项目=b.编码 and a.系统编号='" & 1112060410001# & "'"
    
    Set lobjRec = dafuncGetData(lstrsql)
    
    Set lcolInfo = New Collection
    While lobjRec.EOF
        lstrItem = lobjRec("编码")
        Set lcolItem = New Collection
        lcolItem.Add lstrItem, "编码"
            
        lcolItem.Add lobjRec("名称"), "名称"
        lcolInfo.Add lcolItem, lstrItem
    Wend
    If lcolInfo.Count = 0 Then
        sffuncMsg "必须选择体检项目！", sf警告
        Set lcolInfo = Nothing
        Exit Function
    End If
    
    Set func获取体检人员的体检项目 = lcolInfo
    
    Set lobjRec = Nothing
    
    Exit Function
    
errHandler:
    dasubRollBack
    Exit Function
    Resume

End Function

'2012-06-13 于登淼
'已知体检项目名称获取体检项目编号
Public Function func获取体检项目编号(ByVal paraItemName As String, Optional ByVal paraDept As String = "") As Object
    Dim paraItemNo As String
    Dim lobjTemp As Object
    Dim strSQL As String
    dasubSetQueryTimeout 6000
    strSQL = "select * from 职业病体检_体检项目设置表 where 名称='" & paraItemName & "' and left(编码,Len('" & paraDept & "'))='" & paraDept & "'"
    Set lobjTemp = dafuncGetData(strSQL)
    Set func获取体检项目编号 = lobjTemp
End Function

'2012-06-13 于登淼
'已知单项体检项目编号和单个体检人员的体检编号，获取其对应的的结果
Public Function func获取单人单项体检结果(ByVal paraSysNo As String, ByVal paraItemNo As String) As Object
    Dim deptName As String
    Dim strSQL As String
    Dim lobjRec As Object
    
    Select Case Left(paraItemNo, 2)
    Case "01": deptName = "五官科"
    Case "02": deptName = "内科"
    Case "03": deptName = "外科"
    Case "04": deptName = "血常规化验科"
    Case "05": deptName = "免疫科"
    Case "06": deptName = "尿常规化验科"
    Case "07": deptName = "染色体化验科"
    Case "08": deptName = "电测听科"
    Case "09": deptName = "X光影像科"
    Case "10": deptName = "心电科"
    Case "11": deptName = "B超影像科"
    Case "12": deptName = "肺功能影像科"
    Case "13": deptName = "受检者个人信息录入科"
    Case "17": deptName = "生化科"
    End Select
    
    strSQL = "select * from 职业病体检_结果信息_" & deptName & " where 系统编号='" & paraSysNo & "' and 体检项目='" & paraItemNo & "'"
'    dasubBeginTran
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(strSQL)
'    dasubCommitTran
    Set func获取单人单项体检结果 = lobjRec
    If lobjRec.RecordCount = 0 Then Set func获取单人单项体检结果 = Nothing
End Function

'2012-06-13 于登淼
'已知单项体检项目编号和单个体检人员的体检编号，保存其对应的的结果
Public Function func保存单人单项体检结果(ByVal paraSysNo As String, _
                                        ByVal paraDept As String, _
                                        ByVal paraItemNo As String, _
                                        ByVal paraItemResult As String, _
                                        Optional ByVal paraConclusion As String)
    Dim deptName As String
    Dim strSQL As String
    Dim lobjRec As Object
    
    Select Case paraDept
    Case "01": deptName = "五官科"
    Case "02": deptName = "内科"
    Case "03": deptName = "外科"
    Case "04": deptName = "血常规化验科"
    Case "05": deptName = "免疫科"
    Case "06": deptName = "尿常规化验科"
    Case "07": deptName = "染色体化验科"
    Case "08": deptName = "电测听科"
    Case "09": deptName = "X光影像科"
    Case "10": deptName = "心电科"
    Case "11": deptName = "B超影像科"
    Case "12": deptName = "肺功能影像科"
    Case "13": deptName = "受检者个人信息录入科"
    Case "17": deptName = "生化科"
    End Select
    
    strSQL = "select * from 职业病体检_结果信息_" & deptName & " where 系统编号='" & paraSysNo & "' and 体检项目='" & paraItemNo & "'"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount = 0 Then
        '2012-07-05 于登淼 ↓
        '如果选择项目的时候，没有将体检项加入到结果信息表中，则不保存该结果。（主要配合自由修改体检项目的需求）
        'strSQL = "insert into 职业病体检_结果信息_" & deptName & " values('" & paraSysNo & "','" & paraItemNo & "','" _
                    & paraItemResult & "','" & um用户编号 & "','" & date & "','" & paraConclusion & "')"
        '2012-07-05 于登淼 ↑
    Else
        strSQL = "update 职业病体检_结果信息_" & deptName & " set 体检结果='" & paraItemResult & _
                "',体检医师='" & um用户编号 & "', 填写时间='" & Date & "', 单项结论='" _
                & paraConclusion & "' where 系统编号='" & paraSysNo & "' and 体检项目='" & paraItemNo & "'"
    End If
    dasubBeginTran
    dafuncGetData (strSQL)
    dasubCommitTran
End Function

'2012-06-13 于登淼
'已知单项体检项目编号和单个体检人员的体检编号，保存其对应的的结果
Public Function func删除单人单项体检结果(ByVal paraSysNo As String, _
                                        ByVal paraDept As String, _
                                        ByVal paraItemNo As String)
    Dim deptName As String
    Dim strSQL As String
    
    Select Case paraDept
    Case "01": deptName = "五官科"
    Case "02": deptName = "内科"
    Case "03": deptName = "外科"
    Case "04": deptName = "血常规化验科"
    Case "05": deptName = "免疫科"
    Case "06": deptName = "尿常规化验科"
    Case "07": deptName = "染色体化验科"
    Case "08": deptName = "电测听科"
    Case "09": deptName = "X光影像科"
    Case "10": deptName = "心电科"
    Case "11": deptName = "B超影像科"
    Case "12": deptName = "肺功能影像科"
    Case "13": deptName = "受检者个人信息录入科"
    Case "17": deptName = "生化科"
    End Select
    
    strSQL = "delete from 职业病体检_结果信息_" & deptName & " where 系统编号='" & paraSysNo & "' and 体检项目='" & paraItemNo & "'"
    dasubBeginTran
    dafuncGetData (strSQL)
    dasubCommitTran
End Function
