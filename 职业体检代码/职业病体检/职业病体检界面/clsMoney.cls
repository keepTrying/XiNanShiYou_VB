VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public col体检项目 As Collection
Public mstr系统编号 As String
Public mstr体检人员姓名 As String

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
    dasubSetQueryTimeout 600
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
                    .Add mstr体检人员姓名, "交费人" 'lstr送样人
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
