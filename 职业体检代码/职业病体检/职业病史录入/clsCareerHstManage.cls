VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCareerHstMage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************
'名称：职业病史(受检者个人信息)管理
'函数：
'功能：起动界面
'作者：Yunle Liu
'时间：2012.03
'***************************************


Option Explicit
Private mstr系统编号  As String

'功能：启动本部件所包括的一级操作界面，
'      按照总体结构的要求，必须无模式启动界面，并指定其parent指向主导航界面，并把界面移动到指定位置。
'输入：para操作名称      要启动的界面对应的操作名称。
'返回：界面的句柄（若指定操作不存在，返回-1）。
'修改：2001-12-27（初检登记、年检登记合并为“体检登记”。
Public Function funcStart(ByVal para操作名称 As String) As Long
    Dim lfrmForm As Form '待打开的窗体
    
    On Error GoTo errHandler
    Select Case para操作名称
    'Case "体检登记", "体检管理_体检登记"
    '    Set lfrmForm = frmRegisterManage
        
    'Case "体检管理_体检结果录入"
    '    Set lfrmForm = FrmInputTestResult
    
    'Case "体检管理_体检结论录入"
    '    Set lfrmForm = frmGiveConclusion
        
    'Case "文书打印", "体检管理_文书打印"
    '    Set lfrmForm = frmPrintPaper
    Case "职业病体检_受检者个人信息录入科"
        Set lfrmForm = frmCareerHstMage
'    Case "体检公众查询", "体检管理_体检公众查询"
'        Set lfrmForm = frmPublicQueryResult
'        funcStart = lfrmForm.hWnd
'        lfrmForm.Show 1
'        Exit Function
    
    Case Else
        funcStart = -1
        Err.Raise 6666, , "当前体检管理版本不支持操作“" & para操作名称 & "”。"
    End Select
    '如果找到窗体则打开窗体并调整窗体位置，返回窗体句柄。
    If funcStart <> -1 Then
        If lfrmForm.pblnInUse Then
            funcStart = -2
        Else
            funcStart = lfrmForm.hWnd
        End If
'        lfrmForm.ZOrder 0
        lfrmForm.Show
        lfrmForm.Move 700, 350 '1560, 0
        
        Set lfrmForm = Nothing
    End If
    Exit Function

errHandler:
    Set lfrmForm = Nothing
    sfsub错误处理 "体检界面部件", "clsManageTestForm", "subStart", Err.Number, Err.Description, True
End Function


Public Property Let 系统编号(ByVal vNewValue As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    mstr系统编号 = vNewValue
    
        'Set lobjRec = dafuncGetData("select * from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "' and not 体检状态='" & "0'")
        'If lobjRec.RecordCount > 0 Then
            'MsgBox "该体检人员已进行受检者个人信息登记，是否要修改它？"
            'If MsgBox("该体检人员已进行受检者个人信息登记，是否要修改它？", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
            'End If
            'Exit Property
        'End If
        'lobjRec.Close
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clscarrerhstmage", "Property Let 系统编号", Err.Number, Err.Description, True
End Property

'获取个人生活史
Public Property Get 个人生活史() As Object
    Dim lobjRec As Object
    On Error GoTo errHandler
'        Set lobjRec = dafuncGetData("select * from 职业病体检_个人生活史表 where 系统编号='" & mstr系统编号 & "'")
'        Set 个人生活史 = lobjRec
        '个人生活史 判断有没有以前来录过信息，有的话直接导入 2016-6-2 by 牟俊 ↓
        Dim id As String
        Dim xitong As String
        Dim dan As String
        Dim syno As String
        Dim lobj1 As Object
        Dim lobj2 As Object
        Dim obj As Object
        Dim kk As Integer
        Set lobj1 = dafuncGetData("select 公民身份号码 from 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
        id = lobj1("公民身份号码")
        Set lobj2 = dafuncGetData("select * from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
        xitong = lobj2("各科体检状态")
        dan = Mid(xitong, 13, 1)
        If dan = 1 Then
            Set obj = dafuncGetData("select * from 职业病体检_体检人员基本信息表 where 公民身份号码='" & id & "' order by 建档日期 desc")
            If obj.RecordCount > 1 Then
            obj.MoveFirst
                For kk = obj.RecordCount - 1 To obj.RecordCount
                    syno = obj("系统编号")
            obj.MoveNext
                Next
                Set lobjRec = dafuncGetData("select * from 职业病体检_个人生活史表 where 系统编号='" & syno & "'")
                Set 个人生活史 = lobjRec
            Else
                Set lobjRec = dafuncGetData("select * from 职业病体检_个人生活史表 where 系统编号='" & mstr系统编号 & "'")
                Set 个人生活史 = lobjRec
            End If
        Else
            Set lobjRec = dafuncGetData("select * from 职业病体检_个人生活史表 where 系统编号='" & mstr系统编号 & "'")
            Set 个人生活史 = lobjRec
        End If
        '2016-6-2 by 牟俊 ↑
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clscarrerhstmage", "Property get 个人生活史", Err.Number, Err.Description, True
End Property


'获取职业史
Public Property Get 职业史() As Object
    Dim lobjRec As Object
    On Error GoTo errHandler
        dasubSetQueryTimeout 6000
'        Set lobjRec = dafuncGetData("select 系统编号,编号,工作单位,部门,工种,起始时间,结束时间,危害种类,防护措施,备注,接触时间,放射种类,每日工作量,累积照射量,过量照射史,是否放射性 from 职业病体检_职业史表 where 系统编号='" & mstr系统编号 & "'")
'        Set 职业史 = lobjRec
        '职业史 判断有没有以前来录过信息，有的话直接导入 2016-6-2 by 牟俊 ↓
        Dim id As String
        Dim xitong As String
        Dim dan As String
        Dim syno As String
        Dim lobj1 As Object
        Dim lobj2 As Object
        Dim obj As Object
        Dim kk As Integer
        Set lobj1 = dafuncGetData("select 公民身份号码 from 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
        id = lobj1("公民身份号码")
        Set lobj2 = dafuncGetData("select * from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
        xitong = lobj2("各科体检状态")
        dan = Mid(xitong, 13, 1)
        If dan = 1 Then
            Set obj = dafuncGetData("select * from 职业病体检_体检人员基本信息表 where 公民身份号码='" & id & "' order by 建档日期 desc")
            If obj.RecordCount > 1 Then
            obj.MoveFirst
                For kk = obj.RecordCount - 1 To obj.RecordCount
                    syno = obj("系统编号")
            obj.MoveNext
                Next
                Set lobjRec = dafuncGetData("select 系统编号,编号,工作单位,部门,工种,起始时间,结束时间,危害种类,防护措施,备注,接触时间,放射种类,每日工作量,累积照射量,过量照射史,是否放射性 from 职业病体检_职业史表 where 系统编号='" & syno & "'")
                Set 职业史 = lobjRec
            Else
                Set lobjRec = dafuncGetData("select 系统编号,编号,工作单位,部门,工种,起始时间,结束时间,危害种类,防护措施,备注,接触时间,放射种类,每日工作量,累积照射量,过量照射史,是否放射性 from 职业病体检_职业史表 where 系统编号='" & mstr系统编号 & "'")
                Set 职业史 = lobjRec
            End If
        Else
            Set lobjRec = dafuncGetData("select 系统编号,编号,工作单位,部门,工种,起始时间,结束时间,危害种类,防护措施,备注,接触时间,放射种类,每日工作量,累积照射量,过量照射史,是否放射性 from 职业病体检_职业史表 where 系统编号='" & mstr系统编号 & "'")
            Set 职业史 = lobjRec
        End If
        '2016-6-2 by 牟俊 ↑
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clscarrerhstmage", "Property get 职业史", Err.Number, Err.Description, True
End Property

'获取既往病史
Public Property Get 既往病史() As Object
    Dim lobjRec As Object
    On Error GoTo errHandler
'        Set lobjRec = dafuncGetData("select * from 职业病体检_既往病史表 where 系统编号='" & mstr系统编号 & "'")
'        Set 既往病史 = lobjRec
        '既往病史 判断有没有以前来录过信息，有的话直接导入 2016-6-2 by 牟俊 ↓
        Dim id As String
        Dim xitong As String
        Dim dan As String
        Dim syno As String
        Dim lobj1 As Object
        Dim lobj2 As Object
        Dim obj As Object
        Dim kk As Integer
        Set lobj1 = dafuncGetData("select 公民身份号码 from 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
        id = lobj1("公民身份号码")
        Set lobj2 = dafuncGetData("select * from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
        xitong = lobj2("各科体检状态")
        dan = Mid(xitong, 13, 1)
        If dan = 1 Then
            Set obj = dafuncGetData("select * from 职业病体检_体检人员基本信息表 where 公民身份号码='" & id & "' order by 建档日期 desc")
            If obj.RecordCount > 1 Then
            obj.MoveFirst
                For kk = obj.RecordCount - 1 To obj.RecordCount
                    syno = obj("系统编号")
            obj.MoveNext
                Next
                Set lobjRec = dafuncGetData("select * from 职业病体检_既往病史表 where 系统编号='" & syno & "'")
                Set 既往病史 = lobjRec
            Else
                Set lobjRec = dafuncGetData("select * from 职业病体检_既往病史表 where 系统编号='" & mstr系统编号 & "'")
                Set 既往病史 = lobjRec
            End If
        Else
            Set lobjRec = dafuncGetData("select * from 职业病体检_既往病史表 where 系统编号='" & mstr系统编号 & "'")
            Set 既往病史 = lobjRec
        End If
        '2016-6-2 by 牟俊 ↑
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clscarrerhstmage", "Property get 既往病史", Err.Number, Err.Description, True
End Property

'获取自觉症状
Public Property Get 自觉症状() As Object
    Dim lobjRec As Object
    On Error GoTo errHandler
        Set lobjRec = dafuncGetData("select 系统编号,编号,症状,程度,出现时间 from 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'")
        Set 自觉症状 = lobjRec
'        '自觉症状 判断有没有以前来录过信息，有的话直接导入 2016-6-2 by 牟俊 ↓
'        Dim id As String
'        Dim xitong As String
'        Dim dan As String
'        Dim syno As String
'        Dim lobj1 As Object
'        Dim lobj2 As Object
'        Dim obj As Object
'        Dim kk As Integer
'        Set lobj1 = dafuncGetData("select 公民身份号码 from 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
'        id = lobj1("公民身份号码")
'        Set lobj2 = dafuncGetData("select * from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
'        xitong = lobj2("各科体检状态")
'        dan = Mid(xitong, 13, 1)
'        If dan = 1 Then
'            Set obj = dafuncGetData("select * from 职业病体检_体检人员基本信息表 where 公民身份号码='" & id & "' order by 建档日期 desc")
'            If obj.RecordCount > 1 Then
'            obj.MoveFirst
'                For kk = obj.RecordCount - 1 To obj.RecordCount
'                    syno = obj("系统编号")
'            obj.MoveNext
'                Next
'                Set lobjRec = dafuncGetData("select 系统编号,编号,症状,程度,出现时间 from 职业病体检_自觉症状表 where 系统编号='" & syno & "'")
'                Set 自觉症状 = lobjRec
'            Else
'                Set lobjRec = dafuncGetData("select 系统编号,编号,症状,程度,出现时间 from 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'")
'                Set 自觉症状 = lobjRec
'            End If
'        Else
'            Set lobjRec = dafuncGetData("select 系统编号,编号,症状,程度,出现时间 from 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'")
'            Set 自觉症状 = lobjRec
'        End If
'        '2016-6-2 by 牟俊 ↑
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clscarrerhstmage", "Property get 自觉症状", Err.Number, Err.Description, True
End Property

'获取是否已进行受检者个人信息记录
Public Property Get 已记录标志() As String
    Dim lobjRec As Object
    Dim tmp As String
    On Error GoTo errHandler
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select 各科体检状态 from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "' and 体检状态 in('0','1','2','3','4','5')")
    If Not (lobjRec.EOF Or lobjRec.BOF) Then
        If lobjRec.RecordCount = 1 Then
            tmp = Mid(lobjRec(0), 13, 1)
            If tmp = "0" Then
                已记录标志 = "100"
            ElseIf tmp = "1" Then
                已记录标志 = "1"
            Else
                已记录标志 = "2"
            End If
        Else
            已记录标志 = "101"
        End If
    Else
        已记录标志 = "100"
    End If
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clscarrerhstmage", "Property get 已记录标志", Err.Number, Err.Description, True
End Property
