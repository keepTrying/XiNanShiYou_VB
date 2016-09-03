VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'2012-03-01 于登淼
'增加结果录入内所有函数功能

Option Explicit

Public Function funcStart(ByVal para操作名称 As String) As Long
    Dim lfrmForm As Form '待打开的窗体
    
    On Error GoTo errHandler
    Select Case para操作名称
    Case "职业病体检_五官科结果录入"
        InputFlag = "五官科"
        InputFlagNo = "01"
'        frmResultInput_Routine.Caption = "五官科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_X光影像科结果录入"
        InputFlag = "X光影像科"
        InputFlagNo = "09"
'        frmResultInput_Assay.Caption = "X光影像科结果录入"
'        Set lfrmForm = frmResultInput_Assay
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_B超影像科结果录入"
        InputFlag = "B超影像科"
        InputFlagNo = "11"
'        frmResultInput_Assay.Caption = "B超影像科结果录入"
'        Set lfrmForm = frmResultInput_Assay
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_内科结果录入"
        InputFlag = "内科"
        InputFlagNo = "02"
'        frmResultInput_Routine.Caption = "内科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_外科结果录入"
        InputFlag = "外科"
        InputFlagNo = "03"
'        frmResultInput_Routine.Caption = "外科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_免疫科结果录入"
        InputFlag = "免疫科"
        InputFlagNo = "05"
'        frmResultInput_Routine.Caption = "免疫科结果录入"
        Set lfrmForm = frmResultInput_Routine
    
    Case "职业病体检_生化科结果录入"
        InputFlag = "生化科"
        InputFlagNo = "17"
'        frmResultInput_Routine.Caption = "生化科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_心电科结果录入"
        InputFlag = "心电科"
        InputFlagNo = "10"
'        frmResultInput_Routine.Caption = "心电科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_肺功能影像科结果录入"
        InputFlag = "肺功能影像科"
        InputFlagNo = "12"
'        frmResultInput_Routine.Caption = "肺功能影像科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_尿常规化验科结果录入"
        InputFlag = "尿常规化验科"
        InputFlagNo = "06"
'        frmResultInput_Routine.Caption = "尿常规化验科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_染色体化验科结果录入"
        InputFlag = "染色体化验科"
        InputFlagNo = "07"
'        frmResultInput_Routine.Caption = "染色体化验科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_电测听科结果录入"
        InputFlag = "电测听科"
        InputFlagNo = "08"
'        frmResultInput_Routine.Caption = "电测听科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_血常规化验科结果录入"
        InputFlag = "血常规化验科"
        InputFlagNo = "04"
'        frmResultInput_Routine.Caption = "血常规化验科结果录入"
        Set lfrmForm = frmResultInput_Routine
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
'            lfrmForm.pblnInUse = True
        End If
        lfrmForm.Show
        lfrmForm.Move 700, 350 '1560, 0
        Set lfrmForm = Nothing
    End If
    
    Exit Function
errHandler:
    Set lfrmForm = Nothing
    sfsub错误处理 "职业病体检结果录入", "clsCommon", "funcStart", Err.Number, Err.Description, True
End Function


Public Function funcClose(para操作名 As String) As Boolean
    On Error GoTo errHandler
    Dim lfrmForm As Form '待关闭的窗体
    funcClose = False
    '根据操作名设置不同的待关闭窗体。
'    Select Case para操作名
'    Case "职业病体检_五官科结果录入"
'        Set lfrmForm = frmHEENTnew_ResultInput
'    Case "职业病体检_X光影像科结果录入"
'        Set lfrmForm = frmXRay_ResultInput
'    Case "职业病体检_B超影像科结果录入"
'        Set lfrmForm = frmBUS_ResultInput
'    Case "职业病体检_内科结果录入"
'        Set lfrmForm = FrmInMedi_ResultInput
'    Case "职业病体检_外科结果录入"
'        Set lfrmForm = frmSurgery_ResultInput
'    Case "职业病体检_肝功能化验科结果录入"
'        Set lfrmForm = frmLiverFunc_ResultInput
'    Case "职业病体检_心电科结果录入"
'        Set lfrmForm = frmECG_ResultInput
'    Case "职业病体检_肺功能影像科结果录入"
'        Set lfrmForm = frmPFT_ResultInput
'    Case "职业病体检_尿常规化验科结果录入"
'        Set lfrmForm = frmURT_ResultInput
'    Case "职业病体检_染色体化验科结果录入"
'        Set lfrmForm = frmChromosome_ResultInput
'    Case "职业病体检_电测听科结果录入"
'        Set lfrmForm = frmElectroaudiometer_ResultInput
'    Case "职业病体检_血常规化验科结果录入"
'        Set lfrmForm = frmBloodRoutine_ResultInput
'    Case Else
'        funcClose = True
'    End Select
    Select Case para操作名
    Case "职业病体检_五官科结果录入"
        InputFlag = "五官科"
        InputFlagNo = "01"
'        frmResultInput_Routine.Caption = "五官科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_X光影像科结果录入"
        InputFlag = "X光影像科"
        InputFlagNo = "09"
'        frmResultInput_Assay.Caption = "X光影像科结果录入"
        Set lfrmForm = frmResultInput_Assay
        
    Case "职业病体检_B超影像科结果录入"
        InputFlag = "B超影像科"
        InputFlagNo = "11"
'        frmResultInput_Assay.Caption = "B超影像科结果录入"
        Set lfrmForm = frmResultInput_Assay
        
    Case "职业病体检_内科结果录入"
        InputFlag = "内科"
        InputFlagNo = "02"
'        frmResultInput_Routine.Caption = "内科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_外科结果录入"
        InputFlag = "外科"
        InputFlagNo = "03"
'        frmResultInput_Routine.Caption = "外科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_免疫科结果录入"
        InputFlag = "免疫科"
        InputFlagNo = "05"
'        frmResultInput_Routine.Caption = "免疫科结果录入"
        Set lfrmForm = frmResultInput_Routine
    
    Case "职业病体检_生化科结果录入"
        InputFlag = "生化科"
        InputFlagNo = "17"
'        frmResultInput_Routine.Caption = "生化科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_心电科结果录入"
        InputFlag = "心电科"
        InputFlagNo = "10"
'        frmResultInput_Routine.Caption = "心电科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_肺功能影像科结果录入"
        InputFlag = "肺功能影像科"
        InputFlagNo = "12"
'        frmResultInput_Routine.Caption = "肺功能影像科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_尿常规化验科结果录入"
        InputFlag = "尿常规化验科"
        InputFlagNo = "06"
'        frmResultInput_Routine.Caption = "尿常规化验科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_染色体化验科结果录入"
        InputFlag = "染色体化验科"
        InputFlagNo = "07"
'        frmResultInput_Routine.Caption = "染色体化验科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_电测听科结果录入"
        InputFlag = "电测听科"
        InputFlagNo = "08"
'        frmResultInput_Routine.Caption = "电测听科结果录入"
        Set lfrmForm = frmResultInput_Routine
        
    Case "职业病体检_血常规化验科结果录入"
        InputFlag = "血常规化验科"
        InputFlagNo = "04"
'        frmResultInput_Routine.Caption = "血常规化验科结果录入"
        Set lfrmForm = frmResultInput_Routine
    Case Else
        funcClose = -1
        Err.Raise 6666, , "当前体检管理版本不支持操作“" & para操作名 & "”。"
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
    sfsub错误处理 "职业病界面部件", "clsManageTestForm", "funcClose", Err.Number, Err.Description, True
End Function


Private Function sub保存单项体检结果(ByVal paraSysNo As String, _
                                     ByVal paraItem As String, _
                                     ByVal paraResult As String, _
                                     ByVal paraDoctor As String, _
                                     ByVal paraFillDate As Date, _
                                     ByVal paraConclusion As String, _
                                     ByVal paraTableName As String)
    Dim lstrSql As String
    Dim lobjRec As Object
    Dim lstrItemID As String
    Dim sql As String
    'sql = "exec autoConclusion '" & paraSysNo & "','" & paraItem & "','" & paraResult & "','" & paraDoctor & "','" & paraFillDate & "','" & paraConclusion & "','" & paraTableName & "'"
    'dafuncGetData (sql)
    
    '如果paraItem为文字，查找对应paraItem体检项目的项目编号
    If IsNumeric(paraItem) = False Then
        lstrSql = "select * from 职业病体检_体检项目设置表 where 名称='" & paraItem & "'"
        Set lobjRec = dafuncGetData(lstrSql)
        lstrItemID = lobjRec("编码")
    Else
        lstrItemID = paraItem
    End If
    
    '检查该系统编号和体检项目是否存在?
    '如果没有记录，则insert一条结果记录；否则，update该记录。
    lstrSql = "select * from " & paraTableName & " where 系统编号='" & paraSysNo & "' and 体检项目='" & lstrItemID & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    If lobjRec.RecordCount = 0 Then
        lstrSql = "insert into " & paraTableName & " values('" & paraSysNo & "','" & lstrItemID & "','" & paraResult & "','" & paraDoctor & "',getdate(),'" & paraConclusion & "')"
        dafuncGetData (lstrSql)
    Else
        lstrSql = "update " & paraTableName & " set 体检结果='" & paraResult & "',体检医师='" & paraDoctor & "',填写时间=getdate(),单项结论='" & paraConclusion & "' where 系统编号='" & paraSysNo & "' and 体检项目='" & lstrItemID & "'"
        dafuncGetData (lstrSql)
    End If
    
End Function

'通用的保存体检结果的函数。
'-------暂时没有用特别的判断结果的函数，只是单纯的字符串比较。
'-------自动判断结果的如果搞好了，就可以更改这部分。
'适用的体检信息表有如下字段：系统编号，(项目编号,存入时自动填写)，体检项目，体检结果，体检医师，填写日期，单项结论。
Public Function func保存单人体检结果(ByVal paraSysNo As String, _
                                     ByVal paraDoctor As String, _
                                     ByVal paraFillDate As Date, _
                                     ByVal paraItem As Collection, _
                                     ByVal paraResult As Collection, _
                                     ByVal paraTableName As String) As Boolean
    On Error GoTo errHandler
    Dim i As Integer
    Dim paraConclusion As String
    Dim sql As String
    For i = 1 To paraItem.Count
        If paraResult(i) = "正常" Then
            paraConclusion = "合格"
        Else
            paraConclusion = "不合格"
        End If
        'sql = "exec autoConclusion '" & paraSysNo & "','" & paraItem(i) & "','" & paraResult(i) & "','" & paraDoctor & "','" & paraFillDate & "','" & paraConclusion & "','" & paraTableName & "'"
        'dafuncGetData (sql)
        sub保存单项体检结果 paraSysNo, paraItem(i), paraResult(i), paraDoctor, paraFillDate, paraConclusion, paraTableName
    Next

    func保存单人体检结果 = True
    
    Exit Function
errHandler:
    MsgBox ("保存失败，事务回滚。")
    func保存单人体检结果 = False
End Function

'从 职业病体检_体检人员基本信息表 中获取某系统编号的体检人员的所有基本信息
Public Function func获取体检人员基本信息(ByVal paraSysNo As String) As Object
    Set func获取体检人员基本信息 = dafuncGetData("select * from 职业病体检_体检基本数据库 where 系统编号='" & paraSysNo & "'")
End Function

'从 职业病体检_结果信息_五官科 中获取某系统编号的体检人员的所有基本信息
Public Function func获取体检人员体检科室信息(ByVal paraSysNo As String, ByVal para科室 As String) As Boolean
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select * from 职业病体检_结果信息_" & para科室 & " where 系统编号='" & paraSysNo & "'")
    
    If lobjRec.RecordCount > 0 Then
        func获取体检人员体检科室信息 = True
    Else
        func获取体检人员体检科室信息 = False
    End If
    
End Function

'应该需要获取到 人员体检项目的相关信息。
'只查找出，当前科室的人员信息。
'什么情况下可以“修改结论”呢？暂时等待结论部分的完成。
Public Function func获取可修改结论的_特定科室的_体检人员基本信息(ByVal paraQuerySQL As String, ByVal paraDept As String) As Object
    Dim lstrSql As String
'    If paraDept = "五官科" Then
'        lstrSql = ""        '应该添加五官科体检类别的编号。最后查询的sql里面，暂时没有添加这个。
'    End If
    If paraDept = "" Then
        Set func获取可修改结论的_特定科室的_体检人员基本信息 = dafuncGetData("select * from 职业病体检_体检基本数据库 where 1=1" & paraQuerySQL)
    Else
        Set func获取可修改结论的_特定科室的_体检人员基本信息 = dafuncGetData("select * from 职业病体检_体检基本数据库 where 1=1" & paraQuerySQL & " and 系统编号 in (select 系统编号 from 职业病体检_结果信息_" & paraDept & " group by 系统编号)")
    End If
End Function

Public Function func获取可修改结论_本科室_体检人员信息(ByVal paraQuerySQL As String, ByVal paraDeptName As String) As Object
    Dim lstrSql As String
    
     '2012-07-14 于登淼 ↓
    '增加查找科室信息。
    Dim DeptInfo As Object  '记录科室信息，便于使用科室名称或科室编号时，方便查找。
    lstrSql = "select * from 系统管理_字典_字典内容表 where 描述='职业病体检_科室'"
    dasubSetQueryTimeout 6000
    Set DeptInfo = dafuncGetData(lstrSql)
    
    '修改查询条件，加入判断各科室8/48小时时间判断。
    Dim paraDeptNo As String
    DeptInfo.Filter = "名称='" & paraDeptName & "'"
    paraDeptNo = DeptInfo("编号")
    
    
'    lstrSql = "select distinct a.系统编号,a.姓名,a.性别,a.年龄,a.体检类型,a.单位名称,convert(varchar(10),b.填写时间,2) 填写时间 from 职业病体检_体检基本数据库 a, 职业病体检_结果信息_" & paraDeptName
'    lstrSql = lstrSql & " b where 1=1 and a.系统编号=b.系统编号" & paraQuerySQL & " and (a.体检状态='2' or a.体检状态='3' or a.体检状态='4')"
'    lstrSql = lstrSql & " and (substring(a.各科体检状态," & CInt(paraDeptNo) & ",1)='1' or substring(a.各科体检状态," & CInt(paraDeptNo) & ",1)='2')"
    '2012-07-14 于登淼 ↑
'    dasubSetQueryTimeout 600
'    Set func获取可修改结论_本科室_体检人员信息 = dafuncGetData(lstrSql)
     
    '2012年10月15日  罗李奎↓
     Set func获取可修改结论_本科室_体检人员信息 = dafuncGetData("exec 职业病体检_结果信息 '" & paraDeptName & "','" & paraDeptNo & "','" & paraQuerySQL & "'")
    '2012年10月15日  罗李奎↑
End Function

Public Function func是否已经体检过(ByVal paraSysNo As String, ByVal paraDept As String, Optional ByVal paraFillDate As Date) As Object
    Set func是否已经体检过 = dafuncGetData("select b.*,a.名称 项目名称 from 职业病体检_体检项目设置表 a,职业病体检_结果信息_" & paraDept & " b where a.编码=b.体检项目 and b.系统编号='" & paraSysNo & "'") ' and 填写时间='" & paraFillDate & "'")
End Function

Public Function func获取体检项目编号(ByVal paraItemName As String) As String
    Dim lobjTmp As Object
    Set lobjTmp = dafuncGetData("select 编码 from 职业病体检_体检项目设置表 where 名称='" & paraItemName & "'")
    func获取体检项目编号 = lobjTmp("编码")
End Function

'为减少复杂度，科室暂时为编号，在代码中，如果有必要更改为科室名称，另外写代码实现
'所有科室结论
Public Function func获取体检人员科室结论(ByVal paraSysNo As String) As Object
    Dim lstrSql As String
    lstrSql = "select a.科室,a.文字结论,b.姓名 医师姓名 from 职业病体检_科室结论表 a, 系统管理_员工基本信息表 b  where a.医生编号=b.编号 and a.系统编号='" & paraSysNo & "'"
    Set func获取体检人员科室结论 = dafuncGetData(lstrSql)
End Function

Public Function func获取体检人员单科室体检结果(ByVal paraSysNo As String, ByVal paraDeptName As String) As Object
    On Error Resume Next
    Dim lstrSql As String
    lstrSql = "select a.体检项目 项目编号,b.名称 体检项目,a.体检结果,c.姓名,a.单项结论 from 职业病体检_结果信息_" & paraDeptName & " a, 职业病体检_体检项目设置表 b, 系统管理_员工基本信息表 c where a.系统编号='" & paraSysNo & "' and a.体检项目=b.编码 and c.编号=a.体检医师"
    Set func获取体检人员单科室体检结果 = dafuncGetData(lstrSql)
End Function

Public Function func获取体检人员照片(ByVal paraSysNo As String) As Object
    Set func获取体检人员照片 = dafuncGetData("select * from 系统管理_系统图片管理表 where 图片编号='" & paraSysNo & "'")
End Function

Public Function func查找结果图片(ByVal paraSysNo As String, ByVal paraItemNo As String) As Object
    Set func查找结果图片 = dafuncGetData("select * from 职业病体检_体检人员图片结果表 where 系统编号='" & paraSysNo & "' and 项目编号='" & paraItemNo & "'")
End Function

Public Function func保存结果图片(ByVal ParaPicture As StdPicture, ByVal paraID As String, ByVal paraItemNo As String, ByVal paraDate As Date)
    On Error GoTo errHandler
    Dim lstrSql As String              'SQL语句
    Dim lrecPicture As Object          '根据语句返回图片信息的RecordSet
    Dim lprbPicture As New PropertyBag '将图片信息进行序列化的属性包
    Dim lobjRec  As Object
    
    '将图片写入属性包进行序列化。
    lprbPicture.WriteProperty "Picture", ParaPicture
    
    '根据标识号取出相应的图片。
    lstrSql = "select * from 职业病体检_体检人员图片结果表 where 系统编号='" & paraID & "' and 项目编号='" & paraItemNo & "'"
    Set lrecPicture = dafuncGetData(lstrSql)
    
    '如果返回空记录集，则新增一条记录。
    If lrecPicture.RecordCount = 0 Then
        lrecPicture.AddNew
    End If

    lstrSql = "select getdate()"
    Set lobjRec = dafuncGetData(lstrSql)
    
    '将图片信息写入记录集中。
    lrecPicture("图片").AppendChunk lprbPicture.Contents
    lrecPicture("系统编号") = paraID
    lrecPicture("项目编号") = paraItemNo
    lrecPicture("填写时间") = lobjRec(0)
    
    '保存记录集更新。
    lrecPicture.Update
    
errHandler:
    Set lrecPicture = Nothing
    Set lprbPicture = Nothing
    Set ParaPicture = Nothing
    If Err.Number = 0 Then Exit Function
    Err.Raise Err.Number, , Err.Description
End Function


Public Function func获取结果图片(ByVal paraID As String, ByVal paraItemNo As String, ByVal paraPicturePath As String) As StdPicture
    On Error GoTo errHandler
    Dim lstrSql As String              'SQL语句
    Dim lrecPicture As Object          '根据语句返回图片信息的RecordSet
    Dim lprbPicture As New PropertyBag '将图片信息进行序列化的属性包
    Dim lvarPicture() As Byte          '图片转换成的Byte数组
    
    '根据标识号返回相应图片。
    lstrSql = "select * from 职业病体检_体检人员图片结果表 where 系统编号='" & paraID & "' and 项目编号='" & paraItemNo & "'"
    Set lrecPicture = dafuncGetData(lstrSql)
    
    '从返回的记录集中取出图片信息。
    If lrecPicture.RecordCount = 0 Then
        Set func获取结果图片 = LoadPicture(App.Path & "\" & paraPicturePath)
        Exit Function
    End If
    lvarPicture = lrecPicture("图片").GetChunk(lrecPicture("图片").ActualSize)
    
    '将图片信息由属性包中读出。
    lprbPicture.Contents = lvarPicture
    
    '返回根据唯一标识号所取得的图片。
    Set func获取结果图片 = lprbPicture.ReadProperty("Picture")
    
errHandler:
    Set lrecPicture = Nothing
    Set lprbPicture = Nothing
    If Err.Number = 0 Then Exit Function
End Function

'删除已有的体检结果图片
Public Function func删除结果图片(paraID As String, paraItemNo As String)
    On Error GoTo errHandler
    Dim lstrSql As String 'SQL语句
    
    '删除该标识号所对应的图片。
    lstrSql = "delete from 职业病体检_体检人员图片结果表 where 系统编号='" & paraID & "' and 项目编号='" & paraItemNo & "'"
    dafuncGetData lstrSql
    
    Exit Function
errHandler:
    sfsub错误处理 "职业病体检结果录入", "clscommon", "func删除结果图片", Err.Number, Err.Description, True
End Function


