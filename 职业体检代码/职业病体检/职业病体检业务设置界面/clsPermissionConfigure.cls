VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPermissionConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'2012-03-16 于登淼
'添加权限设置类的全部功能

Option Explicit

'获取单个医师可操作的所有科室。
Public Function func获取职业病体检单个医师科室(ByVal paraDoctorNo As String) As Object
    dasubSetQueryTimeout 6000
    Set func获取职业病体检单个医师科室 = dafuncGetData("select * from 职业病体检_用户科室权限表 where 用户编号='" & paraDoctorNo & "' order by 科室编号")
End Function

'获取单个医师所有操作权限(仅职业病体检模块)
Public Function func获取职业病体检单个医师操作权限(ByVal paraDoctorNo As String) As Object
    dasubSetQueryTimeout 6000
    Set func获取职业病体检单个医师操作权限 = dafuncGetData("select * from 职业病体检_用户操作权限表 where 用户编号='" & paraDoctorNo & "'")
End Function

'获取单个科室下的所有操作权限(paraDeptName: 科室名称，如“五官科”等，不是编号)
Public Function func获取职业病体检单个科室所有操作权限(ByVal paraDeptName As String) As Object
    Dim lstrSql As String
    lstrSql = "select * from 职业病体检_可用操作信息表 where 操作名 like '%" & paraDeptName & "%' order by 上级操作名"
    dasubSetQueryTimeout 6000
    Set func获取职业病体检单个科室所有操作权限 = dafuncGetData(lstrSql)
End Function

'增加或删除某个医师可以操作的单个科室(paraKey: 科室编号，如“01”，代表“五官科”)
Public Function func修改职业病体检医师单个科室(ByVal paraDoctorNo As String, ByVal paraCheck As Boolean, ByVal paraKey As String)
    Dim lobjRec As Object
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select * from 职业病体检_用户科室权限表 where 用户编号='" & paraDoctorNo & "' and 科室编号='" & paraKey & "'")
    If paraCheck = True Then
        If lobjRec.recordcount = 0 Then dafuncGetData ("insert into 职业病体检_用户科室权限表 values('" & paraDoctorNo & "','" & paraKey & "')")
    Else
        dafuncGetData ("delete 职业病体检_用户科室权限表 where 用户编号='" & paraDoctorNo & "' and 科室编号='" & paraKey & "'")
    End If
End Function

'增加或删除某个医师单个科室的系统权限(仅限于窗体显示或不显示)
Public Function func修改医师单个科室系统权限(ByVal paraDoctorNo As String, ByVal paraCheck As Boolean, ByVal paraPermission As String)
    Dim lobjRec As Object
    Dim lobjTmp As Object
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select * from 系统管理_用户操作权限表 where 用户编号='" & paraDoctorNo & "' and 权限名='" & paraPermission & "'")
    If paraCheck = True Then
        If lobjRec.recordcount = 0 Then dafuncGetData ("insert into 系统管理_用户操作权限表 values('" & paraDoctorNo & "','" & paraPermission & "')")
    Else
        dafuncGetData ("delete 系统管理_用户操作权限表 where 用户编号='" & paraDoctorNo & "' and 权限名='" & paraPermission & "'")
    End If
    
    Set lobjTmp = CreateObject("系统管理.cls用户")
    lobjTmp.员工编号 = paraDoctorNo
    lobjTmp.sus自动设置平台 1
End Function

'增加或删除某个医师的单个可用操作(paraKey: 权限名，如“职业病体检_五官科结果录入”等，都为全称)
Public Function func修改职业病体检医师单个可用操作(ByVal paraDoctorNo As String, ByVal paraCheck As Boolean, ByVal paraKey As String)
    Dim lobjRec As Object
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select * from 职业病体检_用户操作权限表 where 用户编号='" & paraDoctorNo & "' and 权限名='" & paraKey & "'")
    If paraCheck = True Then
        If lobjRec.recordcount = 0 Then dafuncGetData ("insert into 职业病体检_用户操作权限表 values('" & paraDoctorNo & "','" & paraKey & "')")
    Else
        dafuncGetData ("delete 职业病体检_用户操作权限表 where 用户编号='" & paraDoctorNo & "' and 权限名='" & paraKey & "'")
    End If
End Function

'--------以下为对外接口部分--------
'传入参数为：医师编号，科室名称(如“五官科”)。科室名称必须与“系统管理_字典_字典内容表”中的“职业病”科室名称一致
Public Function func医师科室权限(ByVal paraDoctorNo As String, ByVal paraDeptName) As Boolean
    Dim lobjRec As Object
    Dim paradeptno As String
    dasubSetQueryTimeout 6000
    Set lobjRec = pobjDict.Fetch("职业病体检科室字典")
    lobjRec.Filter = "名称='" & paraDeptName & "'"
    If lobjRec.recordcount > 0 Then
        paradeptno = lobjRec("编号")
        Set lobjRec = dafuncGetData("select * from 职业病体检_用户科室权限表 where 用户编号='" & paraDoctorNo & "' and 科室编号='" & paradeptno & "'")
        If lobjRec.recordcount > 0 Then func医师科室权限 = True: Exit Function
    End If
    func医师科室权限 = False
End Function

'科室操作权限名称需要写“职业病体检_可用操作信息表”，操作名字段的全称
Public Function func科室操作权限(ByVal paraDoctorNo As String, ByVal paraPermission As String) As Boolean
    Dim lobjRec As Object
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select * from 职业病体检_用户操作权限表 where 用户编号='" & paraDoctorNo & "' and 权限名='" & paraPermission & "'")
    If lobjRec.recordcount > 0 Then
        func科室操作权限 = True
    Else
        func科室操作权限 = False
    End If
End Function

'2012-07-02 于登淼
Public Function func获取职业病体检科室医师基本信息() As Object
    Dim strSQL As String
    dasubSetQueryTimeout 6000
    strSQL = "select * from 系统管理_员工基本信息表" '-- where 科室='06'"
    Set func获取职业病体检科室医师基本信息 = dafuncGetData(strSQL)
End Function


'修改人：张令 2012.11.29
'说明：按照病人编号查询，传入病人编号    ↓
'bug号：0000041
Public Function func获取职业病体检人员基本信息(ByVal paraWhere As String) As Object
    Dim strSQL As String
    dasubSetQueryTimeout 6000
'    strSQL = "select a.系统编号,b.姓名 from 职业病体检_体检基本信息表 a, 职业病体检_体检人员基本信息表 b where (a.体检状态='2' or a.体检状态='3' or a.体检状态='4') and (charindex('3', a.各科体检状态)>0 or left(a.各科体检状态,1)='3') and a.系统编号=b.系统编号 and a.系统编号= " & paraWhere & " order by a.系统编号"
'    If paraWhere = "" Then Exit Function
    strSQL = "select a.系统编号,b.姓名 from 职业病体检_体检基本信息表 a, 职业病体检_体检人员基本信息表 b " _
                & "where (a.体检状态='2' or a.体检状态='3' or a.体检状态='4') and (left(a.各科体检状态,1)='3') " _
                & "and a.系统编号=b.系统编号 and a.系统编号='" & paraWhere & "' order by a.系统编号"
    Set func获取职业病体检人员基本信息 = dafuncGetData(strSQL)
End Function

'2012-07-02 于登淼
'修改单个体检人员单个科室的体检结果录入状态，同时更改当前体检状态
Public Function func单人单科室结果重新修改(ByVal paraSysNo As String, ByVal paraDeptIndex As Integer, ByVal paraCurState As String)
    Dim lobjTmp As Object
    Dim strSQL As String
    dasubSetQueryTimeout 6000
    Set lobjTmp = CreateObject("职业病对象.clsManageMedicalExam")
    lobjTmp.func修改单人各科体检状态 paraSysNo, paraDeptIndex, paraCurState
    lobjTmp.func写入单人当前体检状态 paraSysNo, "03" '03为“体检中”
    lobjTmp.sub修改起始时间 paraSysNo, "", Format(CStr(paraDeptIndex), "#00")
End Function
