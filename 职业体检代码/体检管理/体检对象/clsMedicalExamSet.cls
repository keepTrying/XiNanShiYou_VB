VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedicalExamSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'本对象描述满足指定过滤条件的体检对象集合。
'过滤条件由系统编号、体检表名称、体检状态、体检日期范围、健康档案编号、单位名称等子条件组合成，
'各子条件之间逻辑关系是"AND"。
'若某属性为空，表示忽略该过滤子条件。

Public 从系统编号 As String
Public 到系统编号 As String
Public 从试管编号 As String
Public 到试管编号 As String
Public 从体检日期 As String
Public 到体检日期 As String
Public 体检状态 As String     '体检的当前状态（枚举范围，比如如："1，2"
Public 下结论医师 As String   '医师编号。
Public 健康档案编号 As String
Public 单位名称 As String     '模糊查询。
Public 姓名 As String
Public 性别 As String
Public 年龄 As Long
Public 身份证号 As String
Public 体检表名称 As String
Public 复查标志 As Long       '是否需要复查的标志：0不需要复查/1需要/-1不作复查检查。

Public 复查系统编号 As String '为空，表示还未复查登记；不为空，表示已复查过

'功能：满足过滤条件的体检对象的集合。
'返回：collection[系统编号，健康档案编号，姓名，性别，年龄，单位申请编号，单位名称，体检日期，体检结论，诊断和处理意见，体检状态]。
'作者：邓恒
'修改：2001-12-30（增加参数“paraOrderBy”）。
Public Property Get 元素集(Optional paraFields As String = "", Optional paraOrderby As String = "") As Object
    Dim lstrSql As String   '要执行的sql语句。
    Dim lobjRec As Object   '执行sql语句的结果记录。

    On Error GoTo errHandler
    '拼sql语句。
    If 从系统编号 <> "" Then
        lstrSql = "系统编号>='" & 从系统编号 & "'"
    End If
    If 到系统编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "系统编号<='" & 到系统编号 & "'"
    End If
    If 从试管编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "试管编号>='" & 从试管编号 & "'"
    End If
    If 到试管编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "试管编号<='" & 到试管编号 & "'"
    End If
    If 从体检日期 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检日期>='" & 从体检日期 & "'"
    End If
    If 到体检日期 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检日期<='" & 到体检日期 & "'"
    End If
    If 体检状态 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检状态 in (" & 体检状态 & ")"
    End If
    If 下结论医师 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "下结论医师 ='" & 下结论医师 & "'"
    End If
    If 体检表名称 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检表名称 ='" & 体检表名称 & "'"
    End If
    
    If 健康档案编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "健康档案编号 ='" & 健康档案编号 & "'"
    End If
    If 单位名称 <> "" Then '模糊查询。
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "单位名称 like '" & 单位名称 & "%'"
    End If
    If 姓名 <> "" Then '模糊查询。
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "姓名 like '" & 姓名 & "%'"
    End If
    If 性别 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "性别 ='" & 性别 & "'"
    End If
    If 年龄 <> 0 Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "datediff(year,出生日期,getdate()) =" & 年龄
    End If
    If 身份证号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "公民身份号码 ='" & 身份证号 & "'"
    End If
    If 复查标志 <> -1 Then
        If 复查标志 = 0 Then
            '不需要复查。
            lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "isnull(复查体检表名,'')=''"
        Else
            '需要复查。
            lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "isnull(复查体检表名,'')<>''"
        End If
        
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "isnull(复查系统编号,'')='" & 复查系统编号 & "'"
    End If
    
    dasubSetQueryTimeout 30
    'Set 元素集 = dafuncGetData("select " & IIf(paraFields = "", "*", paraFields) & " from 体检管理_体检基本数据库 " & IIf(lstrSql = "", "", " where " & lstrSql) & IIf(InStr(paraFields, "distinct") = 0, " order by " & IIf(paraOrderby = "", "系统编号", paraOrderby), ""))
    lstrSql = Replace(lstrSql, "'", "''")
    Set 元素集 = dafuncGetData("exec 体检管理_获取可修改的体检记录 '" & um用户编号 & "', '" & lstrSql & "'")
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExamSet", "Property Get 元素集", Err.Number, Err.Description, True
    Exit Property
    Resume
End Property

Public Property Get 元素集old(Optional paraFields As String = "", Optional paraOrderby As String = "") As Object
    Dim lstrSql As String   '要执行的sql语句。
    Dim lobjRec As Object   '执行sql语句的结果记录。

    On Error GoTo errHandler
    '拼sql语句。
    If 从系统编号 <> "" Then
        lstrSql = "系统编号>='" & 从系统编号 & "'"
    End If
    If 到系统编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "系统编号<='" & 到系统编号 & "'"
    End If
    If 从试管编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "试管编号>='" & 从试管编号 & "'"
    End If
    If 到试管编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "试管编号<='" & 到试管编号 & "'"
    End If
    If 从体检日期 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检日期>='" & 从体检日期 & "'"
    End If
    If 到体检日期 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检日期<='" & 到体检日期 & "'"
    End If
    If 体检状态 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检状态 in (" & 体检状态 & ")"
    End If
    If 下结论医师 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "下结论医师 ='" & 下结论医师 & "'"
    End If
    If 体检表名称 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "体检表名称 ='" & 体检表名称 & "'"
    End If
    
    If 健康档案编号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "健康档案编号 ='" & 健康档案编号 & "'"
    End If
    If 单位名称 <> "" Then '模糊查询。
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "单位名称 like '" & 单位名称 & "%'"
    End If
    If 姓名 <> "" Then '模糊查询。
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "姓名 like '" & 姓名 & "%'"
    End If
    If 性别 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "性别 ='" & 性别 & "'"
    End If
    If 年龄 <> 0 Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "datediff(year,出生日期,getdate()) =" & 年龄
    End If
    If 身份证号 <> "" Then
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "公民身份号码 ='" & 身份证号 & "'"
    End If
    If 复查标志 <> -1 Then
        If 复查标志 = 0 Then
            '不需要复查。
            lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "isnull(复查体检表名,'')=''"
        Else
            '需要复查。
            lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "isnull(复查体检表名,'')<>''"
        End If
        
        lstrSql = lstrSql & IIf(lstrSql = "", "", " and ") & "isnull(复查系统编号,'')='" & 复查系统编号 & "'"
    End If
    
    dasubSetQueryTimeout 30
    Set 元素集old = dafuncGetData("select " & IIf(paraFields = "", "*", paraFields) & " from 体检管理_体检基本数据库 " & IIf(lstrSql = "", "", " where " & lstrSql) & IIf(InStr(paraFields, "distinct") = 0, " order by " & IIf(paraOrderby = "", "系统编号", paraOrderby), ""))
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExamSet", "Property Get 元素集old", Err.Number, Err.Description, True
    Exit Property
    Resume
End Property
Private Sub Class_Initialize()
    复查标志 = -1
End Sub

'功能：清空本对象属性。
'作者：邓恒
Public Sub subClear()
    
    从系统编号 = ""
    到系统编号 = ""
    从试管编号 = ""
    到试管编号 = ""
    从体检日期 = ""
    到体检日期 = ""
    体检状态 = ""
    下结论医师 = ""
    健康档案编号 = ""
    单位名称 = ""
    姓名 = ""
    性别 = ""
    年龄 = 0
    身份证号 = ""
    体检表名称 = ""

    复查标志 = -1
    复查系统编号 = ""
End Sub
