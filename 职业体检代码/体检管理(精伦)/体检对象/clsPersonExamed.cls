VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPersonExamed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'属性变量。
Private mstrHealthFileNo As String          '健康档案编号
Private mpicPhoto As StdPicture             '照片
Private mstr公民身份号码 As String

Public 姓名 As String
Public 性别 As String
Public 出生日期 As String
Private mstr年龄 As String

Private mstr单位申请编号 As String
Public 单位名称 As String
Public 建档日期 As String

Public 卫生种类 As String
Public 片区 As String
Public 行业类别 As String

'功能：根据出生日期获取年龄。
'作者：杨春
Public Property Get 年龄() As String
    On Error Resume Next
    年龄 = mstr年龄
End Property
Public Property Let 年龄(ByVal vNewValue As String)
    mstr年龄 = vNewValue
End Property

'功能：获取健康档案编号。
Public Property Get 健康档案编号() As String
    健康档案编号 = mstrHealthFileNo
End Property


'功能：设置健康档案编号，从而获取本对象其他属性。
'作者：邓恒
Public Property Let 健康档案编号(ByVal paraHealthFileNo As String)
    On Error GoTo errHandler
    Dim lobjRec As Object   '根据健康档案获取得体检人员基本信息记录。
    
    mstrHealthFileNo = paraHealthFileNo
    mstr公民身份号码 = ""
    姓名 = ""
    出生日期 = ""
    mstr单位申请编号 = ""
    单位名称 = ""
    建档日期 = ""
    年龄 = ""
    
    '只有当引用像片时，才获取它。
    Set mpicPhoto = Nothing
    
    '健康档案编号不为空时将查询到的值赋值给相关属性
    If mstrHealthFileNo <> "" Then
        Set lobjRec = dafuncGetData("select * from 体检管理_体检人员基本信息表 where 健康档案编号='" & mstrHealthFileNo & "'")
        If lobjRec.recordcount > 0 Then
            mstr公民身份号码 = IIf(IsNull(lobjRec("公民身份号码")), "", lobjRec("公民身份号码"))
            姓名 = IIf(IsNull(lobjRec("姓名")), "", lobjRec("姓名"))
            性别 = IIf(IsNull(lobjRec("性别")), "", lobjRec("性别"))
            出生日期 = IIf(IsNull(lobjRec("出生日期")), "", lobjRec("出生日期"))
            卫生种类 = IIf(IsNull(lobjRec("卫生种类")), "", lobjRec("卫生种类"))
            片区 = IIf(IsNull(lobjRec("片区")), "", lobjRec("片区"))
            行业类别 = IIf(IsNull(lobjRec("行业类别")), "", lobjRec("行业类别"))
            
            单位名称 = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
            建档日期 = IIf(IsNull(lobjRec("建档日期")), "", lobjRec("建档日期"))
            年龄 = IIf(IsNull(lobjRec("年龄")), "", lobjRec("年龄"))
            
            单位申请编号 = IIf(IsNull(lobjRec("单位申请编号")), "", lobjRec("单位申请编号"))
        End If
        lobjRec.Close
    End If
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "clsPersonExamed", "Property Let 健康档案编号", Err.Number, Err.Description, True
End Property
Public Sub sub清空系统编号()
    mstrHealthFileNo = ""
End Sub
'功能：获取公民身份号码。
'作者：邓恒
Public Property Get 公民身份号码() As String
    公民身份号码 = mstr公民身份号码
End Property
'功能：设置公民身份号码，从而自动分解出出生日期和性别。
'作者：邓恒
Public Property Let 公民身份号码(ByVal vNewValue As String)
    Dim lstrDate As String
    Dim lstrSex As String
    On Error GoTo errHandler
    mstr公民身份号码 = vNewValue
    
    '根据公民身份号码获取生日和性别。
    If vNewValue <> "" Then
        sub根据公民身份号码获取生日和性别 mstr公民身份号码, lstrDate, lstrSex
        If IsDate(lstrDate) Then
            出生日期 = lstrDate
            性别 = lstrSex
        End If
    End If
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "clsPersonExamed", "Property Let 公民身份号码", Err.Number, Err.Description, True
End Property

'作者：邓恒
Public Property Get 像片() As StdPicture
    On Error Resume Next
    '获取照片
    If mpicPhoto Is Nothing Then
        Set mpicPhoto = pmfunc获取图片(mstrHealthFileNo, P_SUBSYSNAME)
    End If
    
    Set 像片 = mpicPhoto
End Property

Public Property Let 像片(ByVal vNewValue As StdPicture)
    Set mpicPhoto = vNewValue
End Property
Public Property Let 像片压缩(ByVal vNewValue As StdPicture)
    Dim lobjTransfer As Object
    Dim lstrTmp As String
    
    Set lobjTransfer = CreateObject("BmpToJGP.clsBmpToJPG")
    lstrTmp = App.Path & "\tmp.bmp"
    RidFile lstrTmp
    SavePicture vNewValue, lstrTmp
    lobjTransfer.subSetBMPFile lstrTmp
    Set mpicPhoto = lobjTransfer.funcSaveToJPG()
End Property
Public Function FileExists(FileName As String) As Boolean
    If Len(FileName) > 0 Then FileExists = (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function
Public Sub RidFile(FileName As String)
    On Error Resume Next
    If FileExists(FileName) Then
        SetAttr FileName, vbNormal
        Kill FileName
    End If
End Sub
'功能：分配健康档案编号。
'返回：分配的健康档案编号。
'作者：邓恒
Public Function Func分配健康档案编号(ByVal para预留参数 As Collection) As String
Attribute Func分配健康档案编号.VB_Description = "'功能：分配健康档案编号。\r\n    '返回：分配的健康档案编号。\r\n"
    On Error GoTo errHandler
    Dim lobjRule As Object          'ClsSysNoRule
    Dim lstrHealthNo As String      '生成的健康档案编号。
    Dim lobjRec As Object           '执行“生成编号流水号”存储过程的返回记录。
    
    '通过系统编号规则生成编号（忽略错误）。
    On Error Resume Next
    Set lobjRule = CreateObject("系统编号规则.ClsSysNoRule")
    If Err <> 0 Then
        lobjRule.编号名称 = "健康档案编号"
        lstrHealthNo = lobjRule.Func生成系统编号(para预留参数)
    End If
    
    On Error GoTo errHandler
    '若生成失败，说明生成编号出错，使用缺省方法生成编号：防疫站编号3+服务器代号2+日期6+流水号4。
    If lstrHealthNo = "" Then
        '获取服务日期。
        Set lobjRec = dafuncGetData("select getdate()")
        '生成固定部分：防疫站编号3+服务器代号2+日期6。
        If um防疫站编号 = "" Then
            lstrHealthNo = "000" & "00" & Format(lobjRec(0), "yymmdd")
        Else
            lstrHealthNo = um防疫站编号 & um服务器代号 & Format(lobjRec(0), "yymmdd")
        End If
        
        '生成流水号。
        Set lobjRec = dafuncGetData("exec 体检管理_生成编号流水号 '健康档案编号'")
        
        lstrHealthNo = lstrHealthNo & Format(lobjRec(0), "0000")
    End If
    Func分配健康档案编号 = lstrHealthNo
    
    Exit Function
errHandler:
    sfsub错误处理 "体检对象部件", "clsPersonExamed", "Func分配健康档案编号", Err.Number, Err.Description, True
    Exit Function
    Resume
End Function

'功能：自动分配"健康档案编号"，保存体检人员的健康档案基本信息和像片。
'作者：邓恒
'修改：2001-10-30（处理主健重复错误）。
Public Sub Sub保存()
Attribute Sub保存.VB_Description = "'功能：自动分配""健康档案编号""，保存体检人员的健康档案基本信息和像片。"
    Dim lobjRec As Object
    Dim lcolInfo As Collection '生成健康档案编号所需要的预留参数。
    Dim lstrSql As String
    Dim lstrError As String
    Dim lblnInsert As Boolean
    
    On Error GoTo errHandler
    
    '修改：2002-1-14（判断该编号的记录是否存在）。
    lblnInsert = True
    If mstrHealthFileNo <> "" Then
        Set lobjRec = dafuncGetData("select 健康档案编号 from 体检管理_体检人员基本信息表 where 健康档案编号='" & mstrHealthFileNo & "'")
        If lobjRec.recordcount > 0 Then
            lblnInsert = False
        End If
    End If
    
    
    If lblnInsert Then
        '插入记录。
        
retry:  If mstrHealthFileNo = "" Then
            '生成健康档案编号。
            Set lcolInfo = New Collection
            On Error Resume Next
            mstrHealthFileNo = Func分配健康档案编号(lcolInfo)
        End If
        '新增保存到库中
        lstrSql = "insert into 体检管理_体检人员基本信息表(健康档案编号,公民身份号码,姓名,性别,年龄,出生日期,单位申请编号,单位名称,建档日期,卫生种类,片区,行业类别) values( '" _
                & mstrHealthFileNo & "','" & mstr公民身份号码 & "','" & 姓名 & "','" & 性别 & "','" & 年龄 & "'," & IIf(出生日期 = "", "null", "'" & 出生日期 & "'") & ",'" & mstr单位申请编号 & "','" & 单位名称 & "'," _
                & IIf(建档日期 = "", "getdate()", "'" & 建档日期 & "'") & "," & IIf(卫生种类 = "", "null", "'" & 卫生种类 & "'") & "," _
                & IIf(片区 = "", "null", "'" & 片区 & "'") & "," & IIf(行业类别 = "", "null", "'" & 行业类别 & "'") & ")"
        dafuncGetData lstrSql
        If Err.Number = -2147217873 Then
            '主健重复，重试。
            mstrHealthFileNo = ""
            Err.Clear
            GoTo retry
        ElseIf Err.Number <> 0 Then
            '上举错误。
            lstrError = func错误处理(Err.Number, Err.Description)
            On Error GoTo errHandler
            Err.Raise 6666, , lstrError
        End If
        On Error GoTo errHandler
    Else
        '修改保存到数据库中。
        lstrSql = "update 体检管理_体检人员基本信息表 set 公民身份号码='" & mstr公民身份号码 & "',姓名='" & 姓名 & "'" _
                & ",性别='" & 性别 & "',年龄='" & 年龄 & "',出生日期=" & IIf(出生日期 = "", "null", "'" & 出生日期 & "'") & ",单位申请编号='" & mstr单位申请编号 & "'" _
                & ",单位名称='" & 单位名称 & "',建档日期=" & IIf(建档日期 = "", "getdate()", "'" & 建档日期 & "'") _
                & ",卫生种类=" & IIf(卫生种类 = "", "null", "'" & 卫生种类 & "'") & ",片区=" & IIf(片区 = "", "null", "'" & 片区 & "'") & ",行业类别=" & IIf(行业类别 = "", "null", "'" & 行业类别 & "'") _
                & " where 健康档案编号='" & mstrHealthFileNo & "'"
        dafuncGetData lstrSql
    End If
    
    
    '若像片不为空，则保存到相应目录
    Dim lobjPhoto As StdPicture
    If Not mpicPhoto Is Nothing Then
        Set lobjPhoto = mpicPhoto
        pmsub保存图片 lobjPhoto, mstrHealthFileNo, P_SUBSYSNAME
    End If
    Exit Sub
errHandler:
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsPersonExamed", "Sub保存", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：删除健康档案。
'作者：邓恒
Public Sub Sub删除()
    On Error GoTo errHandler
    Dim lstrSql As String
    
    '删除库中的基本信息
    If mstrHealthFileNo <> "" Then
        lstrSql = "delete 体检管理_体检人员基本信息表 where 健康档案编号='" & mstrHealthFileNo & "'"
        dafuncGetData lstrSql
        
        '当有照片时删除照片
        On Error Resume Next
        pmsub删除图片 mstrHealthFileNo, P_SUBSYSNAME
        
    End If
    Set mpicPhoto = Nothing
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsPersonExamed", "Sub删除", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：获取本人最近一次体检
'作者：邓恒
Public Function Func获取本人最近一次体检() As clsMedicalExam
    Dim lobj最近体检 As clsMedicalExam
    Dim lstrSql As String
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    Set Func获取本人最近一次体检 = Nothing
    
    '获取最近一次体检的日期。
    lstrSql = "select max(体检日期) from 体检管理_体检基本信息表 where 健康档案编号='" & mstrHealthFileNo & "' and 体检类别<>" & P_EXAM_AGAIN
    Set lobjRec = dafuncGetData(lstrSql)
    If IIf(IsNull(lobjRec(0)), "", lobjRec(0)) <> "" Then
        '获取最近一次体检的系统编号。
        lstrSql = "select 系统编号 from 体检管理_体检基本信息表 where 健康档案编号='" & mstrHealthFileNo & "' and 体检日期='" & lobjRec(0) & "'"
        Set lobjRec = dafuncGetData(lstrSql)
        If lobjRec.recordcount > 0 Then
            Set lobj最近体检 = New clsMedicalExam
            lobj最近体检.系统编号 = lobjRec("系统编号")
            Set Func获取本人最近一次体检 = lobj最近体检
        End If
    End If
    lobjRec.Close
    
    Exit Function
errHandler:
    sfsub错误处理 "体检对象部件", "clsPersonExamed", "Func获取本人最近一次体检", Err.Number, Err.Description, True
End Function

Private Sub Class_Initialize()

    建档日期 = Format(Date, "yyyy-mm-dd")
    
End Sub
'修改：2001-8-23（把公用变量属性改为属性过程，以便获取卫生种类、片区、行业类别等属性）。
Public Property Get 单位申请编号() As String
    单位申请编号 = mstr单位申请编号
End Property
Public Property Let 单位申请编号(ByVal vNewValue As String)
    Dim lobjRec As Object
    Dim lobjDict  As Object
    On Error Resume Next
    
    mstr单位申请编号 = vNewValue
    '若没有给卫生种类等属性复制，根据单位申请编号获取这些属性。
    If mstr单位申请编号 <> "" And (卫生种类 = "" Or 片区 = "" Or 行业类别 = "") Then
        Set lobjRec = dafuncGetData("select * from 单位档案_单位定位查询视图 where 申请编号='" & mstr单位申请编号 & "'")
        If lobjRec.recordcount > 0 Then
            If 卫生种类 = "" Then
                卫生种类 = IIf(IsNull(lobjRec!卫生种类BH), "", lobjRec!卫生种类BH)
            End If
            If 片区 = "" Then
                片区 = IIf(IsNull(lobjRec!片区BH), "", lobjRec!片区BH)
            End If
            If 行业类别 = "" Then
                '创建字典对象。
                Set lobjDict = CreateObject("字典管理.clsDictionary")
                
                '获取行业类别编号。
                If IIf(IsNull(lobjRec("行业类别")), "", lobjRec("行业类别")) <> "" Then
                    Set lobjRec = lobjDict.Fetch("行业类别字典视图", "名称='" & lobjRec("行业类别") & "'")
                    If Not lobjRec Is Nothing Then
                        If lobjRec.recordcount > 0 Then
                            行业类别 = IIf(IsNull(lobjRec("编号")), "", lobjRec("编号"))
                        End If
                    End If
                End If
            End If
        
        End If
        
    End If
End Property

Public Property Get 卫生种类名称() As String
    On Error Resume Next
    Dim lstrTemp As String
    If 卫生种类 <> "" Then
        '根据编号获取名称.
        Dim lobjRec As Object
        Set lobjRec = dafuncGetData("select * from 系统管理_卫生种类字典视图 where 编号='" + 卫生种类 + "'")
        If lobjRec.recordcount > 0 Then
            lstrTemp = lobjRec!名称
        End If
    Else
    End If
    卫生种类名称 = lstrTemp
    
End Property

Public Property Get 片区名称() As String
    On Error Resume Next
    Dim lstrTemp As String
    If 片区 <> "" Then
        '根据编号获取名称.
        Dim lobjRec As Object
        Set lobjRec = dafuncGetData("select * from 系统管理_片区字典视图 where 编号='" + 片区 + "'")
        If lobjRec.recordcount > 0 Then
            lstrTemp = lobjRec!名称
        End If
    Else
    End If
    片区名称 = lstrTemp
    
End Property
