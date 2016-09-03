VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedicalExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'本对象关键属性。
Private mstr系统编号 As String

Private mlng体检类别 As PENUM_体检类型
Private mstr体检日期 As String


Private mobj体检人员 As clsPersonExamed   '(只读)
Private mobj体检表 As ClsMedicalExamSheet '(只读)

Private mstr试管编号 As String          '(只读)
Private mstr复查系统编号 As String      '(只读)
Private mstr体检结论 As String          '(只读)
Private mstr诊断和处理意见 As String    '(只读)
Private mstr复查体检表名称 As String    '(只读)
Private mstr下结论医师  As String       '(只读)用户编号。
Private mstr下结论医师姓名 As String    '(只读)
Private mlng体检状态 As PENUM_体检状态  '(只读)
Private mstr收费批号 As String          '(只读)

Private mstr体检单号 As String          '2004-01-09（武侯需求新增）

Private mbln是否已存在 As Boolean

'修改：2002-10-10（杨春）为了嘉定定制要求，增加该属性。
Private mstr系统编号名称 As String
Public Property Get 系统编号名称() As String
    系统编号名称 = mstr系统编号名称
End Property
Public Property Let 系统编号名称(ByVal vNewValue As String)
    If vNewValue = "" Then
        mstr系统编号名称 = "系统编号"
    Else
        mstr系统编号名称 = vNewValue
    End If
End Property

'功能：对像初始化。
'作者：邓横
Private Sub Class_Initialize()
    On Error GoTo errHandler
    
    '创建属性对象。
    Set mobj体检人员 = New clsPersonExamed
    Set mobj体检表 = New ClsMedicalExamSheet
    mbln是否已存在 = False
    
    mstr系统编号名称 = "系统编号"
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Class_Initialize", Err.Number, Err.Description, False
End Sub

Public Property Get 是否已存在() As Boolean
    是否已存在 = mbln是否已存在
End Property

'功能：设置（或取）本对象的关键属性，从而可以获取其他属性。
'作者：邓横
Public Property Get 系统编号() As String
Attribute 系统编号.VB_Description = "'功能：设置（或取）本对象的关键属性，从而可以获取其他属性。"
    系统编号 = mstr系统编号
End Property

'功能：给本对象的关键属性赋值，从而获取本对象其它属性。
'作者：邓横
Public Property Let 系统编号(ByVal vNewValue As String)
    Dim lobjRec As Object  'sql语句执行结果。
    
    On Error GoTo errHandler
    mstr系统编号 = vNewValue
    mbln是否已存在 = False
    subClear
    
    '只有到引用“体检表”属性时，才实例化该对象。
    mobj体检表.系统编号 = ""
    
    If mstr系统编号 <> "" Then
        '获取体检基本信息，赋值给各属性。
        Set lobjRec = dafuncGetData("select * from 体检管理_体检基本数据库 where 系统编号='" & vNewValue & "'")
        If lobjRec.recordcount > 0 Then
            mstr试管编号 = IIf(IsNull(lobjRec!试管编号), "", lobjRec!试管编号)
            mstr收费批号 = IIf(IsNull(lobjRec!收费批号), "", lobjRec!收费批号)
            mstr体检结论 = IIf(IsNull(lobjRec!体检结论), "", lobjRec!体检结论)
            mstr体检日期 = IIf(IsNull(lobjRec!体检日期), "", lobjRec!体检日期)
            mstr下结论医师 = IIf(IsNull(lobjRec!下结论医师), "", lobjRec!下结论医师)
            mstr下结论医师姓名 = IIf(IsNull(lobjRec!下结论医师姓名), "", lobjRec!下结论医师姓名)
            mstr诊断和处理意见 = IIf(IsNull(lobjRec!诊断和处理意见), "", lobjRec!诊断和处理意见)
            mstr复查体检表名称 = IIf(IsNull(lobjRec!复查体检表名), "", lobjRec!复查体检表名)
            mstr复查系统编号 = IIf(IsNull(lobjRec!复查系统编号), "", lobjRec!复查系统编号)
            mlng体检类别 = IIf(IsNull(lobjRec!体检类别), P_EXAM_FIRST, lobjRec!体检类别)
            mlng体检状态 = IIf(IsNull(lobjRec!体检状态), P_LOGIN_STATUS, lobjRec!体检状态)
            
            '2004-01-09（武侯需求新增）
            mstr体检单号 = IIf(IsNull(lobjRec!体检单号), "", lobjRec!体检单号)
            
            mobj体检人员.健康档案编号 = lobjRec!健康档案编号
            
            mbln是否已存在 = True
        End If
        
    Else
        '新增体检。
        mstr体检日期 = Format(Date, "yyyy-mm-dd")
    End If
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Property Let 系统编号", 6666, lstrError, True
    Exit Property
    Resume
End Property
Public Sub sub分配系统编号(ByVal paraNewNo As String)
    mstr系统编号 = paraNewNo
    mstr试管编号 = ""
    mbln是否已存在 = False
    mobj体检表.系统编号 = mstr系统编号
End Sub

Public Property Get 体检类别() As PENUM_体检类型
    体检类别 = mlng体检类别
End Property

Public Property Let 体检类别(ByVal vNewValue As PENUM_体检类型)
    mlng体检类别 = vNewValue
End Property

Public Property Get 体检日期() As String
    体检日期 = mstr体检日期
End Property

Public Property Let 体检日期(ByVal vNewValue As String)
    mstr体检日期 = vNewValue
End Property

'修改：2004-1-9（试管编号可以手工录入）
Public Property Let 试管编号(ByVal vNewValue As String)
    mstr试管编号 = vNewValue
End Property
Public Property Get 试管编号() As String
    试管编号 = mstr试管编号
End Property

'修改：2004-01-09（武侯需求新增）
Public Property Get 体检单号() As String
    体检单号 = mstr体检单号
End Property

Public Property Let 体检单号(ByVal vNewValue As String)
    mstr体检单号 = vNewValue
End Property


Public Property Get 复查系统编号() As String
    复查系统编号 = mstr复查系统编号
End Property
Public Property Get 体检结论() As String
    体检结论 = mstr体检结论
End Property
Public Property Get 诊断和处理意见() As String
    诊断和处理意见 = mstr诊断和处理意见
End Property
Public Property Get 复查体检表名称() As String
    复查体检表名称 = mstr复查体检表名称
End Property
Public Property Get 下结论医师() As String
    下结论医师 = mstr下结论医师
End Property
Public Property Get 下结论医师姓名() As String
    下结论医师姓名 = mstr下结论医师姓名
End Property
Public Property Get 体检状态() As PENUM_体检状态
    体检状态 = mlng体检状态
End Property
Public Property Get 收费批号() As String
    收费批号 = mstr收费批号
End Property
Public Property Let 收费批号(ByVal vNewValue As String)
    mstr收费批号 = vNewValue
End Property

Property Get 体检表() As ClsMedicalExamSheet
    On Error GoTo errHandler
    '若体检表对象还未实例化，则实例化它。
    If mobj体检表.系统编号 <> mstr系统编号 Then
        mobj体检表.系统编号 = mstr系统编号
    End If
    
    Set 体检表 = mobj体检表
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Property Get 体检表", Err.Number, Err.Description, True
End Property
Property Get 体检人员() As clsPersonExamed
    Set 体检人员 = mobj体检人员
End Property

'功能：根据生成规则获取系统编号流水号前面固定部分。
'作者：邓横
Public Property Get 系统编号固定部分() As String
Attribute 系统编号固定部分.VB_Description = "'功能：根据生成规则获取系统编号流水号前面固定部分。\r\n"
    Dim lobjSysRule As Object '系统编号生成规则对象。
    Dim lstrTemp As String
    
    '创建系统编号生成规则对象。
    On Error Resume Next
    Set lobjSysRule = CreateObject("系统编号规则.clsSysNoRule")
    lobjSysRule.编号名称 = mstr系统编号名称
    lstrTemp = lobjSysRule.流水号前面固定部分
    
    On Error GoTo errHandler
    If lstrTemp = "" Then
        '根据缺省生成规则获取系统编号流水号前面固定部分:"单位编号+服务器代号+日期"。
        If um防疫站编号 = "" Then
            lstrTemp = "000" & "0" & Format(Date, "yymmdd")
        Else
            lstrTemp = um防疫站编号 & um服务器代号 & Format(Date, "yymmdd")
        End If
    End If
    
    系统编号固定部分 = lstrTemp
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Property Get 系统编号固定部分", Err.Number, Err.Description, True
End Property

'功能：按系统编号的用户自定义生成规则（或缺省规则）生成系统编号。
'      缺省规则：防疫站编号4+服务器代号3+日期6+流水号4。
'返回：生成的系统编号。
'作者：邓横
Public Function Func分配系统编号() As String
Attribute Func分配系统编号.VB_Description = "'功能：按系统编号的用户自定义生成规则（或缺省规则）生成系统编号。\r\n    '      缺省规则：防疫站编号3+服务器代号1+日期6+流水号4。\r\n    '返回：生成的系统编号。\r\n"
    Dim lobjSysRule As Object             '系统编号生成规则对象。
    Dim lcol预留参数 As New Collection    '需要传递给lobjRule的参数（包括：体检表模板代号）。
    Dim lobjRec As Object                 '执行sql语句的结果记录。
    Dim lstrTemp As String                '生成的系统编号。
    
    On Error Resume Next
    
    '修改：2002-6-28（杨春）首先获取漏掉的号。
    Set lobjRec = dafuncGetData("exec 体检管理_预定漏掉的编号 '" & mstr系统编号名称 & "'")
    If lobjRec.recordcount > 0 Then
        If lobjRec(0) <> "" Then
            Func分配系统编号 = lobjRec(0)
            Exit Function
        End If
    End If
    
    '创建系统编号生成规则对象，先按照用户定义规则生成系统编号。
    Set lobjSysRule = CreateObject("系统编号规则.clsSysNoRule")
    lobjSysRule.编号名称 = mstr系统编号名称
    'lcol预留参数.Add "", "体检表模板代号"
    lstrTemp = lobjSysRule.Func生成系统编号(lcol预留参数)
    
    '若上述处理步骤出错，按缺省规则生成系统编号。
    On Error GoTo errHandler
    If lstrTemp = "" Then
        '获取服务器日期。
        Set lobjRec = dafuncGetData("select getdate()")
        lstrTemp = um防疫站编号 & um服务器代号 & Format(lobjRec(0), "yymmdd")
        
        '生成流水号。
        Set lobjRec = dafuncGetData("exec 体检管理_生成编号流水号 '" & mstr系统编号名称 & "'")
        lstrTemp = lstrTemp & lobjRec(0)
        
    End If
    
    Func分配系统编号 = lstrTemp
    
    Exit Function
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Func分配系统编号", Err.Number, Err.Description, True
End Function
'功能：按如下规则生成复查系统编号：原系统编号-复查流水号。
'创建：2002-7-15（杨春）嘉定要求复查系统编号不能占系统编号的流水号，否则健康证号不连续。
Public Function func分配复查系统编号(ByVal para原系统编号 As String) As String
    Dim lobjRec As Object
    
    func分配复查系统编号 = ""
    
    '首先执行扩展存储过程。
    On Error Resume Next
    Set lobjRec = dafuncGetData("体检管理_生成复查系统编号 '" & para原系统编号 & "'")
    If lobjRec.recordcount > 0 Then
        func分配复查系统编号 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    
    '没有扩展存储过程，按系统编号规则生成。
    If func分配复查系统编号 = "" Then
        func分配复查系统编号 = Func分配系统编号
    End If
    
End Function

'功能：获取系统编号的总长度。
'作者：杨春
Public Property Get 系统编号长度() As Long
    Dim lobjSysRule As Object             '系统编号生成规则对象。
    Dim llng长度 As Long                  '系统编号长度。
    
    '创建系统编号生成规则对象，先按照用户定义规则获取系统编号长度。
    On Error Resume Next
    Set lobjSysRule = CreateObject("系统编号规则.clsSysNoRule")
    lobjSysRule.编号名称 = mstr系统编号名称
    llng长度 = lobjSysRule.编号总长度
    
    '若上述处理步骤出错，按缺省规则获取长度。
    On Error GoTo errHandler
    If llng长度 = 0 Then
        llng长度 = Len(um防疫站编号) + Len(um服务器代号) + 10
    End If
    
    系统编号长度 = llng长度
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Property Get 系统编号长度", Err.Number, Err.Description, True
End Property
'功能：获取系统编号的前一个号。
'返回：指定参数的前一个系统编号。
'作者：杨春
Public Function func获取系统编号的前一个号(ByVal para系统编号 As String) As String
    Dim lobjRec As Object
    Dim lstr流水号 As String
    
    On Error GoTo errHandler
    Set lobjRec = dafuncGetData("exec 体检管理_获取编号的上下一个号 '" & mstr系统编号名称 & "','" & para系统编号 & "',-1")
    If lobjRec.recordcount > 0 Then
        func获取系统编号的前一个号 = lobjRec(0)
    Else
        func获取系统编号的前一个号 = ""
    End If
    If func获取系统编号的前一个号 = "" Then
        '按缺省规则返回前一个号。
        lstr流水号 = Right(para系统编号, 4)
        If Val(lstr流水号) = 1 Then
            Err.Raise 6666, , "当前体检记录没有上一条记录。"
        Else
            lstr流水号 = Format(Val(lstr流水号) - 1, "0000")
        End If
        func获取系统编号的前一个号 = Left(para系统编号, Len(para系统编号) - 4) & lstr流水号
    End If
    Exit Function
    
errHandler:
    Dim lstrError As String
    If InStr(Err.Description, "SQL语句") > 0 Then
        lstrError = Left(Err.Description, InStr(Err.Description, "SQL语句") - 1)
    Else
        lstrError = Err.Description
    End If
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "func获取系统编号的前一个号", Err.Number, lstrError, True
End Function
'功能：获取系统编号的前一个号。
'返回：指定参数的前一个系统编号。
'作者：杨春
Public Function func获取系统编号的后一个号(ByVal para系统编号 As String) As String
    Dim lobjRec As Object
    Dim lstr流水号 As String
    
    On Error GoTo errHandler
    Set lobjRec = dafuncGetData("exec 体检管理_获取编号的上下一个号 '" & mstr系统编号名称 & "','" & para系统编号 & "',1")
    If lobjRec.recordcount > 0 Then
        func获取系统编号的后一个号 = lobjRec(0)
    Else
        func获取系统编号的后一个号 = ""
    End If
    If func获取系统编号的后一个号 = "" Then
        '按缺省规则返回后一个号。
        lstr流水号 = Right(para系统编号, 4)
        lstr流水号 = Format(Val(lstr流水号) + 1, "0000")
        func获取系统编号的后一个号 = Left(para系统编号, Len(para系统编号) - 4) & lstr流水号
    End If
    Exit Function
    
errHandler:
    Dim lstrError As String
    If InStr(Err.Description, "SQL语句") > 0 Then
        lstrError = Left(Err.Description, InStr(Err.Description, "SQL语句") - 1)
    Else
        lstrError = Err.Description
    End If
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "func获取系统编号的后一个号", Err.Number, lstrError, True
End Function

'功能：退回系统编号，即从系统编号的最大流水号减一。
'作者：杨春
Public Sub sub退回系统编号(ByVal para系统编号 As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    dasubBeginTran
        
    '首先判断该系统编号的记录是否已存在。
    Set lobjRec = dafuncGetData("select 系统编号 from 体检管理_体检基本信息表 where 系统编号='" & para系统编号 & "'")
    If lobjRec.recordcount = 0 Then
        '该系统编号没有保存，可以退回。
        dafuncGetData "exec 体检管理_退回编号流水号 '" & mstr系统编号名称 & "','" & para系统编号 & "'"
    End If
    
    dasubCommitTran
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "sub退回系统编号", 6666, lstrError, True
End Sub

'功能：复查系统编号若有扩展的生成规则，不作退回处理，否则需要。
'创建：2002-7-15（杨春）嘉定要求复查系统编号不能占用系统编号的流水号，否则，健康证号不连续。
Public Sub sub退回复查系统编号(ByVal para系统编号 As String)
    Dim lobjRec As Object
    Dim lstr复查编号 As String
    
    
    '判断是否有扩展存储过程。若有，不需要作退回处理，否则需要。
    On Error Resume Next
    lstr复查编号 = ""
    Set lobjRec = dafuncGetData("体检管理_生成复查系统编号 '" & para系统编号 & "'")
    If lobjRec.recordcount > 0 Then
        lstr复查编号 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    
    On Error GoTo errHandler
    If lstr复查编号 = "" Then
        sub退回系统编号 para系统编号
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "sub退回系统编号", 6666, lstrError, True
End Sub

'功能：对当前对象的属性作合法性判断。
'输出：paraErrorInfo   若不合法时，该参数表明不合法的具体错误信息。
'返回：检查结果  true合法/false不合法。
'作者：杨春
Public Function func合法性检查(paraErrorInfo As String) As Boolean
Attribute func合法性检查.VB_Description = "'功能：对当前对象的属性作合法性判断。\r\n    '输出：paraErrorInfo   若不合法时，该参数表明不合法的具体错误信息。\r\n    '返回：检查结果  true合法/false不合法。\r\n"
    Dim lobjRec As Object   '执行sql语句的结果。
    
    On Error GoTo errHandler
    paraErrorInfo = ""
    
    
    '若体检表对象还未实例化，则实例化它。
    If mobj体检表.系统编号 <> mstr系统编号 Then
        mobj体检表.系统编号 = mstr系统编号
    End If
    
    '检查体检表对象的合法性。
    func合法性检查 = mobj体检表.func合法性检查(paraErrorInfo)
    If Not func合法性检查 Then
        Exit Function
    End If
    
    If mstr体检日期 = "" Then
        paraErrorInfo = "必须输入体检日期。"
        func合法性检查 = False
        Exit Function
    End If
    
    '若是修改，判断是否已开始体检。
    If mbln是否已存在 And mstr系统编号 <> "" Then
        Set lobjRec = dafuncGetData("select 体检状态,体检表名称 from 体检管理_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
        If lobjRec.recordcount > 0 Then
            If IIf(IsNull(lobjRec!体检状态), P_LOGIN_STATUS, lobjRec!体检状态) <> P_LOGIN_STATUS And lobjRec!体检表名称 <> mobj体检表.体检表名 Then
                '已开始体检，并且修改了体检表。这是不允许的。
                paraErrorInfo = "你正在修改的体检人员已开始体检，不允许重新选择体检表。请恢复体检表模板"
                
                func合法性检查 = False
                Exit Function
            End If
        End If
    End If
    Exit Function
errHandler:
    paraErrorInfo = Err.Description
    func合法性检查 = False
End Function


'功能：保存本对象属性描述的体检登记信息（包括建立体检人员的健康档案、体检表）。
'      若是新增体检记录，插入系统编号到体检访问标志表。
'作者：邓横
'修改：2001-4-22（杨春）
'      2004-01-09（武侯需求新增“体检单号”，试管编号可以手工录入）
Public Sub Sub保存体检登记信息()
    Dim lobjRec As Object         '执行sql语句的结果记录。
    Dim lcolInfo As Collection
    Dim lstrErrorInfo As String   '合法性检查返回的错误信息。
    Dim lstrSql As String         '要执行的sql语句。
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '修改：2002-1-14（重新获取是否已存在的标志，以免在第一次保存失败后第二次保存时该标志没有复原导致下面的处理不正确）。
    mbln是否已存在 = False
    If mstr系统编号 <> "" Then
        Set lobjRec = dafuncGetData("select * from 体检管理_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
        If lobjRec.recordcount > 0 Then
            mbln是否已存在 = True
        End If
    Else
        '生成系统编号。
        mstr系统编号 = Func分配系统编号
        mobj体检表.系统编号 = mstr系统编号
    End If
    
    '数据合法性检查。
    If Not func合法性检查(lstrErrorInfo) Then
        Err.Raise 6666, , lstrErrorInfo
    End If
    
    '若是新增体检记录，检查"体检管理_试管编号字母"中作"字母"与当前对象的"体检表.试管编号字母"是否一致。
    If Not mbln是否已存在 And mstr试管编号 = "" Then
        lstrSql = "select 字母,是否固定 from 体检管理_试管编号字母表 where 体检表名称='" & mobj体检表.体检表名 & "' and convert(varchar(10),日期,101)=convert(varchar(10),getdate(),101)"
        Set lobjRec = dafuncGetData(lstrSql)
        If lobjRec.recordcount > 0 Then
            If lobjRec!字母 <> mobj体检表.试管编号字母 Then
                '不一致。
                If IIf(IsNull(lobjRec!是否固定), 0, lobjRec!是否固定) = 1 Then
                    '已固定，已库中固定了的字母为准。
                    '重新按表中字母设置"体检表.试管编号字母"。
                    mobj体检表.试管编号字母 = lobjRec!字母
                
                Else
                    '该字母未固定，以当前体检记录的为准。
                    '修改该表的"字母"，并作上固定标志。
                    lstrSql = "update 体检管理_试管编号字母表 set 字母='" & mobj体检表.试管编号字母 & "',是否固定=1 where 体检表名称='" & mobj体检表.体检表名 & "' and convert(varchar(10),日期,101)=convert(varchar(10),getdate(),101)"
                    dafuncGetData lstrSql
                End If
                
            ElseIf IIf(IsNull(lobjRec!是否固定), 0, lobjRec!是否固定) = 0 Then
                '一致，但库中还未作固定标志。作上固定标志。
                lstrSql = "update 体检管理_试管编号字母表 set 是否固定=1 where 体检表名称='" & mobj体检表.体检表名 & "' and convert(varchar(10),日期,101)=convert(varchar(10),getdate(),101)"
                dafuncGetData lstrSql
            End If
            
        Else
            '库中还没有登记当天该体检表模板使用的试管编号字母。
            '插入新纪录，并且作固定标志。
            lstrSql = "insert into 体检管理_试管编号字母表(体检表名称,字母,日期,是否固定) values('" & mobj体检表.体检表名 & "','" & mobj体检表.试管编号字母 & "',convert(varchar(10),getdate(),101),1)"
            dafuncGetData lstrSql
        End If
        
        '分配试管流水号。
        '修改：2002-10-30（杨春）增加参数体检日期。
        Set lobjRec = dafuncGetData("exec 体检管理_生成试管编号 '" & mobj体检表.试管编号字母 & "','" & Format(mstr体检日期, "yyyy-mm-dd") & "'")
        mstr试管编号 = lobjRec(0)
        'If mobj体检表.试管编号字母 = "" Then mstr试管编号 = Right(mstr试管编号, Len(mstr试管编号) - 1)
    End If
    
    '保存健康档案到数据库中。
    mobj体检人员.Sub保存
    
    '若是修改了体检表模板名，先删除旧的体检项目。
    If mbln是否已存在 Then
        Set lobjRec = dafuncGetData("select 体检表名称,体检状态 from 体检管理_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
        If lobjRec!体检表名称 <> mobj体检表.体检表名 Then
            '判断状态，是否已开始体检。
            If IIf(IsNull(lobjRec!体检状态), 0, lobjRec!体检状态) <> P_LOGIN_STATUS Then
                Err.Raise 6666, , "该体检记录已开始登记体检结果，不允许修改体检表。"
            Else
                dafuncGetData "delete 体检管理_体检结果信息表 where 系统编号='" & mstr系统编号 & "'"
            End If
        End If
    End If
    
    
    '保存体检基本信息。
    If mbln是否已存在 Then
        '已存在，修改记录。
        lstrSql = "update 体检管理_体检基本信息表 set 体检单号='" & mstr体检单号 & "',试管编号='" & mstr试管编号 & "',体检表名称='" & mobj体检表.体检表名 & "'" _
                & ",体检日期='" & mstr体检日期 & " ',收费批号='" & mstr收费批号 & "',体检类别=" & mlng体检类别 _
                & "where 系统编号='" & mstr系统编号 & "'"
    Else
        '新增，插入记录。
        lstrSql = "insert into 体检管理_体检基本信息表(系统编号,健康档案编号,体检单号,试管编号,体检表名称,体检类别,体检日期,收费批号,体检状态) values(" _
                & "'" & mstr系统编号 & "','" & mobj体检人员.健康档案编号 & "','" & mstr体检单号 & "','" & mstr试管编号 & "','" & mobj体检表.体检表名 & "'," _
                & mlng体检类别 & ", '" & mstr体检日期 & "','" & mstr收费批号 & "'," & mlng体检状态 & ")"
    End If
    dafuncGetData lstrSql
    
    '删除旧附加信息。
    If mbln是否已存在 Then
        dafuncGetData "delete 体检管理_体检附加信息表 where 系统编号='" & mstr系统编号 & "'"
    End If
    
    '保存附加信息
    Set lcolInfo = mobj体检表.附加信息
    For i = 1 To lcolInfo.Count
        '2006-8-3(新增，自动生成血号)
        Dim lstr体检表号 As String
        If Not mbln是否已存在 And lcolInfo(i)("附加项目") = "血号" And lcolInfo(i)("项目值") = "" Then
            Set lobjRec = dafuncGetData("exec 体检管理_生成体检表号 '" & Me.体检日期 & "'")
            If lobjRec.recordcount > 0 Then
                lstr体检表号 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
                lcolInfo(i).Remove "项目值"
                lcolInfo(i).Add lstr体检表号, "项目值"
            End If
        End If
        '插入记录。
        lstrSql = "insert into 体检管理_体检附加信息表(系统编号,附加项目,项目值,项目值编号) values(" _
            & "'" & mstr系统编号 & "','" & lcolInfo(i)("附加项目") & "','" & lcolInfo(i)("项目值") & "','" & lcolInfo(i)("项目值编号") & "')"
        dafuncGetData lstrSql
    Next i
    
    
    '保存体检结果信息。
    '修改：2001-4-22（杨春）优化性能。
    '修改：2001-12-29（复查体检项目可以重新选择）。
    If mlng体检类别 = P_EXAM_AGAIN Then
        mobj体检表.Sub保存体检结果
    Else
        '可以修改体检项目。
        mobj体检表.Sub保存体检结果
    End If
    
    '若是新增体检记录，插入系统编号到体检访问标志表。
    If Not mbln是否已存在 Then
        dafuncGetData "insert into 体检管理_体检访问标志表(系统编号) values('" & mstr系统编号 & "')"
    End If
    
    '修改：2002-6-28（杨春）把预定的号从"体检管理_漏掉的编号表"中删除。
    dafuncGetData "delete 体检管理_漏掉的编号表 where 编号名称='" & mstr系统编号名称 & "' and 编号='" & mstr系统编号 & "'"
    
    mbln是否已存在 = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Sub保存体检登记信息", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：保存本对象的"体检表"属性中的所有体检结果，并自动下体检结论。
'假设：已调用"体检表.sub设置体检结果"方法设置好体检结果。
'作者：邓横
'修改：2001-4-20（杨春：把下体检结论抽象为一个新的方法。）
Public Sub Sub保存体检结果()
Attribute Sub保存体检结果.VB_Description = "'功能：保存本对象的""体检表""属性中的所有体检结果，并自动下体检结论。\r\n    '      使用本方法的前提是已调用""体检表.sub设置体检结果""方法设置好体检结果。\r\n"
    On Error GoTo errHandler
    
    '若体检表对象还未实例化，则实例化它。
    If mobj体检表.系统编号 <> mstr系统编号 Then
        mobj体检表.系统编号 = mstr系统编号
    End If
    
    '把"体检表"对象中设置的体检结果保存到数据库中。
    mobj体检表.Sub保存体检结果
    
    '若当前体检表所有体检结果均录入，自动得出体检结论。
    sub下体检结论
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Sub保存体检结果", 6666, lstrError, True
End Sub

'功能：若当前体检表所有体检结果均录入，自动得出体检结论。
'假设：体检结果已保存到体检结果信息表中。
'作者：杨春。
'日期：2001-4-20。
Public Sub sub下体检结论()
    Dim lstr体检结论 As String          '自动得出的体检结论。
    Dim llng体检状态 As PENUM_体检状态  '新体检状态。
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    
    '若体检表对象还未实例化，则实例化它。
    If mobj体检表.系统编号 <> mstr系统编号 Then
        mobj体检表.系统编号 = mstr系统编号
    End If
    
    '设置体检状态=2（若所有体检项目均得出体检结果）。
    Set lobjRec = dafuncGetData("select count(*) from 体检管理_体检结果信息表 where 系统编号='" & mstr系统编号 & "' and isnull(体检结果,'')=''")
    If IIf(IsNull(lobjRec(0)), 0, lobjRec(0)) > 0 Then
        llng体检状态 = P_EXAMING_STATUS    '已开始体检，但还未体检完毕。
    Else
        '所有体检项目均录入，得出体检结论。
        lstr体检结论 = mobj体检表.func下体检结论
    
        llng体检状态 = P_CONCLUED_STATUS   '已得出体检结论。
    End If
    
    '保存体检结论、体检状态到数据库中。
    If (mstr体检结论 <> lstr体检结论 And lstr体检结论 <> "") Or mlng体检状态 <> llng体检状态 Then
        dafuncGetData "update 体检管理_体检基本信息表 set 体检结论='" & lstr体检结论 & "',体检状态=" & llng体检状态 & " where 系统编号='" & mstr系统编号 & "'"
    End If
    
    mstr体检结论 = lstr体检结论
    mlng体检状态 = llng体检状态

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "sub下体检结论", 6666, lstrError, True
End Sub

'功能：保存体检结论，并设置体检状态=3（体检完毕）。
'作者：邓横
Public Sub Sub保存体检结论(para体检结论 As String, _
                        para诊断和处理意见 As String, _
                        para下结论医师编号 As String, _
                        para复查体检表名 As String, _
                        para下结论日期 As String)
    Dim lstrSql As String
    
    On Error GoTo errHandler
    '保存体检结论信息（体检状态=3）到数据库中。
    lstrSql = "update 体检管理_体检基本信息表 set 体检结论='" & para体检结论 & "',诊断和处理意见='" & para诊断和处理意见 & "',下结论日期='" & Format(para下结论日期, "yyyy-mm-dd") & "'" _
            & ",下结论医师=" & IIf(para下结论医师编号 = "", "null", "'" & para下结论医师编号 & "'") & ",复查体检表名=" & IIf(para复查体检表名 = "", "null", "'" & para复查体检表名 & "'") & ",体检状态=" & P_ENDED_STATUS _
            & " where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrSql
    
    '修改对象属性。
    mlng体检状态 = P_ENDED_STATUS
    mstr复查体检表名称 = para复查体检表名
    mstr下结论医师 = para下结论医师编号
    mstr诊断和处理意见 = para诊断和处理意见
    mstr体检结论 = para体检结论
        
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Sub保存体检结论", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：设置体检状态=2，并保存到库中。
'作者：邓横
Public Sub Sub取消体检结论审核()
Attribute Sub取消体检结论审核.VB_Description = "'功能：设置体检状态=2，并保存到库中。\r\n"
    Dim lstrSql As String
    
    On Error GoTo errHandler
    '保存体检结论信息（体检状态恢复到2）到数据库中。
    lstrSql = "update 体检管理_体检基本信息表 set 复查系统编号='', 诊断和处理意见='',下结论医师=null,复查体检表名=null,体检状态=" & P_CONCLUED_STATUS _
            & " where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrSql
    
    '恢复对象属性。
    mlng体检状态 = P_EXAMING_STATUS
    mstr复查体检表名称 = ""
    mstr下结论医师 = ""
    mstr诊断和处理意见 = ""
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Sub取消体检结论审核", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：保存本次体检的复查系统编号到库中。
'作者：邓横
Public Sub Sub保存复查系统编号(ByVal para复查系统编号 As String)
Attribute Sub保存复查系统编号.VB_Description = "'功能：保存本次体检的复查系统编号到库中。\r\n"
    On Error GoTo errHandler

    dafuncGetData "update 体检管理_体检基本信息表 set 复查系统编号='" & para复查系统编号 & "' where 系统编号=" & IIf(mstr系统编号 = "", "null", "'" & mstr系统编号 & "'")
    
    '修改属性。
    mstr复查系统编号 = para复查系统编号
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "Sub保存复查系统编号", 6666, lstrError, True
End Sub

'功能：清空对象属性。
'作者：邓横
Private Sub subClear()
    On Error GoTo errHandler
    mstr复查体检表名称 = ""
    mstr复查系统编号 = ""
    mstr试管编号 = ""
    mstr收费批号 = ""
    mstr体检结论 = ""
    mstr体检日期 = Format(Date, "yyyy-mm-dd")
    mstr下结论医师 = ""
    mstr下结论医师姓名 = ""
    mstr诊断和处理意见 = ""
    mlng体检类别 = P_EXAM_FIRST
    mlng体检状态 = P_LOGIN_STATUS
    Set mobj体检表 = New ClsMedicalExamSheet
    Set mobj体检人员 = New clsPersonExamed
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExam", "subClear", Err.Number, Err.Description, True
End Sub


'输入：para编号类型 1试管编号，2体检单号, 3姓名。
Public Function func根据其他编号获取系统编号(ByVal para编号 As String, ByVal para编号类型 As Long) As String
    Dim lobjRec As Object
    Dim lstr编号名称 As String
    If para编号类型 = 1 Then
        lstr编号名称 = "试管编号"
    ElseIf para编号类型 = 2 Then
        lstr编号名称 = "体检单号"
    Else
        lstr编号名称 = "姓名"
    End If
    Set lobjRec = dafuncGetData("select 系统编号 from 体检管理_体检基本数据库 where " & lstr编号名称 & "='" & para编号 & "' order by 体检日期 desc")
    If lobjRec.recordcount > 0 Then
        func根据其他编号获取系统编号 = lobjRec(0)
    Else
        func根据其他编号获取系统编号 = ""
    End If
    
End Function
