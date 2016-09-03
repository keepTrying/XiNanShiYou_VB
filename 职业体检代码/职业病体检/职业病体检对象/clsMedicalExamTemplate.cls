VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedicalExamTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'关键属性。
Private mstr体检表名 As String

'可写的简单属性。
Private mstr试管字母编号 As String
Public 诊断处理意见 As String
Public 代号 As String
Public 收费标准 As String
Public 是否复查体检表 As Boolean
Public 体检单名称 As String
Public 体检人员类型_ger As String 'german
Public 体检类别_ger As String 'german

'修改：2002-7-26（增加“是否年检表”属性）。
Public 是否年检表 As Boolean

'只读的复查属性。
Private mcol附加项目集 As Collection '[附加项目,是否必录,数据类型,数据长度,枚举值]key:附加项目。
Private mcol体检项目集 As Collection '[ClsTestItem]key:编码。
Private mcol体检结论集 As Collection '[体检结论ID,名称]key:体检结论ID。

Private mbln是否已存在 As Boolean    '表明本对象在库中是否已存在。

Private mbln已获取附加项目 As Boolean
Private mbln已获取体检项目 As Boolean
Private mbln已获取体检结论 As Boolean

Public Property Get 是否已存在() As Boolean
    是否已存在 = mbln是否已存在
End Property

Public Property Get 体检表名() As String
Attribute 体检表名.VB_Description = "'功能：设置本杜想的关键属性，从而获取本对象的其他属性。\r\n"
    体检表名 = mstr体检表名
End Property

'功能：设置本杜想的关键属性，从而获取本对象的其他属性。
'作者：刘浩
Public Property Let 体检表名(ByVal vNewValue As String)
    Dim lobjRec As Object        '从库中获取的纪录。
    Dim lstrSql As String        '要执行的sql语句。
    
    On Error GoTo errHandler
    
    mstr体检表名 = vNewValue
    
    '其他附加信息，只有当访问时才获取。
    Set mcol附加项目集 = New Collection
    Set mcol体检结论集 = New Collection
    Set mcol体检项目集 = New Collection
    mbln已获取附加项目 = False
    mbln已获取体检项目 = False
    mbln已获取体检结论 = False
    
    mbln是否已存在 = False
    
    '获取体检表模板的基本信息。
    lstrSql = "select * from " & _
                "职业病体检_体检表模板基本信息表 where 体检表名称 = '" & vNewValue & "'"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(lstrSql)
    If lobjRec.RecordCount > 0 Then
        mstr试管字母编号 = lobjRec!试管编号字母
        诊断处理意见 = IIf(IsNull(lobjRec!诊断处理意见), "", lobjRec!诊断处理意见)
        是否复查体检表 = IIf(IIf(IsNull(lobjRec!是否复查体检表), 0, lobjRec!是否复查体检表) = 1, True, False)
        体检单名称 = IIf(IsNull(lobjRec!体检单名称), "", lobjRec!体检单名称)
        代号 = IIf(IsNull(lobjRec!代号), "", lobjRec!代号)
        收费标准 = IIf(IsNull(lobjRec!收费标准), "", lobjRec!收费标准)
        
        '修改：2002-7-26（增加“是否年检表”属性）。
        是否年检表 = IIf(IIf(IsNull(lobjRec!是否复查体检表), 0, lobjRec!是否复查体检表) = 2, True, False)
        
        lobjRec.Close
        
        mbln是否已存在 = True
    Else
        subClear
        
        '自动分配代号。
        
    End If
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Property Let 体检表名", 6666, lstrError, True
    Exit Property
    Resume
End Property

Public Property Get 试管字母编号() As Variant
    On Error Resume Next
    '替换中文逗号为英文逗号。
    mstr试管字母编号 = gffuncStrReplace(mstr试管字母编号, "，", ",")
    试管字母编号 = mstr试管字母编号
End Property

Public Property Let 试管字母编号(ByVal vNewValue As Variant)
    Dim lstrItem As String
    
    On Error Resume Next
    
    '替换中文逗号为英文逗号。
    vNewValue = gffuncStrReplace(vNewValue, "，", ",")
    
    Do While Right(vNewValue, 1) = ","
        vNewValue = Left(vNewValue, Len(vNewValue) - 1)
    Loop
    mstr试管字母编号 = vNewValue
    
    On Error GoTo errHandler
    '修改：2001-8-23。
    '检查试管子母是否超过2位。
'    If vNewValue <> "" Then
'        If Right(vNewValue, 1) <> "," Then vNewValue = vNewValue & ","
'        Do While vNewValue <> ""
'            lstrItem = Left(vNewValue, InStr(vNewValue, ",") - 1)
'            vNewValue = Right(vNewValue, Len(vNewValue) - InStr(vNewValue, ","))
'            If Len(Trim(lstrItem)) > 2 Then
'                Err.Raise 6666, , "系统没有作保存处理。因为系统规定：“试管字母”中每个可选的字母不能超过2个字符，但你输入的试管字母中“" & lstrItem & "”就超过了2个字符。解决办法：" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "(1) 请减少超过3个字符的试管字母的字符数（比如：把“" & lstrItem & "”改为“" & Left(lstrItem, 2) & "”）。"
'            ElseIf Len(Trim(lstrItem)) = 0 Then
'                Err.Raise 6666, , "系统没有作保存处理。因为系统规定：“试管字母”中每个可选的字母不能为空，即两个逗号之间必须录入字母。"
'            End If
'        Loop
'    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Property Let 试管字母编号", 6666, lstrError, True
End Property

'返回：[附加项目,是否必录,数据类型,数据长度,枚举值]key:附加项目。
'作者：刘浩
Public Property Get 基本附加项目集() As Collection
Attribute 基本附加项目集.VB_Description = "'返回：[附加项目，是否必录]key:附加项目。\r\n"
    Dim lcolItem As Variant     'mcol附加项目中的元素。
    Dim lcolInfo As Collection  '要返回的集合。
    
    On Error GoTo errHandler
    
    '若还未获取附加项目，则获取它。
    If Not mbln已获取附加项目 Then
        sub获取附加项目
    End If
    
    Set lcolInfo = New Collection
    For Each lcolItem In mcol附加项目集
        lcolInfo.Add lcolItem, lcolItem("附加项目")
    Next
    Set 基本附加项目集 = lcolInfo
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Property Get 基本附加项目集", 6666, lstrError, True
End Property

'返回：[ClsTestItem]key:编码。
'作者：刘浩
Public Property Get 体检项目集() As Collection
Attribute 体检项目集.VB_Description = "'返回：[ClsTestItem]key:编码。\r\n"
    Dim lobjItem As Variant     'mcol体检项目集中的元素,clsTestItem。
    Dim lcolInfo As Collection  '要返回的集合。
    
    On Error GoTo errHandler
    
    '若还未获取体检项目，则获取它。
    If Not mbln已获取体检项目 Then
        sub获取体检项目
    End If
    
    Set lcolInfo = New Collection
    For Each lobjItem In mcol体检项目集
        lcolInfo.Add lobjItem, lobjItem.编码
    Next
    Set 体检项目集 = lcolInfo
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Property Get 体检项目集", 6666, lstrError, True
End Property

'返回：[体检结论ID,名称]key:体检结论ID。
'作者：刘浩
Public Property Get 体检结论集() As Collection
Attribute 体检结论集.VB_Description = "'返回：[体检结论ID,名称]key:体检结论ID。\r\n"
    Dim lcolItem As Variant     'mcol体检结论集中的元素。
    Dim lcolInfo As Collection  '要返回的集合。
    
    On Error GoTo errHandler
    '若还未获取体检结论，则获取它。
    If Not mbln已获取体检结论 Then
        sub获取体检结论
    End If
    
    Set lcolInfo = New Collection
    For Each lcolItem In mcol体检结论集
        lcolInfo.Add lcolItem, Trim(Str(lcolItem("体检结论ID")))
    Next
    Set 体检结论集 = lcolInfo

    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Property Get 体检结论集", 6666, lstrError, True
End Property


'功能：在本对象属性"基本附加项目集"中增加元素，不立刻写库。
'作者：刘浩
Public Sub Sub添加附加项目(ByVal para附加项目 As String, ByVal para是否必录 As Boolean)
Attribute Sub添加附加项目.VB_Description = "'功能：在本对象属性""基本附加项目集""中增加元素，不立刻写库。\r\n"
    Dim lcolItem As Collection '[附加项目，是否比录]。
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    '若还未获取附加项目，则获取它。
    If Not mbln已获取附加项目 Then
        sub获取附加项目
    End If
    
    '先判断要添加的项目是否已在集合中存在。
    If sffunc判断集合键值是否存在(mcol附加项目集, para附加项目) Then
        '已存在，只是修改。
        Set lcolItem = mcol附加项目集(para附加项目)
        lcolItem.Remove "是否必录"
        lcolItem.Add para是否必录, "是否必录"
    Else
        '添加。
        Set lcolItem = New Collection
        lcolItem.Add para附加项目, "附加项目"
        lcolItem.Add para是否必录, "是否必录"
        '获取体检附加项目的其他属性。
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData("select * from 职业病体检_体检人员附加项目设置表 where 附加项目='" & para附加项目 & "'")
        If lobjRec.RecordCount > 0 Then
            lcolItem.Add IIf(IsNull(lobjRec("数据类型")), "3", lobjRec("数据类型").Value), "数据类型"
            lcolItem.Add IIf(IsNull(lobjRec("数据长度")), 10, lobjRec("数据长度").Value), "数据长度"
            lcolItem.Add IIf(IsNull(lobjRec("枚举值")), "", lobjRec("枚举值").Value), "枚举值"
        Else
            Err.Raise 6666, , "你选中的附加项目已被人在“业务设置”处删除掉了。请重新选择或输入要修改的体检表，以便刷新可选择的体检附加项目列表。"
        End If
        mcol附加项目集.Add lcolItem, para附加项目
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub添加附加项目", 6666, lstrError, True
End Sub
'功能：附加项目一移一个位置，不立刻写库。
'作者：杨春
Public Sub sub附加项目下移(ByVal para附加项目 As String)
    Dim i As Long
    Dim lcolItem As Collection
    On Error GoTo errHandler
    '查找当前附加项目的位置。
    i = 1
    Do While i <= mcol附加项目集.Count
        If mcol附加项目集(i)("附加项目") = para附加项目 Then
            '找到了当前附加项目的位置。
            Exit Do
        End If
        i = i + 1
    Loop
    
    If i < mcol附加项目集.Count Then
        '在最后一项前面，可以下移。
        Set lcolItem = mcol附加项目集(i)
        mcol附加项目集.Remove i
        If i < mcol附加项目集.Count Then
            mcol附加项目集.Add lcolItem, lcolItem("附加项目"), i + 1
        Else
            mcol附加项目集.Add lcolItem, lcolItem("附加项目")
        End If
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "sub附加项目下移", Err.Number, Err.Description, True
End Sub
'功能：附加项目一移一个位置，不立刻写库。
'作者：杨春
Public Sub sub附加项目上移(ByVal para附加项目 As String)
    Dim i As Long
    Dim lcolItem As Collection
    On Error GoTo errHandler
    '查找当前附加项目的位置。
    i = 1
    Do While i <= mcol附加项目集.Count
        If mcol附加项目集(i)("附加项目") = para附加项目 Then
            '找到了当前附加项目的位置。
            Exit Do
        End If
        i = i + 1
    Loop
    
    If i > 1 Then
        '在第一项后面，可以上移。
        Set lcolItem = mcol附加项目集(i)
        mcol附加项目集.Remove i
        mcol附加项目集.Add lcolItem, lcolItem("附加项目"), i - 1
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "sub附加项目上移", 6666, lstrError, True
End Sub

'功能：在本对象属性"基本附加项目集"中删除元素，不立刻写库。
'作者：刘浩
Public Sub Sub删除附加项目(ByVal para附加项目 As String)
Attribute Sub删除附加项目.VB_Description = "'功能：在本对象属性""基本附加项目集""中删除元素，不立刻写库。\r\n"
    On Error GoTo errHandler
    '若还未获取附加项目，则获取它。
    If Not mbln已获取附加项目 Then
        sub获取附加项目
    End If
    
    If sffunc判断集合键值是否存在(mcol附加项目集, para附加项目) Then
        mcol附加项目集.Remove para附加项目
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub删除附加项目", 6666, lstrError, True
End Sub

'功能：在本对象属性"体检项目集"增加元素，不立刻写库。
'作者：刘浩
Public Sub Sub添加体检项目(ByVal para编码 As String)
Attribute Sub添加体检项目.VB_Description = "'功能：在本对象属性""体检项目集""增加元素，不立刻写库。\r\n"
    Dim lobjItem As ClsTestItem  '新添加的体检项目。
    
    On Error GoTo errHandler
    '若还未获取体检项目，则获取它。
    If Not mbln已获取体检项目 Then
        sub获取体检项目
    End If
    
    '先判断要添加的项目是否已在集合中存在。
    If Not sffunc判断集合键值是否存在(mcol体检项目集, para编码) Then
        '添加。
        Set lobjItem = New ClsTestItem
        lobjItem.编码 = para编码
        
        '判断体检项目是否存在。
        If Not lobjItem.是否存在 Then
            Err.Raise 6666, , "你选择的编码为“" & para编码 & "”的体检项目已不存在。请进入“业务设置”操作界面的“体检项目设置”功能界面上设置体检项目。"
        End If
        mcol体检项目集.Add lobjItem, para编码
    End If

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub添加体检项目", 6666, lstrError, True
End Sub

'功能：在本对象属性"体检项目集"中增加元素，不立刻写库。
'作者：刘浩
Public Sub Sub删除体检项目(ByVal para编码 As String)
Attribute Sub删除体检项目.VB_Description = "'功能：在本对象属性""体检项目集""中增加元素，不立刻写库。\r\n"
    On Error GoTo errHandler
    
    '若还未获取体检项目，则获取它。
    If Not mbln已获取体检项目 Then
        sub获取体检项目
    End If
    
    If sffunc判断集合键值是否存在(mcol体检项目集, para编码) Then
        mcol体检项目集.Remove para编码
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub删除体检项目", 6666, lstrError, True
End Sub

'功能：在本对象属性"体检结论集"中增加元素，不立刻写库。
'作者：刘浩
Public Sub Sub添加体检结论(ByVal para体检结论ID As Long)
Attribute Sub添加体检结论.VB_Description = "'功能：在本对象属性""体检结论集""中增加元素，不立刻写库。\r\n"
    Dim lcolItem As Collection '[体检项目id，名称]。
    Dim lobjRec As Object      '用来获取体检结论名称的记录。
    
    On Error GoTo errHandler
    
    '若还未获取体检结论，则获取它。
    If Not mbln已获取体检结论 Then
        sub获取体检结论
    End If
    
    '先判断要添加的项目是否已在集合中存在。
    If Not sffunc判断集合键值是否存在(mcol体检结论集, Trim(Str(para体检结论ID))) Then
        '添加。
        Set lcolItem = New Collection
        lcolItem.Add para体检结论ID, "体检结论ID"
        
        '用字典对象根据体检结论ID获取名称。
        If pobjDict Is Nothing Then
            Err.Raise 6666, , "字典对象创建不成功。请退出系统，重新注册“字典管理.dll”后，再进行本操作。"
        Else
            Set lobjRec = pobjDict.Fetch("体检结论字典", "InnerID=" & Trim(Str(para体检结论ID)))
            If lobjRec.RecordCount > 0 Then
                lcolItem.Add lobjRec("名称").Value, "名称"
                mcol体检结论集.Add lcolItem, Trim(Str(para体检结论ID))
            Else
                Err.Raise 6666, , "在系统字典中找不到内部ID =" & para体检结论ID & "的体检结论字典内容。请系统管理员进入字典管理操作界面察看“体检结论字典”的内容。"
            End If
        End If
    End If

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub添加体检结论", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：在本对象属性"体检结论集"中删除元素，不立刻写库。
'作者：刘浩
Public Sub Sub删除体检结论(ByVal para体检结论ID As Long)
Attribute Sub删除体检结论.VB_Description = "'功能：在本对象属性""体检结论集""中删除元素，不立刻写库。\r\n"
    On Error GoTo errHandler
    '若还未获取体检项目，则获取它。
    If Not mbln已获取体检结论 Then
        sub获取体检结论
    End If

    If sffunc判断集合键值是否存在(mcol体检结论集, Trim(Str(para体检结论ID))) Then
        mcol体检结论集.Remove Trim(Str(para体检结论ID))
    End If

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub删除体检结论", 6666, lstrError, True
End Sub

'功能：保存本对象所描述的体检表模板信息到数据库中。
'原作者：刘浩
'修改者:何啸天
Public Sub Sub保存模板()
Attribute Sub保存模板.VB_Description = "'功能：保存本对象所描述的体检表模板信息到数据库中。\r\n"
    Dim lobjRec As Object
    Dim lobjItem As Variant '体检项目集合、附加项目集合、体检结论集合中的元素。
    Dim lstrSql As String
    Dim ID As String
    Dim i As Long
    
    On Error GoTo errHandler
    '合法性检查。
    If mstr体检表名 = "" Then
        Err.Raise 6666, , "体检表名称不允许空。"
    End If
    If mstr试管字母编号 = "" Then
        Err.Raise 6666, , "试管字母不允许空。"
    End If
    If UCase(mstr试管字母编号) <> mstr试管字母编号 Then
        Err.Raise 6666, , "试管字母必须全部大写。"
    End If
    If Len(诊断处理意见) > 250 Then
        Err.Raise 6666, , "诊断处理意见过长，已经超过系统的规定长度250个字符。请慎重选择诊断处理意见。"
    End If
'    '判断代号是否重复。
'    If 代号 <> "" Then
'        Set lobjRec = dafuncGetData("select 体检表名称 from 体检管理_体检表模板基本信息表 where 代号='" & 代号 & "' and 体检表名称<>'" & mstr体检表名 & "'")
'        If lobjRec.recordcount > 0 Then
'            Err.Raise 6666, , "体检表代号必须唯一。该体检表代号已经被“" & lobjRec!体检表名称 & "”使用。"
'        End If
'    End If
    
    '开始事务。
    dasubBeginTran
        
    '若没有代号，自动分配代号。
    If 代号 = "" Then
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData("select max(cast(代号 as int)) from 职业病体检_体检表模板基本信息表")
        代号 = Format$(IIf(IsNull(lobjRec(0)), 0, lobjRec(0)) + 1, "00")
    End If
    
    '检查体检表模板在库中的情况与当前是否已存在标志是否一致。
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select * from 职业病体检_体检表模板基本信息表 where 体检表名称 = '" & mstr体检表名 & "'")
    If lobjRec.RecordCount > 0 Then
        '该体检表模板已存在。
        If Not mbln是否已存在 Then
            Err.Raise 6666, , "新建的体检表名称已被使用，请重新输入体检表名称。"
        End If
    ElseIf mbln是否已存在 Then
        '你正在修改的体检表模板已被人删除。
        mbln是否已存在 = False
    End If
    
    '先删除体检表模板附加信息。
    If mbln是否已存在 Then
        dafuncGetData "delete 职业病体检_体检表模板附加项目信息表 where 体检表名称='" & mstr体检表名 & "'"
        dafuncGetData "delete 职业病体检_体检表模板体检项目表 where 体检表名称='" & mstr体检表名 & "'"
        dafuncGetData "delete 职业病体检_体检表模板体检结论表 where 体检表名称='" & mstr体检表名 & "'"
    End If
    
    Set lobjRec = dafuncGetData("select max(cast(编号 as int)) from 职业病体检_体检表模板基本信息表")
    ID = Format$(IIf(IsNull(lobjRec(0)), 0, lobjRec(0)) + 1, "0")
    
    '修改：根据新的信息内容，扩充SQL语句
    '修改者:何啸天、
    '日期:202.3.16
    'german
    If mbln是否已存在 Then
        '修改。
        lstrSql = "update 职业病体检_体检表模板基本信息表 set 试管编号字母='" & 试管字母编号 & "',诊断处理意见='" & 诊断处理意见 & "'," _
                & "体检单名称='" & 体检单名称 & "',代号='" & 代号 & "',是否复查体检表=" & IIf(是否复查体检表, 1, IIf(是否年检表, 2, 0)) & "," _
                & "收费标准='" & 收费标准 & "',体检人员类型='" & 体检人员类型_ger & "',体检类别='" + 体检类别_ger + "' where 体检表名称='" & mstr体检表名 & "'"
    Else
        '新增。
        lstrSql = "insert into 职业病体检_体检表模板基本信息表(体检表名称,试管编号字母,诊断处理意见,体检单名称,代号,是否复查体检表,收费标准,编号,体检人员类型,体检类别) values(" _
                    & "'" & mstr体检表名 & "','" & 试管字母编号 & "','" & 诊断处理意见 & "','" & 体检单名称 & "'," _
                    & "'" & 代号 & "'," & IIf(是否复查体检表, 1, IIf(是否年检表, 2, 0)) & ",'" & 收费标准 & "','" + ID + "','" + 体检人员类型_ger + "','" + 体检类别_ger + "')"
    End If
    
    'MsgBox lstrSql, , "消息"
    dafuncGetData lstrSql
    
    '若还未获取附加项目，则获取它。
    If Not mbln已获取附加项目 Then
        sub获取附加项目
    End If
    '保存附加项目。
    i = 1
    For Each lobjItem In mcol附加项目集
        lstrSql = "insert into 职业病体检_体检表模板附加项目信息表(体检表名称,附加项目,是否必录,序号) values('" & mstr体检表名 & "','" & lobjItem("附加项目") & "','" & IIf(lobjItem("是否必录"), 1, 0) & "'," & i & ")"
        dafuncGetData lstrSql
        i = i + 1
    Next
    
    '若还未获取体检项目，则获取它。
    If Not mbln已获取体检项目 Then
        sub获取体检项目
    End If
    '保存体检项目。
    For Each lobjItem In mcol体检项目集
        lstrSql = "insert into 职业病体检_体检表模板体检项目表(体检表名称,体检项目) values('" & mstr体检表名 & "','" & lobjItem.编码 & "')"
        dafuncGetData lstrSql
    Next
    
    '若还未获取体检项目，则获取它。
    If Not mbln已获取体检结论 Then
        sub获取体检结论
    End If
    '保存体检结论。
'    For Each lobjItem In mcol体检结论集
'        lstrSql = "insert into 职业病体检_体检表模板体检结论表(体检表名称,体检结论) values('" & mstr体检表名 & "'," & lobjItem("体检结论ID") & ")"
'        dafuncGetData lstrSql
'    Next
    
    '根据附加项目设置录入配置信息。
    subSetInputFrameInfo
    
    '设置已存在标志。
    mbln是否已存在 = True
    
    '提交事务。
    dasubCommitTran
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub保存模板", 6666, lstrError, True
    
    Exit Sub
    Resume
End Sub

'功能：删除库中本体检表模板。
'添加者：秦友龙
'功能：触发器没有起做用，貌似数据库里面没有触发器
'修改日期：2012-4-6
Public Sub Sub删除模板()
Attribute Sub删除模板.VB_Description = "'功能：删除库中本体检表模板。\r\n"
    On Error GoTo errHandler
    dasubBeginTran
    
    '删除基本信息，其他附加信息通过触发器级连删除。
    dafuncGetData "delete 职业病体检_体检表模板基本信息表 where 体检表名称='" & mstr体检表名 & "'"
    '//////////////////////////////////////////////////////////////////////
    '修改日期：2012-4-6 修改人：秦友龙
    dafuncGetData "delete 职业病体检_体检表模板体检结论表 where 体检表名称='" & mstr体检表名 & "'"
    dafuncGetData "delete 职业病体检_体检表模板体检项目表 where 体检表名称='" & mstr体检表名 & "'"
    '//////////////////////////////////////////////////////////////////////
    '删除录入配置信息。
    'dafuncGetData "delete 系统管理_录入信息配置表  where 配置名称 ='" & mstr体检表名 & "'"
    dasubCommitTran
    
    '清空本对象。
    subClear
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = Err.Description
    If Err.Number = P_ERRORCODE_VIOLATE Then
        lstrError = "该体检表模板已被使用，不允许删除。"
    Else
        lstrError = func错误处理(Err.Number, Err.Description)
    End If
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub删除模板", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：把本对象复制为一个新体检表模板。即修改"体检表名称"属性，但不立刻写库。
'作者：刘浩。
Public Sub Sub复制模板(ByVal para新体检表名称 As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    '检查复制后名称是否已存在。
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select * from 职业病体检_体检表模板基本信息表 where 体检表名称 = '" & para新体检表名称 & "'")
    If lobjRec.RecordCount > 0 Then
        Err.Raise 6666, , "新体检表名称已被使用，请重新输入新体检表名称。"
    End If
    
    mstr体检表名 = para新体检表名称
    
    '清除已存在标志。
    mbln是否已存在 = False
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "Sub复制模板", 6666, lstrError, True
End Sub

'功能：更换体检表模板名称。即修改"体检表名称"属性，立刻写库。
'作者：刘浩。
Public Sub sub更换体检表名称(ByVal para新体检表名称 As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    dasubBeginTran
    
    '检查复制后名称是否已存在。
    'german 已修改
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select * from 职业病体检_体检表模板基本信息表 where 体检表名称 = '" & para新体检表名称 & "'")
    If lobjRec.RecordCount > 0 Then
        Err.Raise 6666, , "新体检表名称已被使用，请重新输入新体检表名称。"
    End If
    
    If mstr体检表名 <> "" Then
        '修改库中体检表名称。
        'german 已修改
        dafuncGetData "update 职业病体检_体检表模板基本信息表 set 体检表名称 = '" & para新体检表名称 & "' where 体检表名称='" & mstr体检表名 & "'"
        '替换原有的表名？
        
        '修改录入配置信息。
        dafuncGetData "update 系统管理_录入信息配置表 set 配置名称 = '" & para新体检表名称 & "' where 配置名称 ='" & mstr体检表名 & "'"
    End If
    
    '修改对象属性。
    mstr体检表名 = para新体检表名称
    
    dasubCommitTran
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "sub更换体检表名称", 6666, lstrError, True
    Exit Sub
    Resume
End Sub


'功能：清空本对象。
Private Sub subClear()
    On Error Resume Next
'    试管字母编号 = ""
    体检单名称 = ""
    代号 = ""
    收费标准 = ""
    诊断处理意见 = ""
    是否复查体检表 = False
    
    '修改：2002-7-26（增加“是否年检表”属性）。
    是否年检表 = False
    
    Set mcol附加项目集 = New Collection
    Set mcol体检结论集 = New Collection
    Set mcol体检项目集 = New Collection
    
    mbln是否已存在 = False
End Sub

'功能：根据附加项目设置录入配置信息
'      （到录入信息配置表中，以便体检登记界面初始化附加信息录入板控件）。
Private Sub subSetInputFrameInfo()
    Dim llngRow As Long       '当前录入框的行号。
    Dim llngCol As Long       '当前录入框的列号。
    Dim llngWidth As Long     '当前录入框的宽度。
    Dim llngMaxWidth As Long  '录入控件的最大宽度。
    Dim lstrSql As String     '要执行的sql语句。
    Dim lobjItem As Variant   'collection
    Dim llngNo As Long        '序号。
    On Error GoTo errHandler
    
    '若还未获取附加项目，则获取它。
    If Not mbln已获取附加项目 Then
        sub获取附加项目
    End If

    '先删除录入信息配置表中的旧记录。
    lstrSql = "delete 系统管理_录入信息配置表 where 配置名称='" & mstr体检表名 & "'"
    dafuncGetData lstrSql
    
    '依次计算各录入框的行号、列号、宽度。
    llngRow = 1
    llngCol = 1
    llngMaxWidth = 61
    llngNo = 1
    For Each lobjItem In mcol附加项目集
        '计算宽度。
        llngWidth = lobjItem("数据长度")
        If llngCol + llngWidth + 4 > llngMaxWidth Then
            llngRow = llngRow + 1
            llngCol = 1
            llngWidth = lobjItem("数据长度")
            If llngWidth + 5 > llngMaxWidth Then
                llngWidth = llngMaxWidth - 4
            End If
        End If
        lstrSql = "insert into 系统管理_录入信息配置表(配置名称,序号,TabIndex,名称,状态文本,标题," & _
                    "最大行,录入板宽为几列,所在行,起始列,宽度,数据类型,长度," & _
                    "缺省值,保存缺省值,最大值,最小值,可等于最大值,可等于最小值,允许为空,密码框," & _
                    "检查字典项,可见,Enable,主键,字典名称,显示字典字段,保存字典字段)" & _
                    "values('" & mstr体检表名 & "'," & llngNo & "," & llngNo & ", '" & lobjItem("附加项目") & "','','" & lobjItem("附加项目") & "'," & _
                    "5,61," & llngRow & "," & llngCol & "," & llngWidth & "," & lobjItem("数据类型") & "," & lobjItem("数据长度") & "," & _
                    "'','','','',1,1,'" & IIf(lobjItem("是否必录"), 0, 1) & "','0'," & _
                    IIf(lobjItem("枚举值") = "", "'0'", "'1'") & ",'1','1','0','" & lobjItem("枚举值") & "'," & IIf(lobjItem("枚举值") = "", "'',''", "'名称','编号'") & " )"
                    
        dafuncGetData lstrSql
        
        '计算该行下个录入框的列号。
        llngCol = llngCol + llngWidth + 1
        If llngCol >= llngMaxWidth Then
            llngRow = llngRow + 1
            llngCol = 1
        End If
        llngNo = llngNo + 1
    Next
    

    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "subSetInputFrameInfo", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
Private Sub Class_Initialize()
    subClear
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mcol附加项目集 = Nothing
    Set mcol体检结论集 = Nothing
    Set mcol体检项目集 = Nothing
    
End Sub
Private Sub sub获取附加项目()
    Dim lobjRec As Object        '从库中获取的纪录。
    Dim lcolItem As Collection   '“mcol附加项目集”中的一个元素。
    Dim lstrSql As String        '要执行的sql语句。
    
    On Error GoTo errHandler
    
    '获取当前体检表名的所有基本附加项目：[附加项目,是否必录,数据类型,数据长度,枚举值]key:附加项目。
    lstrSql = "select * from 职业病体检_体检表模板附加项目数据库 where 体检表名称 = '" & mstr体检表名 & "' order by 序号"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(lstrSql)
    While Not lobjRec.EOF
        Set lcolItem = New Collection
        lcolItem.Add lobjRec("附加项目").Value, "附加项目"
        lcolItem.Add IIf(IIf(IsNull(lobjRec("是否必录")), "0", lobjRec("是否必录").Value) = 0, False, True), "是否必录"
        lcolItem.Add IIf(IsNull(lobjRec("数据类型")), "3", lobjRec("数据类型").Value), "数据类型"
        lcolItem.Add IIf(IsNull(lobjRec("数据长度")), 10, lobjRec("数据长度").Value), "数据长度"
        lcolItem.Add IIf(IsNull(lobjRec("枚举值")), "", lobjRec("枚举值").Value), "枚举值"
        mcol附加项目集.Add lcolItem, lcolItem("附加项目")
        
        lobjRec.MoveNext
    Wend
    lobjRec.Close
    
    mbln已获取附加项目 = True
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "sub获取附加项目", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
Private Sub sub获取体检项目()
    Dim lobjRec As Object        '从库中获取的纪录。
    Dim lobjItem As ClsTestItem  '“mcol体检项目集”中的一个元素。
    Dim lstrSql As String        '要执行的sql语句。
    
    On Error GoTo errHandler
    '获取当前体检表名的所有的体检项目：[ClsTestItem]key:编码。
    lstrSql = "select 体检项目 from 职业病体检_体检表模板体检项目表 where 体检表名称 = '" & mstr体检表名 & "'"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(lstrSql)
    While Not lobjRec.EOF
        Set lobjItem = New ClsTestItem
        lobjItem.编码 = lobjRec!体检项目
        mcol体检项目集.Add lobjItem, lobjRec!体检项目
        
        lobjRec.MoveNext
    Wend
    lobjRec.Close
    
    mbln已获取体检项目 = True
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "sub获取体检项目", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
Private Sub sub获取体检结论()
    Dim lobjRec As Object        '从库中获取的纪录。
    Dim lcolItem As Collection   '“mcol体检结论集”中的一个元素。
    Dim lstrSql As String        '要执行的sql语句。
    
    On Error GoTo errHandler
    '获取本体检表的所有体检结论：[体检结论ID,名称]key:体检结论ID。
    lstrSql = "select 体检结论,体检结论名称 from 职业病体检_体检表模板体检结论数据库 where 体检表名称 = '" & mstr体检表名 & "'"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(lstrSql)
    While Not lobjRec.EOF
        Set lcolItem = New Collection
        lcolItem.Add lobjRec("体检结论").Value, "体检结论ID"
        lcolItem.Add lobjRec("体检结论名称").Value, "名称"
        mcol体检结论集.Add lcolItem, Trim(Str(lcolItem("体检结论ID")))
        
        lobjRec.MoveNext
    Wend
    lobjRec.Close
    mbln已获取体检结论 = True
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "clsMedicalExamTemplate", "sub获取体检结论", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：获取收费标准总金额。
'创建：2002-10-10（杨春）。
Public Property Get 收费标准金额() As Double
    Dim lobj收费接口 As Object
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    收费标准金额 = 0
    
    If 收费标准 <> "" Then
        '创建收费接口。
        Set lobj收费接口 = CreateObject("收费接口对象.cls对外接口")
            
        '获取收费标准。
        Set lobjRec = lobj收费接口.func查询收费标准("收费标准名称='" & 收费标准 & "'")
        If Not lobjRec Is Nothing Then
            Do While Not lobjRec.EOF
                收费标准金额 = Format(收费标准金额 + lobjRec("单价") * lobjRec("数量"), "0.00")
                lobjRec.MoveNext
            Loop
        End If
    End If
    Exit Property
errHandler:
    '忽略错误。
End Property

