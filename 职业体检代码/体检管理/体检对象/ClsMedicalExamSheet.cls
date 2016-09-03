VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMedicalExamSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------体检表对象-----------

'属性变量。
Private mstr系统编号 As String
Private mstr体检表名 As String
Private mstr试管编号字母 As String
Private mstr体检单名称 As String

Private mcol附加信息 As Collection  '[附加项目,项目值,项目值编号]，key:附加项目。
Private mcol体检项目 As Collection  'item:clsFactTestItem,key:体检项目编码。

Private mbln是否已获取附加项目 As Boolean
Private mbln是否已获取体检项目 As Boolean

Private mbln是否已存在 As Boolean   '表明当前对象是否已在数据库中存在。

Public Property Get 系统编号() As String
    系统编号 = mstr系统编号
End Property

'功能：设置系统编号属性，从而获取其他属性。
'作者：邓恒
Public Property Let 系统编号(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lcolItem As Collection
    Dim i As Integer
        
    On Error GoTo errHandler
    mstr系统编号 = paraSysNo
    
    '先清空其他属性。
    mstr体检表名 = ""
    mstr试管编号字母 = ""
    mstr体检单名称 = ""
    Set mcol附加信息 = New Collection
    Set mcol体检项目 = New Collection
    mbln是否已存在 = False
    mbln是否已获取附加项目 = False
    mbln是否已获取体检项目 = False
    
    If mstr系统编号 <> "" Then
    
        '获取体检表名,试管编号字母。
        Set lobjRec = dafuncGetData("select 体检表名称,试管编号 from 体检管理_体检基本信息表 where 系统编号='" & mstr系统编号 & "'")
        If lobjRec.recordcount > 0 Then
            mstr体检表名 = lobjRec("体检表名称")
            If InStr(lobjRec("试管编号"), ":") > 0 Then
                mstr试管编号字母 = Left(lobjRec("试管编号"), InStr(lobjRec("试管编号"), ":") - 1)
            Else
                mstr试管编号字母 = Left(lobjRec!试管编号, 1)
            End If
            
            lobjRec.Close
            
            '等待引用附加项目、体检项目时才获取。
            'sub获取附加项目
            'sub获取体检项目
            
            mbln是否已存在 = True
        End If
    End If
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Property Let 系统编号", 6666, lstrError, True
    Exit Property
    Resume
End Property
'功能：获取体检表名属性。
'作者：邓恒
Public Property Get 体检表名() As String
Attribute 体检表名.VB_Description = "'功能：设置体检表名属性，同时获取该体检表的当天使用的试管编号字母。"
    体检表名 = mstr体检表名
End Property

'功能：设置体检表名属性，同时获取该体检表的当天使用的试管编号字母。
'作者：邓恒
Public Property Let 体检表名(ByVal para体检表名 As String)
    Dim lobj体检表模板 As clsMedicalExamTemplate
    Dim lcolInfo As Collection            '存放获取的体检表模板的所有附加项目[附加项目，是否必录]。
    Dim lobjItem As Variant               'clsTestItem，体检表模板对象的属性“体检项目集”的元素。
    Dim lcolItem As Collection            '要加入“mcol附加信息”的元素。
    Dim lobjFactItem As ClsFactTestItem   '要加入“mcol体检项目”的元素。
    Dim lobjRec As Object                 '执行sql语句从库中获取的记录。
    Dim i As Long
    
    On Error GoTo errHandler
    
    '若是换了体检表，根据体检表名获取附加项目、体检项目。
    If mstr体检表名 <> para体检表名 Then
        '创建体检表模板对象。
        Set lobj体检表模板 = New clsMedicalExamTemplate
        lobj体检表模板.体检表名 = para体检表名
        mstr体检单名称 = lobj体检表模板.体检单名称
        
        '获取该体检表模板的所有附加项目:[附加项目，是否必录]，并加入mcol附加信息。
        Set lcolInfo = lobj体检表模板.基本附加项目集
        Set mcol附加信息 = New Collection
        For i = 1 To lcolInfo.Count
            Set lcolItem = New Collection
            lcolItem.Add lcolInfo(i)("附加项目"), "附加项目"
            lcolItem.Add "", "项目值"
            lcolItem.Add "", "项目值编号"
            mcol附加信息.Add lcolItem, lcolItem("附加项目")
        Next

        '获取该体检表模板的所有体检项目：[ClsTestItem]，并加入mcol体检项目。
        Set lcolInfo = lobj体检表模板.体检项目集
        Set mcol体检项目 = New Collection
        
        For Each lobjItem In lcolInfo
            Set lobjFactItem = New ClsFactTestItem
            With lobjFactItem
                .系统编号 = mstr系统编号
                .体检项目编号 = lobjItem.编码
                
                '更换了体检表名称，需要重新插入体检结果表。
                .是否已存在 = False
            End With
            mcol体检项目.Add lobjFactItem, lobjFactItem.体检项目编号
        Next
        
        Set lobj体检表模板 = Nothing
        Set lobjFactItem = Nothing
        
        mbln是否已获取附加项目 = True
        mbln是否已获取体检项目 = True
    End If
    
    If Not mbln是否已存在 Then
        '若是新增，根据选择的体检表名获取该体检表当天使用的试管字母。
        mstr试管编号字母 = ""
        If para体检表名 <> "" Then
            '开始事务。
            dasubBeginTran
            
            '判断当天本体检表是否已登记使用了某字母。
            Set lobjRec = dafuncGetData("select 体检表名称,字母,日期,是否固定 from 体检管理_试管编号字母表 where 体检表名称='" & para体检表名 & "' and convert(varchar(10),日期,101)=convert(varchar(10),getdate(),101)")
            If lobjRec.recordcount > 0 Then
                If IIf(IsNull(lobjRec("是否固定")), 0, lobjRec("是否固定")) = 1 Then
                    '获取当天本体检表已使用的字母，设置“试管编号字母”属性。
                    mstr试管编号字母 = lobjRec("字母")
                End If
                
                lobjRec.Close
                If mstr试管编号字母 = "" Then
                    dafuncGetData "delete from 体检管理_试管编号字母表 where 体检表名称='" & para体检表名 & "'"
                End If
            Else
                '没有当天记录，删除该体检表的所有记录。
                dafuncGetData "delete from 体检管理_试管编号字母表 where 体检表名称='" & para体检表名 & "'"
            End If
        
            '提交事务。
            dasubCommitTran
        End If
        
    End If
    
    mstr体检表名 = para体检表名
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Property Let 体检表名", 6666, lstrError, True
End Property

'功能：添加一体检项目,但不立刻写库。（主要用于复查登记）。
'输入：para编码 as string   '体检项目编号。
'创建：2001-8-23。
Public Sub Sub添加体检项目(ByVal para编码 As String)
    Dim lobjFactItem  As ClsFactTestItem
    On Error GoTo errHandler
    
    If Not sffunc判断集合键值是否存在(mcol体检项目, para编码) Then
        Set lobjFactItem = New ClsFactTestItem
        With lobjFactItem
            .系统编号 = mstr系统编号
            .体检项目编号 = para编码
            
            '更换了体检表名称，需要重新插入体检结果表。
            .是否已存在 = False
        End With
        mcol体检项目.Add lobjFactItem, para编码
    End If
    
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "sub添加体检项目", 6666, lstrError, True
End Sub
'功能：删除所有体检项目,但不立刻写库。（主要用于复查登记）。
'输入：无。
'创建：2001-8-23。
Public Sub Sub删除所有体检项目()
    Set mcol体检项目 = New Collection
End Sub

Public Sub Sub删除体检项目(ByVal para编码 As String)
    If sffunc判断集合键值是否存在(mcol体检项目, para编码) Then
        mcol体检项目.Remove para编码
    End If
End Sub


'功能：设置或获取当天该类体检表所采用的试管编号字母。若该属性为空，表示当前体检表在当天是第一张（表示体检登记的界面上用户可以选择试管编号字母，否则用户不可以选择字母）。
'作者：邓恒
Public Property Get 试管编号字母() As String
Attribute 试管编号字母.VB_Description = "'功能：设置或获取当天该类体检表所采用的试管编号字母。若该属性为空，表示当前体检表在当天是第一张（表示体检登记的界面上用户可以选择试管编号字母，否则用户不可以选择字母）。\r\n"
    试管编号字母 = mstr试管编号字母
End Property

'功能：设置本体检表要使用的试管编号字母，同时写入数据库中。
'作者：邓恒
Public Property Let 试管编号字母(ByVal para试管编号字母 As String)
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    If para试管编号字母 <> "" Then
        '开始事务。
        dasubBeginTran
        
        '判断本类体检表当天是否已登记使用了别的字母。
        Set lobjRec = dafuncGetData("select 体检表名称,字母,是否固定 from 体检管理_试管编号字母表 where 体检表名称='" & mstr体检表名 & "'  and convert(varchar(10),日期,101)=convert(varchar(10),getdate(),101)")
        If lobjRec.recordcount > 0 Then
            If IIf(IsNull(lobjRec!是否固定), 0, lobjRec!是否固定) = 0 Then
                '已有记录，但没有固定，可以修改。
                If lobjRec!字母 <> para试管编号字母 Then
                    '修改了字母。
                    dafuncGetData "update 体检管理_试管编号字母表 set 字母='" & para试管编号字母 & "' where 体检表名称='" & mstr体检表名 & "'"
                End If
            Else
                '已有记录，并且已固定，不允许修改。
                If para试管编号字母 <> lobjRec!字母 Then
                    mstr试管编号字母 = lobjRec!字母
                    '返回错误给界面，不允许重新选择字母。
                    Err.Raise 6666, , "该类型体检表今天已被固定使用字母“" & lobjRec!字母 & "”了，不允许再修改试管编号字母。"
                End If
            End If
            
        Else
            '没有当天记录，删除该体检表的所有记录。并插入记录(缺省，不固定)。
            dafuncGetData "insert into 体检管理_试管编号字母表(体检表名称,字母,日期,是否固定) values('" & mstr体检表名 & "','" & para试管编号字母 & "',convert(varchar(10),getdate(),101),0)"
        End If
        
        '提交事务。
        dasubCommitTran
    End If
    
    mstr试管编号字母 = para试管编号字母
    
    Exit Property
errHandler:
    Dim lstrError As String
    If Err.Number = -2147217873 Then
        lstrError = "系统无法继续当前处理，因为：" & Chr(13) & Chr(10) & "当前你正在使用的体检表已被人删除了。"
    Else
        lstrError = func错误处理(Err.Number, Err.Description)
    End If
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Property Let 试管编号字母", 6666, lstrError, True
    Exit Property
    Resume
End Property

'功能：获取当前体检表所使用的体检单名称。
'作者：杨春
Public Property Get 体检单名称() As String
    On Error GoTo errHandler
    
    Dim lobj体检表模板 As clsMedicalExamTemplate
    
    '创建体检表模板对象，获取体检单名称。
    If mstr体检单名称 = "" Then
        Set lobj体检表模板 = New clsMedicalExamTemplate
        lobj体检表模板.体检表名 = mstr体检表名
        mstr体检单名称 = lobj体检表模板.体检单名称
    End If
    
    体检单名称 = mstr体检单名称
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Property Get 体检单名称", 6666, lstrError, True
End Property


'功能：获取“附加信息”。
'返回：collection[附加项目，附加值]key:附加项目。
'作者：邓恒
Public Property Get 附加信息() As Collection
Attribute 附加信息.VB_Description = "'功能：获取“附加信息”。\r\n'返回：collection[附加项目，附加值]key:附加项目。\r\n"
    Dim lcolTemp As New Collection
    Dim lcolItem As Variant
    
    On Error GoTo errHandler
    
    '若还未获取过附加项目，获取附加项目。
    If Not mbln是否已获取附加项目 Then
        sub获取附加项目
    End If
    
    For Each lcolItem In mcol附加信息
        lcolTemp.Add lcolItem, lcolItem("附加项目")
    Next
    Set 附加信息 = lcolTemp
    
    Set lcolTemp = Nothing
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Property Get 附加信息", 6666, lstrError, True
End Property

'功能：获取本体检表的指定属性（或所有）的体检项目。
'输入：para属性 as string   '常规/化验。若该参数为空，表示获取所有体检项目。
'返回：collection[clsFactTestItem]key:体检项目编码。
'作者：邓恒
Public Property Get 体检项目集(ByVal para属性 As String) As Collection
Attribute 体检项目集.VB_Description = "'功能：获取本体检表的指定属性（或所有）的体检项目。\r\n    '输入：para属性 as string   '常规/化验。若该参数为空，表示获取所有体检项目。\r\n    '返回：collection[clsFactTestItem]key:体检项目编码。\r\n"
    Dim lcol体检项目集 As New Collection    '要范围的体检项目集合。
    Dim lobj体检项目 As Variant            '集合中的某个元素。
   
    On Error GoTo errHandler
    
    '若还未获取过体检项目，获取体检项目。
    If Not mbln是否已获取体检项目 Then
        sub获取体检项目
    End If
    
    Set 体检项目集 = New Collection
    For Each lobj体检项目 In mcol体检项目
        '判断属性是否符合参数要求。
        If para属性 = "" Or lobj体检项目.属性 = para属性 Then
            '符合，则加入返回集合中。
            lcol体检项目集.Add lobj体检项目, lobj体检项目.体检项目编号
        End If
    Next
    Set 体检项目集 = lcol体检项目集
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Property Get 体检项目集", 6666, lstrError, True
End Property

'功能：获取本体检表的指定属性（或所有）的体检项目。
'输入：para属性 as string   '常规/化验。若该参数为空，表示获取所有体检项目。
'返回：Recordset(体检项目,体检项目名称,体检结果)。
'作者：杨春
'日期：2001-4-22。
Public Property Get 优化体检项目集(ByVal para属性 As String) As Object
    On Error GoTo errHandler
    Set 优化体检项目集 = dafuncGetData("select 体检项目,体检项目名称,体检结果,标准值,单位,单项结论 from 体检管理_体检结果数据库 where 系统编号='" & mstr系统编号 & "'" & IIf(para属性 = "", "", " and 属性='" & para属性 & "'"))

    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Property Get 优化体检项目集", 6666, lstrError, True
    Exit Property
    Resume
End Property
'功能：设置基本信息中的附加项目的值，但不立刻写库。
'输入：para附加项目  附加项目的名称。
'      para项目值    项目值：编号+“ ”+名称。
'作者：邓恒
Public Sub Sub填附加信息值(ByVal para附加项目 As String, ByVal para项目值 As String)
Attribute Sub填附加信息值.VB_Description = "'功能：设置基本信息中的附加项目的值，但不立刻写库。\r\n    '输入：para附加项目  附加项目的名称。\r\n    '      para项目值    项目值：编号+“ ”+名称。\r\n"
    Dim lcolTemp As Collection  '要加入mcol附加信息中的元素。
    Dim lstrCode As String
    Dim lstrName As String
    
    On Error GoTo errHandler
    '若还未获取过附加项目，获取附加项目。
    If Not mbln是否已获取附加项目 Then
        sub获取附加项目
    End If
    
    '分解项目值中的编码和名称。
    If InStr(para项目值, " ") > 0 Then
        lstrCode = Left(para项目值, InStr(para项目值, " ") - 1)
        lstrName = Right(para项目值, Len(para项目值) - InStr(para项目值, " "))
    Else
        lstrName = para项目值
        lstrCode = ""
    End If
    '元素已存在，修改。
    If sffunc判断集合键值是否存在(mcol附加信息, para附加项目) Then
        Set lcolTemp = mcol附加信息(para附加项目)
        lcolTemp.Remove "项目值"
        lcolTemp.Remove "项目值编号"
        
        lcolTemp.Add lstrName, "项目值"
        lcolTemp.Add lstrCode, "项目值编号"
    Else
        Set lcolTemp = New Collection
        lcolTemp.Add para附加项目, "附加项目"
        lcolTemp.Add lstrName, "项目值"
        lcolTemp.Add lstrCode, "项目值编号"
        mcol附加信息.Add lcolTemp, para附加项目
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Sub填附加信息值", 6666, lstrError, True
End Sub

'功能：设置指定体检项目的体检结果，不立刻写库。
'作者：邓恒
Public Sub Sub填体检结果(ByVal Para体检项目 As String, ByVal para体检结果 As String, ByVal para医师编号 As String, ByVal para填写日期 As Date)
    Dim lobjItem As ClsFactTestItem 'mcol体检项目中药修改的元素。
        
    On Error GoTo errHandler
    
    '若还未获取过体检项目，获取体检项目。
    If Not mbln是否已获取体检项目 Then
        sub获取体检项目
    End If
    
    '获取要修改的体检项目。
    If sffunc判断集合键值是否存在(mcol体检项目, Para体检项目) Then
        Set lobjItem = mcol体检项目(Para体检项目)
        '修改体检项目。
        With lobjItem
            .体检结果 = para体检结果
            .医师编号 = para医师编号
            .填写日期 = para填写日期
        End With
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Sub填体检结果", 6666, lstrError, True
End Sub

'功能：保存该体检表上所有体检结果到库中。
'作者：邓恒
Public Sub Sub保存体检结果()
Attribute Sub保存体检结果.VB_Description = "'功能：保存该体检表上所有体检结果到库中。\r\n"
    Dim lobjItem As ClsFactTestItem 'mcol体检项目中的元素。
    Dim lstrDeleted As String
    Dim lstr系统编号 As String
    
    On Error GoTo errHandler
    '若还未获取过体检项目，获取体检项目。
    If Not mbln是否已获取体检项目 Then
        sub获取体检项目
    End If
    '依次保存所有项目的体检结果。
    lstrDeleted = ""
    For Each lobjItem In mcol体检项目
        lstrDeleted = lstrDeleted + "'" + lobjItem.体检项目编号 + "',"
        lstr系统编号 = lobjItem.系统编号
        lobjItem.SubSave
        
    Next
    
    '删除所有不在当前体检项目集中的项目。
    If lstrDeleted = "" Then
        '删除所有体检项目。
    Else
        lstrDeleted = Left(lstrDeleted, Len(lstrDeleted) - 1)
        dafuncGetData "delete from 体检管理_体检结果信息表 where 系统编号='" & lstr系统编号 & "' and 体检项目 not in (" & lstrDeleted & ")"
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "Sub保存体检结果", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：保存该体检表上所有体检结果到库中。
'返回：系统得出的体检结论串。
'作者：邓恒
Public Function func下体检结论() As String
Attribute func下体检结论.VB_Description = "'功能：保存该体检表上所有体检结果到库中。\r\n    '返回：系统得出的体检结论串。\r\n"
    Dim lobj体检表模板 As clsMedicalExamTemplate '体检表模板对象。
    Dim lobj体检结论 As ClsMedicalExamConclusion '体检结论对象。
    Dim lcol体检结论 As Collection               '体检表模板的属性“体检结论集”。
    Dim lstr体检结论  As String                  '最后应返回的体检结论串。
    Dim i As Integer
    
    On Error GoTo errHandler
    func下体检结论 = ""
    
    '创建体检表模板对象。
    Set lobj体检表模板 = New clsMedicalExamTemplate
    lobj体检表模板.体检表名 = mstr体检表名
    
    '获取该体检表模板的所有可选的体检结论。
    Set lcol体检结论 = lobj体检表模板.体检结论集
    
    '依次判断是否可以得出该体检表模板的各体检结论，并拼在一起。
    lstr体检结论 = ""
    For i = 1 To lcol体检结论.Count
        '创建体检结论对象。
        Set lobj体检结论 = New ClsMedicalExamConclusion
        lobj体检结论.ID = lcol体检结论(i)("体检结论ID")
        
        '判断该结论是否可用。
        If lobj体检结论.Func判断是否可下本结论(mstr系统编号) = True Then
            '可用，则拼在串中。
            lstr体检结论 = lstr体检结论 & lcol体检结论(i)("名称") & ","
        End If
    Next i
    If lstr体检结论 <> "" Then
        func下体检结论 = Left(lstr体检结论, Len(lstr体检结论) - 1)
    Else
        func下体检结论 = "正常"
    End If
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "func下体检结论", 6666, lstrError, True
End Function

'功能：对当前对象的属性作合法性判断。
'输出：paraErrorInfo   若不合法时，该参数表明不合法的具体错误信息。
'返回：检查结果  true合法/false不合法。
'作者：邓恒
Public Function func合法性检查(paraErrorInfo As String) As Boolean
Attribute func合法性检查.VB_Description = "'功能：对当前对象的属性作合法性判断。\r\n    '输出：paraErrorInfo   若不合法时，该参数表明不合法的具体错误信息。\r\n    '返回：检查结果  true合法/false不合法。\r\n"
    func合法性检查 = True
    If mstr体检表名 = "" Then
        paraErrorInfo = "必须选择体检表模板。"
        func合法性检查 = False
        Exit Function
    End If
    
    Dim lobj业务对象 As Object
    Set lobj业务对象 = New clsManageMedicalExam
    If lobj业务对象.业务设置("试管编号自动生成") = "是" Then
        If mstr试管编号字母 = "" Then
            paraErrorInfo = "必须选择试管字母。"
            func合法性检查 = False
            Exit Function
        End If
    End If
    
    '若还未获取过体检项目，获取体检项目。
    If Not mbln是否已获取体检项目 Then
        sub获取体检项目
    End If
    If mcol体检项目.Count = 0 Then
        paraErrorInfo = "该次体检选择的体检表模板上没有体检项目。请进入“业务设置”的“体检表模板设置”操作界面，设置该类体检表应作的体检项目。"
        func合法性检查 = False
        Exit Function
    End If
End Function
Private Sub Class_Initialize()
    Set mcol附加信息 = New Collection
    Set mcol体检项目 = New Collection
    mbln是否已存在 = False
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mcol附加信息 = Nothing
    Set mcol体检项目 = Nothing
End Sub

'功能：获取当前体检表上的所有附加项目及其结果。
'作者：杨春
Private Sub sub获取附加项目()
    Dim lobj体检表模板 As clsMedicalExamTemplate
    Dim lobjRec As Object '执行sql语句结果。
    Dim lcolItem As Collection
    Dim lcolInfo As Collection
    Dim i As Long
    
    On Error GoTo errHandler
    Set mcol附加信息 = New Collection
    
    '获取附加项目,项目值。
    Set lobjRec = dafuncGetData("select 附加项目,项目值,项目值编号 from 体检管理_体检附加信息表 where 系统编号='" & mstr系统编号 & "'")
    If lobjRec.recordcount > 0 Then
        lobjRec.movefirst
        Do While Not lobjRec.EOF
            Set lcolItem = New Collection
            lcolItem.Add lobjRec("附加项目").Value, "附加项目"
            lcolItem.Add lobjRec("项目值").Value, "项目值"
            lcolItem.Add IIf(IsNull(lobjRec("项目值编号")), "", lobjRec("项目值编号").Value), "项目值编号"
            
            mcol附加信息.Add lcolItem, lcolItem("附加项目")
            
            lobjRec.MoveNext
        Loop
    ElseIf mstr体检表名 <> "" Then
    
        '创建体检表模板对象。
        Set lobj体检表模板 = New clsMedicalExamTemplate
        lobj体检表模板.体检表名 = mstr体检表名
        Set lcolInfo = lobj体检表模板.基本附加项目集
        For i = 1 To lcolInfo.Count
            Set lcolItem = New Collection
            lcolItem.Add lcolInfo(i)("附加项目"), "附加项目"
            lcolItem.Add "", "项目值"
            lcolItem.Add "", "项目值编号"
            mcol附加信息.Add lcolItem, lcolItem("附加项目")
        Next
    End If
    lobjRec.Close
    mbln是否已获取附加项目 = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "sub获取附加项目", 6666, lstrError, True
End Sub

'功能：获取当前体检表上的所有体检项目及其结果。
'作者：杨春
Private Sub sub获取体检项目()
    Dim lobj体检表模板 As clsMedicalExamTemplate
    Dim lobjFactItem As ClsFactTestItem
    Dim lobjItem As Variant              'clsTestItem。
    Dim lcolInfo As Collection
    Dim lobjRec As Object                '执行sql语句结果。
    
    On Error GoTo errHandler
    Set mcol体检项目 = New Collection
    
    '获取体检项目集。
    Set lobjRec = dafuncGetData("select 体检项目 from 体检管理_体检结果信息表 where 系统编号='" & mstr系统编号 & "'")
    If lobjRec.recordcount > 0 Then
        lobjRec.movefirst
        Do While Not lobjRec.EOF
            '创建体检项目对象。
            Set lobjFactItem = New ClsFactTestItem
            lobjFactItem.系统编号 = mstr系统编号
            lobjFactItem.体检项目编号 = lobjRec("体检项目")
            
            mcol体检项目.Add lobjFactItem, lobjRec("体检项目").Value
            
            lobjRec.MoveNext
        Loop
    ElseIf mstr体检表名 <> "" Then
    
        '创建体检表模板对象。
        Set lobj体检表模板 = New clsMedicalExamTemplate
        lobj体检表模板.体检表名 = mstr体检表名
        
        '获取该体检表模板的所有体检项目：[ClsTestItem]，并加入mcol体检项目。
        Set lcolInfo = lobj体检表模板.体检项目集
        For Each lobjItem In lcolInfo
            Set lobjFactItem = New ClsFactTestItem
            lobjFactItem.系统编号 = mstr系统编号
            lobjFactItem.体检项目编号 = lobjItem.编码
            
            mcol体检项目.Add lobjFactItem, lobjFactItem.体检项目编号
        Next
    End If
    
    lobjRec.Close
    mbln是否已获取体检项目 = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象", "ClsMedicalExamSheet", "sub获取体检项目", 6666, lstrError, True
End Sub

