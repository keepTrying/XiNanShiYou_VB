VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsConclusionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mlng编号 As Long
Private mlngID As Long
Private mstr描述 As String
Private mstr体检结论名称 As String

Private mcol判断条件 As Collection '[序号，体检项目，体检项目名称，判断条件，标准值]key:序号。

Private mbln是否已存在 As Boolean  '表明当前分组的条件是否已在数据库中存在。

Private Sub Class_Initialize()
    On Error GoTo errHandler
    subClear
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "Class_Initialize", Err.Number, Err.Description, False
End Sub

'功能：表明当前分组的条件是否已在数据库中存在。
'作者：杨春
Public Property Get 是否已存在() As Boolean
Attribute 是否已存在.VB_Description = "'功能：表明当前分组的条件是否已在数据库中存在。"
    是否已存在 = mbln是否已存在
End Property

Property Get 编号() As Long
    编号 = mlng编号
End Property
'功能：给本对象的关键属性赋值。
'作者：杨春
Property Let 编号(ByVal vNewValue As Long)
    Dim lobjDict As Object
    On Error GoTo errHandler
    
    mlng编号 = vNewValue
    mbln是否已存在 = False
    Set mcol判断条件 = New Collection
    
    '获取对象属性。
    If mlngID <> 0 And mlng编号 <> 0 Then
        subGetProperty
    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "Property Let 编号", 6666, lstrError, True
End Property
Property Get ID() As Long
    ID = mlngID
End Property
'功能：给本对象的关键属性赋值。
'作者：杨春
Property Let ID(ByVal vNewValue As Long)
    On Error GoTo errHandler
    
    mlngID = vNewValue
    mbln是否已存在 = False
    mstr体检结论名称 = ""
    Set mcol判断条件 = New Collection
    
    '获取对象属性。
    If mlngID <> 0 And mlng编号 <> 0 Then
        subGetProperty
    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "Property Let ID", 6666, lstrError, True
End Property

Property Get 描述() As String
    描述 = mstr描述
End Property
Property Let 描述(ByVal vNewValue As String)
    mstr描述 = vNewValue
End Property

Public Property Get 体检结论名称() As String
    Dim lobjRec As Object   '字典内容。
    
    On Error GoTo errHandler
    
    '用字典对象，根据id获取体检结论名称。
    If mlngID <> 0 And mstr体检结论名称 = "" Then
        If pobjDict Is Nothing Then
            Err.Raise 6666, , "无法创建“字典管理”的对象。请退出本系统，重新注册“字典管理.dll”。"
        Else
            Set lobjRec = pobjDict.Fetch("体检结论字典", "InnerID=" & mlngID)
            If lobjRec.recordcount > 0 Then
                mstr体检结论名称 = lobjRec("名称")
            Else
                Err.Raise 6666, , "体检结论字典中没有InnerID=" & mlngID & "的字典项。"
            End If
            lobjRec.Close
        End If
    End If
    
    体检结论名称 = mstr体检结论名称
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "Property Get 体检结论名称", 6666, lstrError, True
End Property

'功能：返回本对象可以识别的判断条件:=,<,>,<=,>=,<>,in,b,lb,rb,lrb。
'返回：collection[符号，说明],key:符号。
'作者：杨春
Public Property Get 判断条件枚举() As Collection
    Dim lcolItem As Collection
    
    On Error GoTo errHandler
    Set 判断条件枚举 = New Collection
    
    With 判断条件枚举
        Set lcolItem = New Collection
        lcolItem.Add "=", "符号"
        lcolItem.Add "等于", "说明"
        .Add lcolItem, lcolItem("符号")
        
        Set lcolItem = New Collection
        lcolItem.Add "<", "符号"
        lcolItem.Add "小于", "说明"
        .Add lcolItem, lcolItem("符号")
        
        Set lcolItem = New Collection
        lcolItem.Add ">", "符号"
        lcolItem.Add "大于", "说明"
        .Add lcolItem, lcolItem("符号")
        
        Set lcolItem = New Collection
        lcolItem.Add ">=", "符号"
        lcolItem.Add "大于等于", "说明"
        .Add lcolItem, lcolItem("符号")
        
        Set lcolItem = New Collection
        lcolItem.Add "<=", "符号"
        lcolItem.Add "小于等于", "说明"
        .Add lcolItem, lcolItem("符号")
        
        Set lcolItem = New Collection
        lcolItem.Add "<>", "符号"
        lcolItem.Add "不等于", "说明"
        .Add lcolItem, lcolItem("符号")
        
        Set lcolItem = New Collection
        lcolItem.Add "in", "符号"
        lcolItem.Add "包含于", "说明"
        .Add lcolItem, lcolItem("符号")
    
'        Set lcolItem = New Collection
'        lcolItem.Add "B", "符号"
'        lcolItem.Add "大于开始值，并且小于结束值", "说明"
'        .Add lcolItem, lcolItem("符号")
'
'        Set lcolItem = New Collection
'        lcolItem.Add "LB", "符号"
'        lcolItem.Add "大于等于开始值，并且小于结束值", "说明"
'        .Add lcolItem, lcolItem("符号")
'
'        Set lcolItem = New Collection
'        lcolItem.Add "RB", "符号"
'        lcolItem.Add "大于开始值，并且小于等于结束值", "说明"
'        .Add lcolItem, lcolItem("符号")
'
'        Set lcolItem = New Collection
'        lcolItem.Add "LRB", "符号"
'        lcolItem.Add "大于等于开始值，并且小于等于结束值", "说明"
'        .Add lcolItem, lcolItem("符号")
    End With
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "Property Get 判断条件枚举", Err.Number, Err.Description, True
End Property

'功能：获取当前分组的所有判断条件。
'返回：[序号，体检项目，判断条件，标准值]key:体检项目。
'作者：杨春
Public Property Get 判定条件() As Collection
Attribute 判定条件.VB_Description = "'功能：获取当前分组的所有判断条件。\r\n    '返回：[序号，体检项目，判断条件，标准值]key:体检项目。\r\n"
    Set 判定条件 = mcol判断条件
End Property

'功能：添加子条件，不立刻写库。
'作者：杨春
Public Sub SubAddFilter(ByVal para序号 As Integer, ByVal Para体检项目 As String, ByVal para体检项目名称 As String, ByVal para判断条件 As String, ByVal para标准值 As String)
Attribute SubAddFilter.VB_Description = "'功能：添加子条件，不立刻写库。\r\n"
    Dim colItem As Collection
    On Error GoTo errHandler
    
    '先删除旧的。
    If sffunc判断集合键值是否存在(mcol判断条件, Trim(Str(para序号))) Then
        mcol判断条件.Remove Trim(Str(para序号))
    End If
    
    '添加旧的。
    Set colItem = New Collection
    colItem.Add para序号, "序号"
    colItem.Add Para体检项目, "体检项目"
    colItem.Add para体检项目名称, "体检项目名称"
    colItem.Add para判断条件, "判断条件"
    colItem.Add para标准值, "标准值"
    
    mcol判断条件.Add colItem, Trim(Str(para序号))
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "SubAddFilter", 6666, lstrError, False
End Sub
'功能：删除子条件，不立刻写库。
Public Sub SubRemoveFilter(ByVal para序号 As Integer)
Attribute SubRemoveFilter.VB_Description = "'功能：删除子条件，不立刻写库。\r\n"
    On Error Resume Next
    mcol判断条件.Remove Trim(Str(para序号))
End Sub
'功能：删除所有子条件，不立刻写库。
Public Sub SubRemoveAllFilter()
Attribute SubRemoveAllFilter.VB_Description = "'功能：删除所有子条件，不立刻写库。\r\n"
    Set mcol判断条件 = New Collection
    
End Sub

'功能：保存本对象所包含的多有条件导数据库中。
'作者：杨春
Public Sub SubSave()
Attribute SubSave.VB_Description = "'功能：保存本对象所包含的多有条件导数据库中。\r\n"
    Dim lobjRec As Object
    Dim lcolItem As Collection
    Dim i As Long
    
    On Error GoTo errHandler
    '开始事务。
    dasubBeginTran
    
    '先删除旧记录。
    If mbln是否已存在 Then
        dafuncGetData "delete 体检管理_体检结论判断条件表 where 编号=" & mlng编号
    Else
        '生成新编号。
        Set lobjRec = dafuncGetData("select max(编号) from 体检管理_体检结论判断条件表")
        mlng编号 = IIf(IsNull(lobjRec(0)), 0, lobjRec(0)) + 1
    End If
    
    '添加新的纪录。
    For Each lcolItem In mcol判断条件
        dafuncGetData "insert into 体检管理_体检结论判断条件表(体检结论,编号,序号,体检项目,判断条件,标准值,描述) values(" & mlngID & "," & mlng编号 & "," & lcolItem("序号") & ",'" & lcolItem("体检项目") & "','" & lcolItem("判断条件") & "','" & lcolItem("标准值") & "','" & mstr描述 & "')"
    Next
    
    '结束事务。
    dasubCommitTran
    
    mbln是否已存在 = True
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "subSave", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'功能：删除库中本对象所包含的条件。
'作者：杨春
Public Sub subDelete()
Attribute subDelete.VB_Description = "'功能：删除库中本对象所包含的条件。\r\n"
    On Error GoTo errHandler
    '先删除旧记录。
    dafuncGetData "delete 体检管理_体检结论判断条件表 where 编号=" & mlng编号
    mbln是否已存在 = False
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "subDelete", 6666, lstrError, True
End Sub

'功能：判断指定系统编号所对应的体检记录是否满足当前分组的所有条件。
'输入：系统编号
'返回：true满足/false不满足。
'作者：杨春
Public Function Func判断是否满足条件(ByVal para系统编号 As String) As Boolean
Attribute Func判断是否满足条件.VB_Description = "'功能：判断指定系统编号所对应的体检记录是否满足当前分组的所有条件。\r\n"
    Dim lobj体检表 As Object       'ClsMedicalExamSheet
    Dim lcol体检项目 As Collection '指定系统编号的体检项目集合。
    Dim lcolItem As Collection     '判断条件中的一个子条件。
    Dim lstr体检结果 As String     '一个体检项目的体检结果。
    Dim lstr标准值 As String       '一个体检项目的标准值。
    Dim i As Long
    
    On Error GoTo errHandler
    If mcol判断条件.Count = 0 Then
        Func判断是否满足条件 = False
    Else
        Func判断是否满足条件 = True
    End If
    '创建“体检表对象”。
    Set lobj体检表 = New ClsMedicalExamSheet
    lobj体检表.系统编号 = para系统编号
    
    '依次判断条件中各值是否与体检表中对应体检项目的体检结果满足对应关系，
    '只要有一项不满足，则返回false，否则返回true。
    Set lcol体检项目 = lobj体检表.体检项目集("")
    For i = 1 To mcol判断条件.Count
        Set lcolItem = mcol判断条件(i)
        '首先判断当前子条件中的体检项目是否在体检表上。
        If Not sffunc判断集合键值是否存在(lcol体检项目, lcolItem("体检项目")) Then
            '不在，肯定不满足。
            Func判断是否满足条件 = False
            Exit For
        End If
        
        lstr体检结果 = lcol体检项目(lcolItem("体检项目")).体检结果
        
        '体检结果值中可能带有关系运算符号，应先去掉该符号。
        lstr体检结果 = Trim(lstr体检结果)
        If lstr体检结果 <> "" Then
            If Left(lstr体检结果, 1) = "=" Then
                lstr体检结果 = Trim(Right(lstr体检结果, Len(lstr体检结果) - 1))
            ElseIf Left(lstr体检结果, 1) = ">" Then
                lstr体检结果 = Trim(Right(lstr体检结果, Len(lstr体检结果) - 1))
                If Left(lstr体检结果, 1) = "=" Then
                    lstr体检结果 = Trim(Right(lstr体检结果, Len(lstr体检结果) - 1))
                Else

                End If
            End If
        End If
        
        lstr标准值 = lcolItem("标准值")
        Select Case UCase(lcolItem("判断条件"))
        Case "="
            Func判断是否满足条件 = (lstr体检结果 = lstr标准值)
        Case "<"
            If IsNumeric(lstr体检结果) Then
                Func判断是否满足条件 = (Val(lstr体检结果) < Val(lstr标准值))
            Else
                Func判断是否满足条件 = (lstr体检结果 < lstr标准值)
            End If
        Case ">"
            If IsNumeric(lstr体检结果) Then
                Func判断是否满足条件 = (Val(lstr体检结果) > Val(lstr标准值))
            Else
                Func判断是否满足条件 = (lstr体检结果 > lstr标准值)
            End If
        Case ">="
            If IsNumeric(lstr体检结果) Then
                Func判断是否满足条件 = (Val(lstr体检结果) >= Val(lstr标准值))
            Else
                Func判断是否满足条件 = (lstr体检结果 >= lstr标准值)
            End If
        Case "<="
            If IsNumeric(lstr体检结果) Then
                Func判断是否满足条件 = (Val(lstr体检结果) <= Val(lstr标准值))
            Else
                Func判断是否满足条件 = (lstr体检结果 <= lstr标准值)
            End If
        Case "<>"
            If IsNumeric(lstr体检结果) Then
                Func判断是否满足条件 = (Val(lstr体检结果) <> Val(lstr标准值))
            Else
                Func判断是否满足条件 = Trim(lstr体检结果) <> Trim(lstr标准值)
            End If
        Case "in"
            lstrTemp = lstr标准值
            If Right(lstrTemp, 1) <> "," Then lstrTemp = lstrTemp & ","
            If InStr(1, lstrTemp, lstr体检结果 & ",") = 0 Then
                Func判断是否满足条件 = False
            End If
            
        Case "B", "LB", "RB", "LRB"
            Dim lstrStart As String '标准值中的最小值。
            Dim lstrEnd As String   '标准值中的最大值。
            If InStr(lstr标准值, ",") = 0 Then
                Err.Raise 6666, , "判断条件为“B、LB、RB、LRB”的标准值不符合格式要求。格式必须是：开始值+“,”+结束值"
            Else
                lstrStart = Left(lstr标准值, InStr(lstr标准值, ",") - 1)
                lstrEnd = Right(lstr标准值, Len(lstr标准值) - InStr(lstr标准值, ","))
                Select Case UCase(lcolItem("判断条件"))
                Case "B"
                    Func判断是否满足条件 = (Val(lstr体检结果) > Val(lstrStart) And Val(lstr体检结果) < Val(lstrEnd))
                Case "LB"
                    Func判断是否满足条件 = (Val(lstr体检结果) >= Val(lstrStart) And Val(lstr体检结果) < Val(lstrEnd))
                Case "RB"
                    Func判断是否满足条件 = (Val(lstr体检结果) > Val(lstrStart) And Val(lstr体检结果) <= Val(lstrEnd))
                Case "LRB"
                    Func判断是否满足条件 = (Val(lstr体检结果) >= Val(lstrStart) And Val(lstr体检结果) <= Val(lstrEnd))
                End Select
            End If
        Case Else
            
        End Select
        
        If Not Func判断是否满足条件 Then
            Exit For
        End If
    Next
    
    Exit Function
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    Func判断是否满足条件 = False
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "Func判断是否满足条件", 6666, lstrError, True
    Exit Function
    Resume
End Function
'功能：清空本对象的属性。
'作者：杨春
Public Sub subClear()
    On Error Resume Next
    mlng编号 = 0
    mlngID = 0
    Set mcol判断条件 = New Collection
    mbln是否已存在 = False
    
End Sub

'功能：根据对象唯一属性获取对象其他属性。
'作者：杨春
Private Sub subGetProperty()
    Dim lobjRec As Object
    Dim lcolItem As Collection
    
    On Error GoTo errHandler
    Set mcol判断条件 = New Collection
    
    '从库中获取当前编号的所有子条件。
    Set lobjRec = dafuncGetData("select * from 体检管理_体检结论判断条件数据库 where 体检结论ID=" & mlngID & " and 编号=" & mlng编号)
    If lobjRec.recordcount > 0 Then
        mstr描述 = IIf(IsNull(lobjRec!描述), "", lobjRec!描述)
        mstr体检结论名称 = IIf(IsNull(lobjRec!体检结论名称), "", lobjRec!体检结论名称)
        
        '把所有分组添加到集合中:[序号，体检项目，体检项目名称，判断条件，标准值]key:序号。
        Do While Not lobjRec.EOF
            Set lcolItem = New Collection
            With lcolItem
                .Add lobjRec("序号").Value, "序号"
                .Add lobjRec("体检项目").Value, "体检项目"
                .Add lobjRec("体检项目名称").Value, "体检项目名称"
                .Add lobjRec("判断条件").Value, "判断条件"
                .Add lobjRec("标准值").Value, "标准值"
            End With
            mcol判断条件.Add lcolItem, Trim(Str(lobjRec("序号")))
            
            lobjRec.MoveNext
        Loop
        mbln是否已存在 = True
    Else
        mstr描述 = ""
        mstr体检结论名称 = ""
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "ClsConclusionFilter", "subGetProperty", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mcol判断条件 = Nothing
End Sub
