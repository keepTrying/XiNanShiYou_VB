VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInputGridServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mctrlInput As Control             'dyInputGrid,dyInputPad 录入网格控件或录入模板。
Private mInputTemplate As clsDyInputItems '道源录入模板。
Private mstrName As String                '录入名称。
Private mstrTableName As String           '表名称。
Private mstrCritical As String            '录入条件。
Private mstrPriKey As String              '主键列表，以chr(1)割开。

Private mcolPriKey As Collection          '主键集合。
Private mcolFieldNames As Collection      '字段名集合，以显示名为键。
Private mcolFieldTypes As Collection      '字段类型，以显示名为键。
Private mcolEnumMethod As Collection      '枚举值方式集合，以显示名为键。
Private mcolInnerNames As Collection      '显示名集合，以内部名称为键。

Private mblnAutoIncre As Boolean          '主键是否自增长。
Private mcolFieldPos As Collection
Private mDbInterfaceObj As Object

Private mblnAllowDelete As Boolean        '是否允许删除。

Private Sub Class_Initialize()
    Set mcolPriKey = New Collection
    Set mcolFieldNames = New Collection
    Set mcolFieldTypes = New Collection
    Set mcolEnumMethod = New Collection
    Set mcolInnerNames = New Collection
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    
    Set mDbInterfaceObj = Nothing
    Set mctrlInput = Nothing
    Set mInputTemplate = Nothing

    
End Sub
'功能：获取某表的录入模板。
'输入：paraTableName 表名。
'输出：mcolFieldNames    字段名集合，以显示名为键。
'      mcolEnumMethod    枚举值方式集合，以显示名为键。
'      mcolInnerNames    各显示名的内部名，以显示名为键。
'      mcolFieldTypes    各字段类型，以显示名为键。
Private Sub sub_GetInputTemplate(paraInputTemplate As clsDyInputItems, ByVal paraTableName As String)
Attribute sub_GetInputTemplate.VB_Description = "获取某表的录入模板：clsDyInputItems对象。"
    Dim recTmp As Object
    Dim strEnum As String
    Dim recDict As Object
    Dim intType As enumDyInputType
    Dim strTmp As String
    
    On Error GoTo err_handler
    
    Set mcolFieldNames = New Collection
    Set mcolInnerNames = New Collection
    Set mcolEnumMethod = New Collection
    Set mcolFieldTypes = New Collection
    
    '获取模板信息。
    mDbInterfaceObj.subExecuteSql "select * from 系统_录入模板信息表 where 表名='" & paraTableName & "' order by 序号", recTmp
    With paraInputTemplate
        Do While Not recTmp.EOF
            strEnum = IIf(IsNull(recTmp!枚举值), "", recTmp!枚举值)
            intType = IIf(IsNull(recTmp!类型), 0, recTmp!类型)
            mcolFieldTypes.Add intType, recTmp!显示名
            If strEnum <> "" And InStr(strEnum, ",") = 0 Then
                '是枚举字典，从字典中获取枚举值。
                If InStr(strEnum, "[") > 0 Then
                    mDbInterfaceObj.subExecuteSql "select " & Mid(strEnum, InStr(strEnum, "[") + 1, InStr(strEnum, "]") - InStr(strEnum, "[") - 1) & " from " & Left(strEnum, InStr(strEnum, "[") - 1), recDict
                Else
                    mDbInterfaceObj.subExecuteSql "select 编码+' '+名称 from " & strEnum, recDict
                End If
                strEnum = ""
                If recDict.recordcount > 0 Then
                    recDict.MoveFirst
                    Do While Not recDict.EOF
                        strEnum = strEnum & recDict(0) & ","
                        recDict.movenext
                    Loop
                End If
            End If
            If strEnum <> "" Then
                If intType <> dyInputSingleselecttext And intType <> dyInputMultiselecttext Then
                    intType = dyInputVartext
                End If
            End If
            .subAddItem recTmp!显示名, intType, recTmp!必须有, recTmp!允许录入, IIf(IsNull(recTmp!长度), 10, recTmp!长度), IIf(IsNull(recTmp!小数点后位数), 0, recTmp!小数点后位数), strEnum, recTmp!最大值, recTmp!最小值, recTmp!缺省值
            mcolFieldNames.Add recTmp("字段名").Value, recTmp!显示名
            mcolEnumMethod.Add IIf(IsNull(recTmp!枚举值方式), CODE_NAME, recTmp("枚举值方式").Value), recTmp!显示名
            
            '该录入模板没有“内部名称”字段，以显示名作为内部名称。
            On Error Resume Next
            strTmp = IIf(IsNull(recTmp!内部名称), "", recTmp!内部名称)
            If strTmp = "" Or Err <> 0 Then
                strTmp = recTmp("显示名").Value
            End If
            mcolInnerNames.Add recTmp("显示名").Value, strTmp
            Err.Clear
            
            On Error GoTo err_handler
            recTmp.movenext
        Loop
    End With
    Exit Sub
    
err_handler:
    Err.Raise Err, , "通用业务库.clsInputServer.sub_GetInputTemplate：获取录入模板信息失败！" & Error
End Sub
Public Property Get InputControl() As Variant
Attribute InputControl.VB_Description = "录入网格控件。"
    Set InputControl = mctrlInput
End Property

Public Property Set InputControl(ByVal vNewValue As Variant)
    Set mctrlInput = vNewValue
End Property

Public Property Get Name() As Variant
Attribute Name.VB_Description = "录入名称（表名）"
    Name = mstrName
End Property

'设置录入模板名。
Public Property Let Name(ByVal vNewValue As Variant)
    Dim i As Integer
    On Error GoTo err_handler
    
    mstrName = vNewValue
    Set mcolFieldPos = New Collection
    mctrlInput.InputTemplate.subRemoveAllItem
    '获取录入模板,和字段名集合。
    sub_GetInputTemplate mctrlInput.InputTemplate, mstrName
    
    '初始化录入网格。
    For i = 1 To mctrlInput.InputTemplate.Count
        mcolFieldPos.Add i - 1, mctrlInput.InputTemplate(i).Name
    Next i
    mctrlInput.subDraw
    Set mInputTemplate = mctrlInput.InputTemplate
    Exit Property
err_handler:
    Err.Raise Err, , "通用业务库.clsInputServer.PropertyLet Name：根据模板名设置录入模板失败，" & Error
End Property
'功能：把object内容加入录入网格中。
'输入：paraRec 记录集，
'      paraAdd 是否追加。
'      paraType  0 记录中字段是显示名/1 字段名/2 内部名称。
Public Sub subLoadGridFromRec(ByVal paraRec As Object, _
                            Optional paraAdd As Boolean = False, _
                            Optional paraType As Integer = 0)
    Dim i As Integer
    Dim colRow As Collection
    Dim strFieldName As String
    Dim strShowName As String
    On Error GoTo err_handler
    
    If Not paraAdd Then
        mctrlInput.subClear
    End If
    If paraRec.recordcount > 0 Then
        paraRec.MoveFirst
        Do While Not paraRec.EOF
            Set colRow = New Collection
            Select Case paraType
            Case 0 '记录中字段是显示名
                For i = 0 To paraRec.Fields.Count - 1
                    colRow.Add paraRec(i).Value, paraRec(i).Name
                Next i
            Case 1 '字段名。
                For i = 1 To mctrlInput.InputTemplate.Count
                    '根据显示名获取字段名。
                    On Error Resume Next
                    strFieldName = FieldName(mctrlInput.InputTemplate(i).Name)
                    If Err <> 0 Then
                        colRow.Add paraRec(strFieldName).Value, mctrlInput.InputTemplate(i).Name
                    End If
                    Err.Clear
                    On Error GoTo err_handler
                Next i
                
            Case 2 '内部名。
                For i = 0 To paraRec.Fields.Count - 1
                    '根据内部名获取显示名。
                    On Error Resume Next
                    strShowName = ShowName(paraRec(i).Name)
                    If Err <> 0 Then
                        colRow.Add paraRec(i).Value, strShowName
                    End If
                    Err.Clear
                    On Error GoTo err_handler
                Next i
            End Select
            
            mctrlInput.subAddRow colRow
            paraRec.movenext
        Loop
    End If
    Exit Sub
err_handler:
    '忽略错误。
End Sub
'功能：把object内容加入录入板中。
'输入：paraRec 记录集，
'      paraType  0 记录中字段是显示名/1 字段名/2 内部名称。
Public Sub subLoadPadFromRec(ByVal paraRec As Object, _
                            Optional paraType As Integer = 0)
    Dim i As Long
    Dim strFieldName As String '记录字段名。
    
    On Error GoTo err_handler
    
    With mctrlInput
        .subClear
        If Not paraRec.EOF Then
            Select Case paraType
            Case 0 '记录中字段名是显示名。
                For i = 1 To .InputTemplate.Count
                    On Error Resume Next
                    '获取显示名。
                    strFieldName = .InputTemplate(i).Name
                    .Value(i) = Null
                    '根据字段名获取字段值。
                    If Not IsNull(paraRec(strFieldName)) Then
                        If .InputTemplate(i).InputType = 0 And paraRec(strFieldName).Type = 3 Then
                            .Value(i) = Trim(Str(paraRec(strFieldName).Value))
                        Else
                            .Value(i) = paraRec(strFieldName).Value
                        End If
                    End If
                    Err.Clear
                    On Error GoTo err_handler
                Next i
            Case 1 '记录中字段名是字段名。
                For i = 1 To .InputTemplate.Count
                    On Error Resume Next
                    .Value(.InputTemplate(i).Name) = Null
                    '根据显示名获取字段名。
                    strFieldName = FieldName(.InputTemplate(i).Name)
                    If Err = 0 Then
                        '根据字段名获取字段值。
                        If Not IsNull(paraRec(strFieldName)) Then
                            If .InputTemplate(i).InputType = 0 And paraRec(strFieldName).Type = 3 Then
                                .Value(i) = Trim(Str(paraRec(strFieldName).Value))
                            Else
                                .Value(i) = paraRec(strFieldName).Value
                            End If
                        End If
                    End If
                    Err.Clear
                    On Error GoTo err_handler
                Next i
            Case 2 '记录中字段名是内部名。
                For i = 0 To paraRec.Fields.Count - 1
                    On Error Resume Next
                    '根据内部名获取显示名。
                    strFieldName = ShowName(paraRec.Fields(i).Name)
                    .Value(strFieldName) = Null
                    '根据字段名获取字段值。
                    If Not IsNull(paraRec(i)) Then
                        If .InputTemplate(strFieldName).InputType = 0 And paraRec(i).Type = 3 Then
                            .Value(strFieldName) = Trim(Str(paraRec(i).Value))
                        Else
                            .Value(strFieldName) = paraRec(i).Value
                        End If
                    End If
                    Err.Clear
                    On Error GoTo err_handler
                Next
            End Select
            
        End If
    End With
    Exit Sub
    
err_handler:
    '忽略错误。
End Sub
'功能：获取InputControl中当前新增记录。
'输入：paraValue  新增行值(或录入网格行号)，对录入板可缺省该参数。
'返回：修改信息："字段列表","值列表","值集合"
Public Function funcGetNewInfo(Optional paraValue) As Collection
Attribute funcGetNewInfo.VB_Description = "获取录入网格中新增行的信息：""fields"",""values""。"
    Dim strFields As String       '字段名列表。
    Dim strValues As String       '值列表。
    Dim colValues As New Collection '值集合。
    Dim colValue As Collection
    Dim i As Long
    Dim varValue                  '某字段值。
    Dim strShowName As String     '字段显示名。
    
    On Error GoTo err_handler
    
    If Not IsMissing(paraValue) Then
        If Not IsObject(paraValue) Then
            Set colValue = mctrlInput.ValueList(paraValue + 0)
        Else
            Set colValue = paraValue
        End If
    Else
        Set colValue = mctrlInput.ValueList
    End If
    For i = 1 To colValue.Count
        strShowName = mInputTemplate(i).Name
        If Not IsNull(mcolFieldNames(strShowName)) Then
            If mcolFieldNames(strShowName) <> "" Then
                strFields = strFields & mcolFieldNames(strShowName) & SEPARATOR_SIG
                varValue = func_convert_value(strShowName, colValue(strShowName)("值"))
                '添加到值集合。
                colValues.Add varValue, strShowName
                '添加到值列表。
                strValues = strValues & varValue & SEPARATOR_SIG
            End If
        End If
    Next i
    strFields = Left(strFields, Len(strFields) - 1)
    strValues = Left(strValues, Len(strValues) - 1)
    Set funcGetNewInfo = New Collection
    With funcGetNewInfo
        .Add strFields, "字段列表"
        .Add strValues, "值列表"
        .Add colValues, "值集合"
    End With
    
    Exit Function
err_handler:
    Err.Raise Err, , "通用业务库.clsInputServer.funcGetNewInfo：获取录入信息失败！" & Error
End Function
'功能：获取InputControl中制定行的各字段值。
'输入：paraRow  行号或ValueList。
'返回：值集合,key:字段名。
Public Function funcGetFieldsValues(Optional paraRow) As Collection
    Dim colValue As Collection
    Dim i As Long
    Dim varValue                  '某字段值。
    Dim strField As String        '字段名。
    Dim strShowName As String     '字段显示名。
    
    On Error GoTo err_handler
    
    Set funcGetFieldsValues = New Collection
    If Not IsMissing(paraRow) Then
        If Not IsObject(paraRow) Then
            Set colValue = mctrlInput.ValueList(paraRow + 0)
        Else
            Set colValue = paraRow
        End If
    Else
        Set colValue = mctrlInput.ValueList
    End If
    For i = 1 To colValue.Count
        strShowName = mInputTemplate(i).Name '字段显示名。
        '根据字段显示名获取字段名称。
        If Not IsNull(mcolFieldNames(strShowName)) Then
            strField = mcolFieldNames(strShowName)
            '获取值，并转换数据类型。
            If strField <> "" Then
                varValue = func_convert_value(strShowName, colValue(strShowName)("值"))
                '添加到值集合。
                funcGetFieldsValues.Add varValue, strField
            End If
        End If
    Next i
    Exit Function
err_handler:
    Err.Raise Err, , "通用业务库.clsInputServer.funcGetFieldsValues：获取录入字段值失败！" & Error
    Exit Function
    Resume
End Function
'功能：获取InputControl中指定行的各字段值(字符串和日期不加单引号)。
'输入：paraRow  行号或ValueList。
'返回：值集合,key:字段名。
Public Function funcGetFieldsValuesNotConvert(Optional paraRow) As Collection
    Dim colValue As Collection
    Dim i As Long
    Dim varValue                  '某字段值。
    Dim strField As String        '字段名。
    Dim strShowName As String     '字段显示名。
    
    On Error GoTo err_handler
    
    Set funcGetFieldsValuesNotConvert = New Collection
    If Not IsMissing(paraRow) Then
        If Not IsObject(paraRow) Then
            Set colValue = mctrlInput.ValueList(paraRow + 0)
        Else
            Set colValue = paraRow
        End If
    Else
        Set colValue = mctrlInput.ValueList
    End If
    For i = 1 To colValue.Count
        strShowName = mInputTemplate(i).Name '字段显示名。
        '根据字段显示名获取字段名称。
        If Not IsNull(mcolFieldNames(strShowName)) Then
            strField = mcolFieldNames(strShowName)
            '获取值，并转换数据类型。
            If strField <> "" Then
                varValue = func_convert_value(strShowName, colValue(strShowName)("值"), False)
                '添加到值集合。
                funcGetFieldsValuesNotConvert.Add varValue, strField
            End If
        End If
    Next i
    Exit Function
err_handler:
    Err.Raise Err, , "通用业务库.clsInputServer.funcGetFieldsValuesNotConvert：" & Error
    Exit Function
    Resume
End Function
'功能：获取录入字段列表。
Public Function funcGetFieldsList() As String
    Dim strFields As String       '字段名列表。
    Dim i As Integer
    Dim strShowName As String     '字段显示名。
    
    For i = 1 To mctrlInput.InputTemplate.Count
        strShowName = mInputTemplate(i).Name
        If Not IsNull(mcolFieldNames(strShowName)) Then
            strFields = strFields & mcolFieldNames(strShowName) & SEPARATOR_SIG
        End If
    Next i
    strFields = Left(strFields, Len(strFields) - 1)
    funcGetFieldsList = strFields
End Function
'功能：获取修改记录的信息。
'输入：paraRow  被修改的行。
'      paraNewValue  新值。
'返回：修改信息，"字段列表","值列表","值集合","主键旧值集合","条件"。
Public Function funcGetUpdateInfo(ByVal paraRow As Long, _
                                ByVal paraNewValue As Collection) As Collection
    Dim colInfo As New Collection
    Dim colPriValues As New Collection '主键旧值。
    
    '获取新值的字段，值列表，值集合。
    Set colInfo = funcGetNewInfo(paraNewValue)
    
    '获取主键条件和旧值集合。
    Set colPriValues = funcGetOldPriInfo(paraRow)
    
    colInfo.Add colPriValues("条件"), "条件"          '条件。
    colInfo.Add colPriValues("主键旧值集合"), "主键旧值集合"  '主键旧值。
    Set funcGetUpdateInfo = colInfo
End Function
'功能：获取记录的旧主键信息。
'输入：paraRow  被修改的行。
'      isPad    是否是录入板。
'返回：删除条件,"条件","主键旧值集合"。
Public Function funcGetOldPriInfo(ByVal paraRow As Long, Optional isPad) As Collection
Attribute funcGetOldPriInfo.VB_Description = "获取录入网格中删除记录的信息：""critical""。"
    Dim strCritical As String
    Dim i As Integer
    Dim colValues As New Collection
    Dim strShowName As String       '主键字段显示名。
    Dim strFieldName As String
    Dim varValue                    '控件上主键值。
    
    On Error GoTo err_handler
    
    '获取主键条件。
    For i = 1 To mcolPriKey.Count
        strShowName = mcolPriKey(i)
        On Error Resume Next
        strFieldName = mcolFieldNames(strShowName)
        If Err <> 0 Then strFieldName = strShowName
        Err.Clear
        On Error GoTo err_handler
        
        '是录入网格。
        varValue = mctrlInput.Value(paraRow, strShowName)
        
        '转换值格式。
        If Not mblnAutoIncre Then
            strValue = func_convert_value(strShowName, varValue)
        Else
            strValue = varValue
        End If
        colValues.Add strValue, strShowName
        strCritical = strCritical & strFieldName & "=" & strValue & " and "
    Next i
    If strCritical <> "" Then strCritical = Left(strCritical, Len(strCritical) - 5)
    
    Set funcGetOldPriInfo = New Collection
    With funcGetOldPriInfo
        .Add strCritical, "条件"
        .Add colValues, "主键旧值集合"
    End With
    Exit Function
err_handler:
    Err.Raise Err, , "通用业务库.clsInputServer.funcGetOldPriInfo：" & Error
End Function
'功能：把值转换为字符串。
'输入：paraShowName 显示名。
'      paraValue    值。
'      paraConvert  转换后的字符串和日期型数据是否加单引号（缺省要加）。
Private Function func_convert_value(ByVal paraShowName As String, _
                                    ByVal paraValue, _
                                    Optional paraConvert As Boolean = True) As String
    Dim i As Integer
    Dim strValue As String
    Dim strItem As String
    '若是字典项，截取规定部分。
    If IsNull(paraValue) Then
        paraValue = "null"
    Else
        Select Case mcolEnumMethod(paraShowName)
            Case "编码名称"
            Case "编码"
                strValue = paraValue
                If InStr(strValue, ",") > 0 Then
                    paraValue = ""
                    Do While InStr(strValue, ",") > 0
                        strItem = Left(strValue, InStr(strValue, ",") - 1)
                        If InStr(strItem, " ") > 0 Then
                            strItem = Left(strItem, InStr(strItem, " ") - 1)
                        End If
                        paraValue = paraValue & strItem & ","
                        strValue = Right(strValue, Len(strValue) - InStr(strValue, ","))
                    Loop
                Else
                    If InStr(paraValue, " ") > 0 Then
                        paraValue = Left(paraValue, InStr(paraValue, " ") - 1)
                    End If
                End If
                If Right(paraValue, 1) = "," Then paraValue = Left(paraValue, Len(paraValue) - 1)
                If mcolFieldTypes(paraShowName) = 3 Then
                    '转换为数字。
                    paraValue = Val(paraValue)
                End If
            Case "名称"
                If InStr(paraValue, " ") > 0 Then
                    paraValue = Right(paraValue, Len(paraValue) - InStr(paraValue, " "))
                End If
        End Select
        If paraConvert Then
            '转换值为SQL语句规定的格式，
            '比如，字符串和日期型加单引号，布尔变为小整型。
            paraValue = funcVarToStr(paraValue)
        Else
            '只转换布尔型为整形。
            If VarType(paraValue) = vbBoolean Then
                paraValue = IIf(paraValue, 1, 0)
            End If
        End If
    End If
    func_convert_value = paraValue
End Function
Public Property Get PrimaryKey() As Variant
Attribute PrimaryKey.VB_Description = "被录入表的主键。"
    PrimaryKey = mstrPriKey
End Property
Public Property Let PrimaryKey(ByVal vNewValue As Variant)
    Dim strKeys As String
    Dim strField As String
    mstrPriKey = Trim(vNewValue)
    If mstrPriKey <> "" And Right(mstrPriKey, 1) <> SEPARATOR_SIG Then
        mstrPriKey = mstrPriKey & SEPARATOR_SIG
    End If
    '获取主键集合。
    Set mcolPriKey = New Collection
    strKeys = mstrPriKey
    Do While InStr(strKeys, SEPARATOR_SIG) > 0
        strField = Left(strKeys, InStr(strKeys, SEPARATOR_SIG) - 1)
        strKeys = Right(strKeys, Len(strKeys) - InStr(strKeys, SEPARATOR_SIG))
        mcolPriKey.Add strField, strField
    Loop
End Property
'功能：检查输入的合法性。
'输入：paraInfo:"值集合","主键旧值集合","动作"
Public Sub subCheckValid(ByVal paraInfo As Collection)
    Dim i As Integer
    Dim strSql As String
    Dim recTmp As Object
    Dim strFldName As String  '主键字段名称。
    
    If mblnAutoIncre Then Exit Sub
    
    '检查主键唯一性。
    If mcolPriKey.Count > 0 Then
        strSql = "select " & mcolFieldNames(mcolPriKey(1)) & " from " & mstrTableName & " where "
        For i = 1 To mcolPriKey.Count
            strFldName = mcolFieldNames(mcolPriKey(i))
            strSql = strSql & " " & strFldName & "=" & paraInfo("值集合")(mcolPriKey(i)) & " and "
        Next i
        If paraInfo("动作") = UPDATE_ACTION Then
            '修改。
            For i = 1 To mcolPriKey.Count
                strFldName = mcolFieldNames(mcolPriKey(i))
                strSql = strSql & " " & strFldName & "<>" & paraInfo("主键旧值集合")(mcolPriKey(i)) & " and "
            Next i
        End If
        strSql = Left(strSql, Len(strSql) - 5)
        If mstrCritical <> "" Then strSql = strSql & " and " & mstrCritical
        mDbInterfaceObj.subExecuteSql strSql, recTmp
        If recTmp.recordcount > 0 Then
            On Error GoTo 0
            Err.Raise ECS_COMM_OP + EC_INVALID_UNIQUE, , mstrPriKey & "不能重复。"
        End If
    End If
End Sub

Public Property Get PriKeyCol() As Variant
    Set PriKeyCol = mcolPriKey
End Property
'功能：在录入网格中查找给定主键值的行。
'输入：paraInfo: 若主键是单字段，该参数是主键值；
'                若主键是多字段，该参数是主键值集合,key:主键值集合。
'返回：找到的行号（未找到返回0)。
Public Function funcFindInInputGrid(ByVal paraInfo As Variant) As Integer
    Dim i As Integer, j As Integer
    With mctrlInput
        For i = 1 To .Rows
            If mcolPriKey.Count > 1 Then
                For j = 1 To mcolPriKey.Count
                    If .Value(i, mcolPriKey(j)) <> paraInfo(mcolPriKey(j)) Then
                        Exit For
                    End If
                Next j
            Else
                If .Value(i, mcolPriKey(1)) = paraInfo Then
                    j = 2
                End If
            End If
            If j > mcolPriKey.Count Then
                '找到。
                funcFindInInputGrid = i
                Exit For
            End If
        Next i
    End With
End Function

Public Property Get TableName() As Variant
    TableName = mstrTableName
End Property

Public Property Let TableName(ByVal vNewValue As Variant)
    mstrTableName = vNewValue
End Property

Public Property Get Critical() As Variant
    Critical = mstrCritical
End Property

Public Property Let Critical(ByVal vNewValue As Variant)
    mstrCritical = vNewValue
End Property

Public Property Get DbInterfaceObj() As Variant
    Set DbInterfaceObj = mDbInterfaceObj
End Property

Public Property Set DbInterfaceObj(ByVal vNewValue As Variant)
    Set mDbInterfaceObj = vNewValue
End Property

'根据显示名获取字段名。
Public Property Get FieldName(ByVal paraShowName As String) As String
    FieldName = IIf(IsNull(mcolFieldNames(paraShowName)), paraShowName, mcolFieldNames(paraShowName))
End Property
'根据显示名获取字段类型。
Public Property Get FieldType(ByVal paraShowName As String) As Integer
    On Error Resume Next
    FieldType = mcolFieldTypes(paraShowName)
    If Err <> 0 Then
        FieldType = 0
    End If
End Property
'功能：恢复控件录入项的“允许录入”等属性。
Public Sub subRefresh()
    Dim recTmp As Object
    
    '获取模板信息。
    mDbInterfaceObj.subExecuteSql "select * from 系统_录入模板信息表 where 表名='" & mstrName & "' order by 序号", recTmp
    With mctrlInput
        Do While Not recTmp.EOF
            On Error Resume Next
            .ItemEnabled(recTmp!显示名) = IIf(IsNull(recTmp!允许录入), True, recTmp!允许录入)
            Err.Clear
            On Error GoTo err_handler
            
            recTmp.movenext
        Loop
    End With

End Sub

Public Property Get AutoIncre() As Boolean
    AutoIncre = mblnAutoIncre
End Property

Public Property Let AutoIncre(ByVal vNewValue As Boolean)
    mblnAutoIncre = vNewValue
End Property

'功能：根据内部名称获取显示名称。
'输入：paraInnerName 内部名称。
'返回：显示名称。
Public Property Get ShowName(ByVal paraInnerName As String) As String
    On Error Resume Next
    ShowName = mcolInnerNames(paraInnerName)
    If Err <> 0 Then ShowName = paraInnerName
End Property
'根据内部名称获取录入网格值。
'输入：paraRow         行号（-1取当前行）。
'      paraInnerName   内部名称。
Public Property Get GridInnerValue(ByVal paraRow As Long, ByVal paraInnerName As String) As Variant
    Dim strItemName As String '显示名。
    '根据内部名称获取显示名。
    On Error Resume Next
    strItemName = mcolInnerNames(paraInnerName)
    '根据显示名获取值。
    If paraRow = -1 Then
        '当前行。
        GridInnerValue = mctrlInput.Value(mctrlInput.Row, strItemName)
    Else
        GridInnerValue = mctrlInput.Value(paraRow, strItemName)
    End If
End Property
Public Property Let GridInnerValue(ByVal paraRow As Long, ByVal paraInnerName As String, ByVal vNewValue As Variant)
    Dim lngLen As Long '列宽。
    Dim strItemName As String '显示名。
    
    '根据内部名称获取显示名。
    On Error Resume Next
    strItemName = mcolInnerNames(paraInnerName)
    '根据显示名设置值。
    If paraRow = -1 Then
        mctrlInput.Value(mctrlInput.Row, strItemName) = vNewValue
    Else
        mctrlInput.Value(paraRow, strItemName) = vNewValue
    End If
End Property
