VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalFunction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'作者：杨春

'功能：弹出ComboBox控件的列表。
'作者：杨春
Public Sub gfsubShowComboList(paraComboBox As Object)
    On Error GoTo errHandler
    Call SendMessage(paraComboBox.hwnd, CB_SHOWDROPDOWN, 1, 0)
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gfsubShowComboList", Err.Number, Err.Description, True
End Sub
'功能：消失ComboBox控件的弹出的列表。
'作者：杨春
'时间：2001-5-12。
Public Sub gfsubHideComboList(paraComboBox As Object)
    On Error GoTo errHandler
    Call SendMessage(paraComboBox.hwnd, CB_SHOWDROPDOWN, 0, 0)
    
    Exit Sub
errHandler:
    
End Sub


'目的：子串替换。
'输入：paraStr      ：被替换的串；
'      paraSubStr   ：被替换的子串；
'      paraNewSubStr：替换为的子串。
'返回：被替换后的串。
'作者：杨春
Public Function gffuncStrReplace(ByVal paraStr As String, _
                                ByVal paraSubStr As String, _
                                ByVal paraNewSubStr As String) As String
    Dim i As Integer
    Dim intSubStrLen As Integer
    Dim intNewSubStrLen As Integer
    Dim strNew As String
    
    On Error GoTo errHandler
    intSubStrLen = Len(paraSubStr)
    intNewSubStrLen = Len(paraNewSubStr)
    i = 1
    Do While (i + intSubStrLen - 1) <= Len(paraStr)
        If Mid(paraStr, i, intSubStrLen) = paraSubStr Then
            paraStr = Left(paraStr, i - 1) & paraNewSubStr & Right(paraStr, Len(paraStr) - i - intSubStrLen + 1)
            i = i + intNewSubStrLen
        Else
            i = i + 1
        End If
    Loop
    gffuncStrReplace = paraStr
    
    Exit Function
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gffuncStrReplace", Err.Number, Err.Description, True
End Function

' 输入: paraString :字符串;
'       paraItemIndex:被查找的项的位置,从1开始
'       paraSep   分割符(长度为1)
' 返回: 被查找的分量;
'作者：杨春
Public Function gffuncGetItemFromList(ByVal paraString As String, _
                                    ByVal paraItemIndex As Integer, _
                                    ByVal paraSep As String) As String
    Dim intCount As Integer
    
    On Error GoTo errHandler
    intCount = 1
    paraString = Trim(paraString)
    If paraString <> "" Then
        paraString = paraString & paraSep
    End If
    Do While intCount < paraItemIndex And InStr(paraString, paraSep) > 0
        If Left(paraString, 1) = "'" Then
            paraString = Right(paraString, Len(paraString) - InStr(paraString, "'" & paraSep) - 1)
        Else
            paraString = Right(paraString, Len(paraString) - InStr(paraString, paraSep))
        End If
        intCount = intCount + 1
    Loop
    If intCount = paraItemIndex And InStr(paraString, paraSep) > 0 Then
        If Left(paraString, 1) <> "'" Then
            gffuncGetItemFromList = Left(paraString, InStr(paraString, paraSep) - 1)
        Else
            gffuncGetItemFromList = Left(paraString, InStr(paraString, "'" & paraSep))
        End If
    Else
        gffuncGetItemFromList = ""
    End If
    
    Exit Function
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gffuncGetItemFromList", Err.Number, Err.Description, True
End Function

' 功能：获取列表中制定项目的序号。
' 输入：paraList  字符串;
'       paraItem  被查找的项
'       paraSep   分割符（长度为1）。
' 返回：位置;
'作者：杨春
Public Function gffuncGetItemIndexFromList(ByVal paraList As String, _
                                ByVal paraItem As String, _
                                ByVal paraSep As String) As Integer
    Dim intIndex As Integer
    Dim strItem As String
    
    On Error GoTo errHandler
    intIndex = 0
    gffuncGetItemIndexFromList = 0
    paraList = Trim(paraList)
    If paraList <> "" Then paraList = paraList & paraSep
    Do While InStr(paraList, paraSep)
        intIndex = intIndex + 1
        If Left(paraList, 1) = "'" Then
            strItem = Left(paraList, InStr(paraList, "'" & paraSep))
            paraList = Right(paraList, Len(paraList) - InStr(paraList, "'" & paraSep))
        Else
            strItem = Left(paraList, InStr(paraList, paraSep) - 1)
            paraList = Right(paraList, Len(paraList) - InStr(paraList, paraSep))
        End If
        If strItem = paraItem Then
            gffuncGetItemIndexFromList = intIndex
            Exit Function
        End If
    Loop
    
    Exit Function
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gffuncGetItemIndexFromList", Err.Number, Err.Description, True
End Function

'功能：获取串中子串个数。
'输入：paraList  串；
'      paraSep   分割符(长度为1)；
'返回：子串个数。
'作者：杨春
Public Function gffuncGetItemsCountInList(ByVal paraList As String, _
                                        ByVal paraSep As String) As Integer
    Dim intCount As Integer
    
    On Error GoTo errHandler
    intCount = 0
    paraList = Trim(paraList)
    If paraList <> "" Then
        paraList = paraList & paraSep
    End If
    Do While InStr(paraList, paraSep) > 0
        intCount = intCount + 1
        If Left(paraList, 1) = "'" Then
            paraList = Right(paraList, Len(paraList) - InStr(paraList, "',"))
        Else
            paraList = Right(paraList, Len(paraList) - InStr(paraList, paraSep))
        End If
    Loop
    gffuncGetItemsCountInList = intCount

    Exit Function
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gffuncGetItemsCountInList", Err.Number, Err.Description, True
End Function

' 目的： 检查某项目是否在Listbox控件中；
' 输入： paraListbox :listbox控件；
'        paraItem    :被检查的项目；
' 输出：无；
' 返回：若找到了，即找到项目的下标，否则-1。
'作者：杨春
Public Function gffuncItemIsInListBox(ByVal paraListBox As Object, _
                                    ByVal paraItem As String) As Integer
    Dim i As Integer
    On Error GoTo errHandler
    For i = 0 To paraListBox.ListCount - 1
        If Left(paraListBox.List(i), Len(Trim(paraItem))) = Trim(paraItem) Then
            gffuncItemIsInListBox = i
            Exit Function
        End If
    Next i
    gffuncItemIsInListBox = -1

    Exit Function
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gffuncItemIsInListBox", Err.Number, Err.Description, True
End Function


' 目的：检查某项是否在ComboBox控件中；
' 输入：paraComboBox :ComboBox控件；
'       paraItem  :被检查的项；
'       paraMethod:可选，取值方式(0:第一部分，1第二部分)；
' 返回：找到项的listindex，若未找到，返回-1。
'作者：杨春
Public Function gffuncItemIsInComboBox(paraComboBox As Object, _
                                    paraItem As String, _
                                    Optional paraMethod) As Integer
    Dim i As Integer
    
    On Error GoTo errHandler
    If Trim(paraItem) = "" Then
        gffuncItemIsInComboBox = -1
        Exit Function
    End If
    For i = 0 To paraComboBox.ListCount - 1
        If IsMissing(paraMethod) Then
            If Left(paraComboBox.List(i), Len(Trim(paraItem))) = Trim(paraItem) Then
                gffuncItemIsInComboBox = i
                Exit Function
            End If
        Else
            If paraMethod = 0 And Left(paraComboBox.List(i), Len(Trim(paraItem))) = Trim(paraItem) Then
                gffuncItemIsInComboBox = i
                Exit Function
            End If
            If paraMethod = 1 And InStr(paraComboBox.List(i), " ") > 0 Then
                If Mid(paraComboBox.List(i), InStr(paraComboBox.List(i), " ") + 1, InStr(InStr(paraComboBox.List(i), " ") + 1, paraComboBox.List(i), " ") - InStr(paraComboBox.List(i), " ") - 1) = Trim(paraItem) Then
                    gffuncItemIsInComboBox = i
                    Exit Function
                End If
            End If
        End If
    Next i
    gffuncItemIsInComboBox = -1

    Exit Function
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gffuncItemIsInComboBox", Err.Number, Err.Description, True
End Function


' 目的：把某记录的某字段加入下拉列表框中；
' 输入： paraComboBox    :下拉列表框；;
'        paraRecordset   :记录;
'        paraFieldIndexes:(可选)字段号(或字段名)；缺省是0
'作者：杨春
Public Sub gfsubLoadComboFromRecordset(paraComboBox As Object, _
                                paraRecordset As Object, _
                                Optional paraFieldIndexes)
    Dim field_value
    Dim field_index
    Dim values_list As String
    Dim pos As Integer
    
    On Error GoTo errHandler
    paraComboBox.Clear
    If paraRecordset.recordcount > 0 Then
        If Not IsMissing(paraFieldIndexes) Then
            paraFieldIndexes = Trim(paraFieldIndexes) & ","
        End If
        paraRecordset.MoveFirst
        Do While Not paraRecordset.EOF
            If IsMissing(paraFieldIndexes) Then
                values_list = IIf(IsNull(paraRecordset.Fields(0)), "", paraRecordset.Fields(0))
            Else
                pos = 0:  values_list = ""
                Do
                    field_index = Mid(paraFieldIndexes, pos + 1, InStr(pos + 1, paraFieldIndexes, ",") - pos - 1)
                    On Error Resume Next
                    '先按字段名获取值：
                    field_value = IIf(IsNull(paraRecordset.Fields(field_index)), "", paraRecordset.Fields(field_index))
                    If Err <> 0 Then
                        '若失败，再按字段号获取值：
                        field_index = Val(field_index)
                        field_value = IIf(IsNull(paraRecordset.Fields(field_index)), "", paraRecordset.Fields(field_index))
                        Err.Clear
                        On Error GoTo 0
                    End If
                    values_list = values_list & field_value & " "
                    pos = InStr(pos + 1, paraFieldIndexes, ",")
                Loop Until pos = Len(paraFieldIndexes)
            End If
                
            paraComboBox.AddItem values_list
            paraRecordset.MoveNext
        Loop
    End If

    Exit Sub
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gfsubLoadComboFromRecordset", Err.Number, Err.Description, True
End Sub
' 目的：把某记录的某字段内容加入列表框中;
' 输入：paraListBox     ：列表框;
'       paraRec         ：记录;
'       paraFieldIndexes：字段名串或字段号串，各项间以逗号隔开;
'作者：杨春
Public Sub gfsubLoadListBoxFromRecordset(paraListBox As Object, _
                                    paraRec As Object, _
                                    ByVal paraFieldIndexes As String)
    Dim varFieldValue
    Dim intFieldIndex
    Dim strValues As String
    Dim pos As Integer
    
    On Error GoTo errHandler
    paraListBox.Clear
    If paraRec.recordcount > 0 Then
        paraRec.MoveFirst
        paraFieldIndexes = Trim(paraFieldIndexes) & ","
        Do While Not paraRec.EOF
            pos = 0:  strValues = ""
            Do
                intFieldIndex = Mid(paraFieldIndexes, pos + 1, InStr(pos + 1, paraFieldIndexes, ",") - pos - 1)
                On Error Resume Next
                varFieldValue = IIf(IsNull(paraRec.Fields(intFieldIndex)), "", paraRec.Fields(intFieldIndex))
                If Err <> 0 Then
                    intFieldIndex = Val(intFieldIndex)
                    varFieldValue = IIf(IsNull(paraRec.Fields(intFieldIndex)), "", paraRec.Fields(intFieldIndex))
                    Err.Clear
                    On Error GoTo 0
                End If
                strValues = strValues & varFieldValue & " "
                pos = InStr(pos + 1, paraFieldIndexes, ",")
            Loop Until pos = Len(paraFieldIndexes)
            paraListBox.AddItem strValues
            paraRec.MoveNext
        Loop
    End If

    Exit Sub
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gfsubLoadListBoxFromRecordset", Err.Number, Err.Description, True
End Sub
'功能：把记录装入网格中。
'       paraNotResize      :是否不调整行列宽度（缺省要调整）。
'       paraFields          要显示的字段名串（各项以','隔开）。
'作者：杨春
Public Sub gfsubLoadGridFromRec(ByVal paraGrid As Object, _
                            ByVal paraRec As Variant, _
                            Optional paraNotResize As Boolean = False, _
                            Optional paraFields)
    Dim i As Integer
    Dim blnShow As Boolean
    Dim strFieldName As String
    Dim blnVsFlexGrid As Boolean
    
    On Error GoTo errHandler
    blnVsFlexGrid = True
    If Not IsMissing(paraFields) Then paraFields = IIf(Right(paraFields, 1) <> ",", paraFields & ",", paraFields)
    With paraGrid
        .rows = 2: .cols = 1
        .Row = 0
        On Error Resume Next
        .reDraw = False
        On Error GoTo errHandler
        For i = 1 To paraRec.Fields.count
            blnShow = True
            If Not IsMissing(paraFields) Then
                If InStr(paraFields, paraRec(i - 1).Name & ",") = 0 Then
                    blnShow = False
                End If
            End If
            If blnShow Then
                .Col = .cols - 1
                .Text = paraRec.Fields(i - 1).Name
                If Not blnVsFlexGrid Then
                    .ColWidth(.Col) = DyLen(Trim(.Text)) * 105
                End If
                .cols = .cols + 1
            End If
        Next i
        If .cols > 1 Then
            .cols = .cols - 1
        End If
        If paraRec.recordcount > 0 Then
            paraRec.MoveFirst
            Do While Not paraRec.EOF
                For i = 0 To .cols - 1
                    .Col = i
                    .Row = 0: strFieldName = .Text
                    .Row = .rows - 1
                    If Not IsNull(paraRec(strFieldName)) Then
                        .Text = paraRec(strFieldName)
                    Else
                        .Text = ""
                    End If
                    
                    '调整列宽
                    If Not paraNotResize And .ColWidth(i) < DyLen(Trim(.Text)) * 105 Then
                        .ColWidth(i) = DyLen(Trim(.Text)) * 105
                    End If
                Next i
                paraRec.MoveNext
                .rows = .rows + 1
            Loop
            .Row = 1
        Else
            .Row = 1
            For i = 0 To .cols - 1
                .Col = i: .Text = ""
            Next i
        End If
        On Error Resume Next
        If Not paraNotResize Then
            .AutoSize 0, .cols - 1
        End If
        .reDraw = True
    End With
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gfsubLoadGridFromRec", Err.Number, Err.Description, True
End Sub
'功能：把记录以追加方式装入网格中，并保证执行列不重复。
'输入： paraGrid             vsFlexGrid
'       paraNotResize      :是否不调整行列宽度（缺省要调整）。
'       paraUniqueCol       不能重复的列。
'       paraFields          要显示的字段名串（各项以','隔开）。
'作者：杨春
Public Sub gfsubAppendGridFromRecWithUnique(ByVal paraGrid As Object, _
                                    ByVal paraRec As Variant, _
                                    ByVal paraUniqueCol As Long, _
                                    Optional paraNotResize As Boolean = False)
    Dim strFieldName As String '网格中某列的列名，也是字段名。
    Dim llngMaxRow As Long     '网格中已存在的行数。
    Dim llngRow As Long
    Dim i As Integer
    
    On Error Resume Next
    llngMaxRow = paraGrid.rows - 1
    With paraGrid
        If paraRec.recordcount > 0 Then
            paraRec.MoveFirst
            Do While Not paraRec.EOF
                '判断该记录在网格中是否已存在。
                llngRow = 1
                Do While llngRow <= llngMaxRow
                    If .TextMatrix(llngRow, paraUniqueCol) = paraRec(.TextMatrix(0, paraUniqueCol)) Then
                        '该记录在网歌中已存在。
                        Exit Do
                    End If
                    llngRow = llngRow + 1
                Loop
                If llngRow > llngMaxRow Then
                    '添加一行。
                    .rows = .rows + 1
                    llngRow = .rows - 1
                    For i = 0 To .cols - 1
                        strFieldName = .TextMatrix(0, i)
                        If Not IsNull(paraRec(strFieldName)) Then
                            .TextMatrix(llngRow, i) = paraRec(strFieldName)
                        Else
                            .TextMatrix(llngRow, i) = ""
                        End If
                        
                        '调整列宽
                        If Not paraNotResize And .ColWidth(i) < DyLen(Trim(.TextMatrix(llngRow, i))) * 105 Then
                            .ColWidth(i) = DyLen(Trim(.TextMatrix(llngRow, i))) * 105
                        End If
                    Next i
                End If
                
                paraRec.MoveNext
            Loop
        End If
    End With

End Sub
'功能：把记录以追加方式装入网格中。
'       paraNotResize      :是否不调整行列宽度（缺省要调整）。
'       paraFields          要显示的字段名串（各项以','隔开）。
'作者：杨春
Public Sub gfsubAppendGridFromRec(ByVal paraGrid As Object, _
                            ByVal paraRec As Variant, _
                            Optional paraNotResize As Boolean = False)
    Dim i As Integer
    Dim blnShow As Boolean
    Dim strFieldName As String
    
    On Error Resume Next
    With paraGrid
        If paraRec.recordcount > 0 Then
            paraRec.MoveFirst
            Do While Not paraRec.EOF
                .rows = .rows + 1
                For i = 0 To .cols - 1
                    .Col = i
                    .Row = 0: strFieldName = .Text
                    .Row = .rows - 1
                    If Not IsNull(paraRec(strFieldName)) Then
                        .Text = paraRec(strFieldName)
                    Else
                        .Text = ""
                    End If
                    
                    '调整列宽
                    If Not paraNotResize And .ColWidth(i) < DyLen(Trim(.Text)) * 105 Then
                        .ColWidth(i) = DyLen(Trim(.Text)) * 105
                    End If
                Next i
                paraRec.MoveNext
            Loop
        End If
    End With

End Sub
'作者：杨春
Private Function DyLen(ByVal paraString As String) As Long
    On Error GoTo errHandler
    Dim i As Long
    Dim lngLen As Long
    lngLen = 0
    i = 1
    Do While i <= Len(paraString)
        If Asc(Mid(paraString, i, 1)) < 0 Then
            lngLen = lngLen + 2
        Else
            lngLen = lngLen + 1
        End If
        i = i + 1
    Loop
    DyLen = lngLen
    Exit Function
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "DyLen", Err.Number, Err.Description, True
End Function

'功能：把Recordset内容加入道源录入网格中。
'作者：杨春
Public Sub gfsubLoadDyGridFromRec(ByVal paraControl As Object, ByVal paraRec As Object)
'paraControl 网格对象 paraRec 数据库对象
    Dim i As Integer
    Dim colRow As Collection
    
    On Error GoTo errHandler
    paraControl.subClear
    If paraRec.recordcount > 0 Then
        paraRec.MoveFirst
        Do While Not paraRec.EOF
            Set colRow = New Collection
            With colRow
                For i = 0 To paraRec.Fields.count - 1
                    colRow.Add paraRec(i).Value, paraRec(i).Name
                Next i
            End With
            paraControl.subAddRow colRow
            paraRec.MoveNext
        Loop
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gfsubLoadDyGridFromRec", Err.Number, Err.Description, True
End Sub

'--------------------------------------------------------------------------
' 功能: 把字符串(以逗号隔开各分量)中某分量的值替换为指定值;
' 输入: paraList      :字符串;
'       paraIndex:被替换分量的位置(从1开始数);
'       item      :替换成的串(若不存在，表示删除该项);
'       paraSep   :分割符号（一个字符）。
' 输出: paraList;
'作者：杨春
Public Sub gfsubReplaceListItem(paraList As String, _
                            ByVal paraIndex As Integer, _
                            ByVal paraSep As String, _
                            Optional item)
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim count As Integer
    
    On Error GoTo errHandler
    count = 1
    pos2 = 1
    pos1 = 1
    paraList = Trim(paraList)
    If paraList = "" Then
        Exit Sub
    End If
    paraList = paraList & paraSep
    pos2 = InStr(pos1, paraList, paraSep)
    Do While count < paraIndex And pos2 < Len(paraList)
        pos1 = pos2 + 1
        count = count + 1
        pos2 = InStr(pos1, paraList, paraSep)
    Loop
    If count = paraIndex Then
        If Not IsMissing(item) Then
            paraList = Left(paraList, IIf(pos1 > 0, pos1 - 1, 0)) & item & Right(paraList, Len(paraList) - pos2 + 1)
        Else
            paraList = Left(paraList, IIf(pos1 > 1, pos1 - 2, 0)) & IIf(Right(paraList, Len(paraList) - pos2) = "", "", paraSep) & Right(paraList, Len(paraList) - pos2)
        End If
    End If
    If paraList <> "" Then
        If Left(paraList, 1) = paraSep Then
            paraList = Right(paraList, Len(paraList) - 1)
        End If
        If Right(paraList, 1) = paraSep Then
            paraList = Left(paraList, Len(paraList) - 1)
        End If
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检公用信息部件", "clsGlobalFunction", "gfsubReplaceListItem", Err.Number, Err.Description, True
End Sub

'功能：根据公民身份号码获取生日和性别。
'作者：杨春
Public Sub sub根据公民身份号码获取生日和性别(ByVal para公民身份号码 As String, para出生日期 As String, para性别 As String)
    On Error GoTo errHandler
    If Len(para公民身份号码) >= 15 And Len(para公民身份号码) < 18 Then
        para性别 = IIf((Mid(para公民身份号码, 15, 1) Mod 2) = 0, "女", "男")
        para出生日期 = "19" & Mid(para公民身份号码, 7, 6)
        para出生日期 = Left(para出生日期, 4) & "-" & Mid(para出生日期, 5, 2) & "-" & Right(para出生日期, 2)
    ElseIf Len(para公民身份号码) >= 18 Then
        para性别 = IIf((Mid(para公民身份号码, 17, 1) Mod 2) = 0, "女", "男")
        para出生日期 = Mid(para公民身份号码, 7, 8)
        para出生日期 = Left(para出生日期, 4) & "-" & Mid(para出生日期, 5, 2) & "-" & Right(para出生日期, 2)
    Else
        para性别 = ""
        para出生日期 = ""
    End If
    Exit Sub
errHandler:

End Sub

