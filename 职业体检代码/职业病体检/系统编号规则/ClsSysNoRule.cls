VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSysNoRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'作者：邓恒
'最后修改者：杨春
Private mstr编号名称 As String
Private mcol分段规则 As New Collection '[分段号，组成，该段是否允许用户修改，分段长度，常量值]。
Private mstr流水号前面固定部分  As String
Private mlng总长度 As Long

Public Property Get 编号名称() As String
    编号名称 = mstr编号名称
End Property

'功能：设置本对象关键属性，从而获取其它属性。
'作者：邓恒
'最后修改：杨春（2001-4-4）
Public Property Let 编号名称(ByVal vNewValue As String)
    Dim lobjRec As Object
    Dim lcolItem As Collection
    Dim lstrConstant As String '分段的常量值。
    Dim lstrSql As String
    Dim i As Integer
    
    On Error GoTo errHandler
    mstr编号名称 = vNewValue
    mstr流水号前面固定部分 = ""
    Set mcol分段规则 = New Collection
    mlng总长度 = 0
    If mstr编号名称 <> "" Then
        '获取编号生成规则
        lstrSql = "select * from 体检管理_编号生成规则表 where 编号名称='" & mstr编号名称 & "' order by 分段号 asc"
        Set lobjRec = dafuncGetData(lstrSql)
        If lobjRec.recordcount > 0 Then
            lobjRec.moveFirst
            Do While Not lobjRec.EOF
                Set lcolItem = New Collection
                lcolItem.Add lobjRec("分段号").Value, "分段号"
                lcolItem.Add IIf(IsNull(lobjRec!组成), "", lobjRec("组成").Value), "组成"
                lcolItem.Add IIf(IsNull(lobjRec!该段是否允许用户修改), "1", lobjRec("该段是否允许用户修改").Value), "该段是否允许用户修改"
                lcolItem.Add IIf(IsNull(lobjRec!分段长度), 1, lobjRec("分段长度").Value), "分段长度"
                
                lstrConstant = func获取分段常量值(lcolItem("组成"), lcolItem("分段长度"))
                
                lcolItem.Add lstrConstant, "常量值"
                
                mcol分段规则.Add lcolItem, Trim(Str(lcolItem("分段号")))
                
                mlng总长度 = mlng总长度 + lcolItem("分段长度")
                
                lobjRec.movenext
            Loop
        Else
            Err.Raise 6666, , "库中没有名称为“" & mstr编号名称 & "”的编号规则。"
        End If
    
    End If
    Exit Property
errHandler:
    'sfsub错误处理 "系统编号规则部件", "ClsSysNoRule", "Property Let 编号名称", Err.Number, Err.Description, True
    Exit Property
    Resume
End Property

'功能：获取分段规则。
'作者：邓恒
Public Property Get 分段规则() As Collection
    On Error GoTo errHandler
    Set 分段规则 = mcol分段规则
    Exit Property
errHandler:
    sfsub错误处理 "系统编号规则部件", "ClsSysNoRule", "Property Get 分段规则", Err.Number, Err.Description, True
End Property

'功能：获取流水号前面固定部分
'作者：邓恒
Public Property Get 流水号前面固定部分() As String
    Dim lcolItem As Collection '各分段的规则。
    Dim lstrTemp As String
    
    On Error GoTo errHandler
    
    lstrTemp = ""
    流水号前面固定部分 = ""
    For Each lcolItem In mcol分段规则
        If lcolItem("常量值") <> "" Then
            lstrTemp = lstrTemp & lcolItem("常量值")
        ElseIf lcolItem("组成") <> "流水号" Then
            Exit For
        End If
        
    Next
    流水号前面固定部分 = lstrTemp
    Exit Property
errHandler:
    sfsub错误处理 "系统编号规则部件", "ClsSysNoRule", "Property Get 流水号前面固定部分", Err.Number, Err.Description, True
End Property

'功能：获取当前编号名称的编码总长度。
'作者：杨春
Public Property Get 编号总长度() As Long
    编号总长度 = mlng总长度
End Property

'功能：按当前编号名称的编码规则生成系统编号。
'作者：邓恒
'最后修改：杨春（2001-4-4）
Public Function Func生成系统编号(Optional para预留参数 As Collection) As String
Attribute Func生成系统编号.VB_Description = "'功能：获取指定分组名的常量值。"
    Dim lstr系统编号 As String
    Dim lobjRec As Object
    Dim i As Integer
    
    On Error GoTo errHandler
    lstr系统编号 = ""
    
    For i = 1 To mcol分段规则.Count
        '单位编号
        If mcol分段规则(i)("常量值") <> "" Then
            lstr系统编号 = lstr系统编号 & mcol分段规则(i)("常量值")
        Else
            Select Case mcol分段规则(i)("组成")
            Case "流水号"
                Set lobjRec = dafuncGetData("exec 体检管理_生成编号流水号 '" & mstr编号名称 & "'")
                lstr系统编号 = lstr系统编号 & lobjRec(0)
            Case Else
                Err.Raise 6666, , "编号的生成规则中组成“" & mcol分段规则(i)("组成") & "”还不能识别。请系统管理员修改编号名称为“" & mstr编号名称 & "”的生成规则"
            End Select
        End If
    Next i
    Func生成系统编号 = lstr系统编号
    
    Exit Function
errHandler:
    Func生成系统编号 = ""
    'sfsub错误处理 "系统编号规则部件", "ClsSysNoRule", "Func生成系统编号", Err.Number, Err.Description, True
End Function

'功能：填加或修改一段分段规则。
'输入：paraItem as collection[分段号，组成，该段是否允许用户修改，分段长度]。
'作者：杨春
Public Sub sub设置分段规则(ByVal paraItem As Collection)
    Dim lcolItem As Collection
    Dim lstrConstant As String
    Dim lstrSql As String

    On Error GoTo errHandler
    If sffunc判断集合键值是否存在(mcol分段规则, Trim(Str(paraItem("分段号")))) Then
        '修改,先删除属性和库中记录。
        mcol分段规则.Remove mcol分段规则(Trim(Str(paraItem("分段号"))))
        dafuncGetData "delete 体检管理_编号生成规则表 where 编号名称='" & mstr编号名称 & "' and 分段号=" & paraItem("分段号")
    End If
    
    '在属性中添加元素。
    Set lcolItem = New Collection
    lcolItem.Add paraItem("分段号"), "分段号"
    lcolItem.Add IIf(IsNull(paraItem("组成")), "", paraItem("组成")), "组成"
    lcolItem.Add IIf(IsNull(paraItem("该段是否允许用户修改")), "1", paraItem("该段是否允许用户修改")), "该段是否允许用户修改"
    lcolItem.Add IIf(IsNull(paraItem("分段长度")), 1, paraItem("分段长度")), "分段长度"
    '获取分段的常量值。
    lstrConstant = func获取分段常量值(lcolItem("组成"), lcolItem("分段长度"))
    lcolItem.Add lstrConstant, "常量值"
    mcol分段规则.Add lcolItem, Trim(Str(paraItem("分段号")))
    
    '在库中添加记录。
    lstrSql = "insert into 体检管理_编号生成规则表(编号名称,分段号,组成,该段是否允许用户修改,分段长度) values('" & mstr编号名称 & "','" & lcolItem("分段号") & _
              "','" & lcolItem("组成") & "','" & lcolItem("该段是否允许用户修改") & "','" & lcolItem("分段长度") & "')"
    dafuncGetData lstrSql
    
    Exit Sub
errHandler:
    sfsub错误处理 "系统编号规则部件", "ClsSysNoRule", "sub设置分段规则", Err.Number, Err.Description, True
End Sub

'作者：杨春
Public Sub sub删除分段规则(ByVal para分段号 As Long)
    On Error GoTo errHandler
    
    If sffunc判断集合键值是否存在(mcol分段规则, Trim(Str(para分段号))) Then
        '删除属性种元素。
        mcol分段规则.Remove mcol分段规则(Trim(Str(para分段号)))
    End If
    
    '删除库中记录。
    dafuncGetData "delete 体检管理_编号生成规则表 where 编号名称='" & mstr编号名称 & "' and 分段号=" & para分段号

    Exit Sub
errHandler:
    sfsub错误处理 "系统编号规则部件", "ClsSysNoRule", "sub删除分段规则", Err.Number, Err.Description, True
End Sub

'功能：获取指定分组名的常量值。
'作者：杨春
Private Function func获取分段常量值(ByVal para组成名 As String, ByVal para分段长度 As String) As String
    Dim lstrConstant As String
    
    On Error GoTo errHandler
    '计算分段的常量值。
    Select Case para组成名
    Case "单位编号", "防疫站编号"
        If Len(um防疫站编号) = para分段长度 Then
            lstrConstant = um防疫站编号
        ElseIf Len(um防疫站编号) < para分段长度 Then
            lstrConstant = um防疫站编号 & String(para分段长度 - Len(um防疫站编号), "0")
        Else
            Err.Raise 6666, , "编号组成“单位编号”和“防疫站编号”长度必须是" & Len(um防疫站编号) & "。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "工作站代号", "工作站编号"
        If Len(um工作站编号) = para分段长度 Then
            lstrConstant = um工作站编号
        ElseIf Len(um工作站编号) < para分段长度 Then
            lstrConstant = um工作站编号 & String(para分段长度 - Len(um工作站编号), "0")
        Else
            Err.Raise 6666, , "编号组成“工作站代号”和“工作站编号”长度必须是" & Len(um工作站编号) & "。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "服务器代号"
        If Len(um服务器代号) = para分段长度 Then
            lstrConstant = um服务器代号
        ElseIf Len(um服务器代号) < para分段长度 Then
            lstrConstant = um服务器代号 & String(para分段长度 - Len(um服务器代号), "0")
        Else
            Err.Raise 6666, , "编号组成中“服务器代号”长度必须是" & Len(um服务器代号) & "。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "yyyymmdd"
        If para分段长度 = 8 Then
            lstrConstant = Format(Date, "yyyymmdd")
        Else
            Err.Raise 6666, , "编号组成“yyyymmdd”长度必须是8。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "yymmdd"
        If para分段长度 = 8 Then
            lstrConstant = Format(Date, "yyyymmdd")
        ElseIf para分段长度 = 6 Then
            lstrConstant = Format(Date, "yymmdd")
        Else
            Err.Raise 6666, , "编号组成“yymmdd”长度必须是6或8。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "yyyy"
        If para分段长度 = 4 Then
            lstrConstant = Left(Format(Date, "yyyymmdd"), 4)
        Else
            Err.Raise 6666, , "编号组成“yyyy”长度必须是4。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "yy"
        If para分段长度 = 2 Then
            lstrConstant = Left(Format(Date, "yymmdd"), 2)
        Else
            Err.Raise 6666, , "编号组成“yy”长度必须是2。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "mm"
        If para分段长度 = 2 Then
            lstrConstant = Format$(Month(Date), "00")
        Else
            Err.Raise 6666, , "编号组成“mm”长度必须是2。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "dd"
        If para分段长度 = 2 Then
            lstrConstant = Format$(Day(Date), "00")
        Else
            Err.Raise 6666, , "编号组成“dd”长度必须是2。请系统管理员重新设置编号名称为“" & mstr编号名称 & "”的系统编号规则中的分段长度。"
        End If
    Case "流水号"
        lstrConstant = ""
    Case Else
        lstrConstant = para组成名
    End Select

    func获取分段常量值 = lstrConstant
    
    Exit Function
errHandler:
    sfsub错误处理 "系统编号规则部件", "ClsSysNoRule", "func获取分段常量值", Err.Number, Err.Description, True
End Function
