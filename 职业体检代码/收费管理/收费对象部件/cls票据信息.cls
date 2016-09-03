VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls票据信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Mstr编号 As String
Private Mint票据类型 As Integer
Private mstr票据名称 As String
Private Mstr票据格式文件名称  As String
Private Mstr对应业务   As String
Private Mbln是否预览    As Boolean
Private Mint最大项数 As Integer

Public 项目汇总 As String


Public Property Get 编号() As String
    编号 = Mstr编号
End Property

Public Property Let 编号(ByVal vNewValue As String)
    Mstr编号 = vNewValue
End Property

Public Property Get 票据类型() As Integer
    票据类型 = Mint票据类型
End Property

Public Property Let 票据类型(ByVal vNewValue As Integer)
    Mint票据类型 = vNewValue
End Property

Public Property Get 票据格式文件名称() As String
    票据格式文件名称 = Mstr票据格式文件名称
End Property

Public Property Let 票据格式文件名称(ByVal vNewValue As String)
    Mstr票据格式文件名称 = vNewValue
End Property

Public Property Get 票据名称() As String
    票据名称 = mstr票据名称
End Property

Public Property Let 票据名称(ByVal vNewValue As String)
    mstr票据名称 = vNewValue
End Property

Public Property Get 对应业务() As String
    对应业务 = Mstr对应业务
End Property

Public Property Let 对应业务(ByVal vNewValue As String)
    Mstr对应业务 = vNewValue
End Property



Public Property Get 是否预览() As Boolean
    是否预览 = Mbln是否预览
End Property

Public Property Let 是否预览(ByVal vNewValue As Boolean)
    Mbln是否预览 = vNewValue
End Property
Public Property Let 最大项数(ByVal vNewValue As Integer)
    Mint最大项数 = vNewValue
End Property
Public Property Get 最大项数() As Integer
    最大项数 = Mint最大项数
End Property

'功能：打印票据
'输入：票据格式文件名称，要打印的内容
'输出：无
'返回：True：保存成功  False：保存失败
'注意事项：无
'作者：杨川
'创建时间：2001-3-28
'修改：2002-6-25（杨春）退费收据要求打印“负”字，所以增加参数“para退费”。
Public Function func打印票据(parastr票据文件名称 As String, _
                        paraCol As Collection, _
                        paraint打折比率 As Single, _
                        paraint最大项数 As Integer, _
                        Optional para退费 As Boolean = False, _
                        Optional paraModal As Boolean = False) As Boolean
    Dim i As Integer
    Dim lintColCount As Integer
    Dim lcol() As New Collection
    Dim lbln As Boolean
    Dim lbln提示框 As Boolean
    Dim lobj打印票据对象 As Object
    Dim lobjRec As Object
    Dim j As Integer
    On Error GoTo errhandle
    
    func打印票据 = False
    
    '创建打印对象
    On Error Resume Next
'    Set lobj打印票据对象 = CreateObject("收费票据打印对象.cls票据打印")
    
    '判是否创建成功
'    If Not lobj打印票据对象 Is Nothing Then
'        lobj打印票据对象.func票据打印 parastr票据文件名称, paraCol, paraint打折比率, paraint最大项数
'        If Err.Number = 0 Then Exit Function
'    End If
    
    
    '修改：2004-3-24（根据票据格式“项目汇总”设置，判断并且作项目汇总处理）。
    If 项目汇总 = "是" Then
        '汇总项目
        Call sub费用项目汇总(paraCol)
    End If
    
    If paraCol.Count Mod paraint最大项数 = 0 Then
        lintColCount = paraCol.Count \ paraint最大项数
    Else
        lintColCount = paraCol.Count \ paraint最大项数 + 1
    End If
    
    ReDim lcol(lintColCount - 1)
    
    
    For i = 1 To paraCol.Count
        If (i Mod paraint最大项数) > 0 Then
            lcol(i \ (paraint最大项数)).Add paraCol(i)
        Else
            lcol(i \ (paraint最大项数) - 1).Add paraCol(i)
        End If
    Next i
    
   
    For i = 0 To UBound(lcol())
        Dim lstr收据号 As String, lstr收费编号 As String
        
        '生成起始收据号。
        Set lobjRec = dafuncGetData("收费管理_产生收据号 '" & um用户编号 & "'")
        If lobjRec.recordcount > 0 Then
            lstr收据号 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
        End If
        Set lobjRec = dafuncGetData("收费管理_产生收费编号 '" & um服务器代号 & "'")
        If lobjRec.recordcount > 0 Then
            lstr收费编号 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
        End If
        
        For j = 1 To lcol(i).Count
            lcol(i)(j).Remove "收据号"
            lcol(i)(j).Add lstr收据号, "收据号"
            dafuncGetData "update 收费管理_费用信息表 set 收费编号='" & lstr收费编号 & "',收据号='" & lstr收据号 & "',票据张数=1 where 收费批号='" & paraCol(1)("收费批号") & "' and left(收费项目编号," & Len(lcol(i)(j)("收费项目编号")) & ")='" & lcol(i)(j)("收费项目编号") & "'"
        Next
    
        lbln = func打印票据分页(parastr票据文件名称, lcol(i), paraint打折比率, paraint最大项数, para退费, paraModal)
        If lbln = False Then
            lbln提示框 = sffuncMsg("打印错误", sf警告)
            Exit Function
        End If
    Next i
    Err.Clear
    
    Exit Function
errhandle:
    func打印票据 = False
    Call sfsub错误处理("收费数据对象", "cls票据信息", "func打印票据", Err.Number, Err.Description, True)
End Function

'汇总项目到第一级。
Private Sub sub费用项目汇总(paraCol As Collection)
    Dim lcol费用信息 As Collection
    Dim lcolItem As Collection
    Dim lstrParent As String
    Dim lobjRec As Object
    Dim i As Long
    Dim ldbl金额 As Double
    
    Set lcol费用信息 = New Collection
    i = 1
    Do While i <= paraCol.Count
        '获取当前项目的1级项目的信息。
        Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & Left(paraCol(i)("收费项目编号"), 3) & "'")
        If lobjRec.recordcount = 0 Then Err.Raise 6666, , "该批费用项目中“" & paraCol(i)("收费项目名称") & "”对应的上级项目不存在！"
        lstrParent = lobjRec("收费项目编号")
        
        If sffunc判断集合键值是否存在(paraCol, lstrParent) Then
            Set lcolItem = paraCol(lstrParent)
            ldbl金额 = Val(lcolItem("金额")) + Val(paraCol(i)("金额"))
            lcolItem.Remove "数量"
            lcolItem.Remove "单价"
            lcolItem.Remove "金额"
            lcolItem.Add Format(ldbl金额, "0.00"), "金额"
            lcolItem.Add 1, "数量"
            lcolItem.Add lcolItem("金额"), "单价"
            paraCol.Remove i
        Else
            Set lcolItem = paraCol(i)
            lcolItem.Remove "收费项目编号"
            lcolItem.Remove "收费项目名称"
            lcolItem.Remove "数量"
            lcolItem.Remove "单价"
            lcolItem.Add lstrParent, "收费项目编号"
            lcolItem.Add lobjRec("收费项目名称").Value, "收费项目名称"
            lcolItem.Add 1, "数量"
            lcolItem.Add lcolItem("金额"), "单价"
            paraCol.Remove i
            If paraCol.Count = 0 Then
                paraCol.Add lcolItem, lstrParent
            ElseIf paraCol.Count = i - 1 Then
                paraCol.Add lcolItem, lstrParent, , i - 1
            Else
                paraCol.Add lcolItem, lstrParent, i
            End If
            i = i + 1
        End If
            
    Loop
    
    

End Sub


'功能: 将给定金额转换为人民币的大写字符串
'输入: money       金额
'输出: FuncConvertToCapsStr     转换的大写字符串
'最后修改时间: 96.6.11
'--------------------------------------------------
Public Function FuncConvertToCapsStr(ByVal Money As Currency) As String
    Const digit_str = "零壹贰叁肆伍陆柒捌玖"
    Const unit_str = "仟佰拾万仟佰拾元角分"
    Dim money_str As String
    
    If Money > 99999999.99 Then
        FuncConvertToCapsStr = ""
    ElseIf Money = 0 Then
        FuncConvertToCapsStr = "零元整"
    Else
        Dim temp_str As String
        Dim i, j As Integer
        
        If Money < 0 Then
            money_str = "负"
            Money = -Money
        Else
            money_str = ""
        End If
        
        temp_str = Format(Money, "00000000.00")
        
        '转换整数部分
        For i = 1 To 8
            If Mid(temp_str, i, 1) <> "0" Then Exit For
        Next
        For i = i To 8
            j = CInt(Mid(temp_str, i, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & Mid(unit_str, i, 1)
            Else
                If i = 4 Then
                    money_str = money_str & "万"
                ElseIf i = 8 Then
                    money_str = money_str & "元"
                ElseIf Mid(temp_str, i + 1, 1) <> "0" Then
                    money_str = money_str & Mid(digit_str, j + 1, 1)
                End If
            End If
        Next
        
        '转换小数部分
        If Right(temp_str, 2) = "00" Then
            money_str = money_str & "整"
        Else
            '转换角
            j = CInt(Mid(temp_str, 10, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "角"
            Else
                money_str = money_str & "零"
            End If
            '转换分
            j = CInt(Mid(temp_str, 11, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "分"
            Else
                money_str = money_str & "整"
            End If
        End If
        
        FuncConvertToCapsStr = money_str
    End If
End Function
 

'功能：添加票据信息
'输入：无
'输出：无
'返回：True：保存成功  False：保存失败
'注意事项：无
'作者：杨川
'创建时间：2001-3-28
Public Function func添加票据信息() As Boolean
    
    On Error GoTo errhandle
    Dim lstr As String
    Dim lobj As Object
    Dim aa As Variant
    
    func添加票据信息 = False
    If func数据校验 = True Then
        lstr = "收费管理_添加票据信息 '" & Mstr编号 & "'," & Mint票据类型 & ",'" & mstr票据名称 & "'," & _
               "'" & Mstr票据格式文件名称 & "','" & Mstr对应业务 & "'," & Mint最大项数 & ""
               
        Set lobj = dafuncGetData(lstr)
        func添加票据信息 = True
    End If
    Exit Function
errhandle:
    Call sfsub错误处理("收费数据对象", "cls票据信息", "func添加票据信息", Err.Number, Err.Description, True)
End Function
'功能：修改票据信息
'输入：无
'输出：无
'返回：True：保存成功  False：保存失败
'注意事项：无
'作者：杨川
'创建时间：2001-3-28
Public Function func修改票据信息() As Boolean
    On Error GoTo errhandle
    Dim lstr As String
    Dim lobj As Object
    
    func修改票据信息 = False
    If func数据校验 = True Then
        lstr = "收费管理_修改票据信息 '" & Mstr编号 & "'," & Mint票据类型 & ",'" & mstr票据名称 & "'," & _
               "             '" & Mstr票据格式文件名称 & "','" & Mstr对应业务 & "'," & Mint最大项数 & ""
        Set lobj = dafuncGetData(lstr)
        
        func修改票据信息 = True
    End If
    Exit Function
errhandle:
    Call sfsub错误处理("收费数据对象", "cls票据信息", "func修改票据信息", Err.Number, Err.Description, True)
End Function
'功能：删除票据信息
'输入：无
'输出：无
'返回：True：保存成功  False：保存失败
'注意事项：无
'作者：杨川
'创建时间：2001-3-28
Public Function func删除票据信息() As Boolean
    On Error GoTo errhandle
    Dim lstr As String
    Dim lobj As Object
    func删除票据信息 = False
    If Mstr编号 <> "" Then
        lstr = "收费管理_删除票据信息 '" & Mstr编号 & "'"
        Set lobj = dafuncGetData(lstr)
        
        func删除票据信息 = True
    End If
    Exit Function
errhandle:
    Call sfsub错误处理("收费数据对象", "cls票据信息", "func删除票据信息", Err.Number, Err.Description, True)
End Function
'功能：查询票据信息
'输入：parastr查询条件
'输出：无
'返回：object '包含返回的recordset
'注意事项：无
'作者：杨川
'创建时间：2001-3-28
Public Function func查询票据信息(Optional parastr查询条件 As String) As Object
    On Error GoTo errhandle
    Dim lstr As String
  
    If parastr查询条件 = "" Then
        lstr = "select 编号,票据类型编号,票据名称,票据格式文件名称,对应业务,最大项数 " & _
             "from 收费管理_票据设置信息表"
    Else
        lstr = "select 编号,票据类型编号,票据名称,票据格式文件名称,对应业务,最大项数 " & _
             "from 收费管理_票据设置信息表  where " & parastr查询条件
    End If
    
    Set func查询票据信息 = dafuncGetData(lstr)
    
    Exit Function
errhandle:
    Call sfsub错误处理("收费数据对象", "cls票据信息", "func查询票据信息", Err.Number, Err.Description, True)
End Function
Private Sub Class_Initialize()
'Dim lstr As String
'lstr = "DSN=hospital99;uid=user26;pwd=welcome;database=防疫2001"

'dasubInitialize lstr
End Sub

'没有校验票据类型
'功能：校验票据信息
'输入：无
'输出：无
'返回：True：保存成功  False：保存失败
'注意事项：无
'作者：杨川
'创建时间：2001-3-29

Private Function func数据校验() As Boolean
    On Error GoTo errhandle
    
    If Mstr编号 = "" Then
    'Or mstr票据名称 = "" Or Mstr票据格式文件名称 Or Mstr对应业务 Then
        Err.Raise 32000, "cls票据信息", "编号不能为空"
    End If
    
    If mstr票据名称 = "" Then
        Err.Raise 32001, "cls票据信息", "票据名称不能为空"
    End If
    If Mstr票据格式文件名称 = "" Then
        Err.Raise 32002, "cls票据信息", "票据格式文件名称不能为空"
    End If
    If Mstr对应业务 = "" Then
        Err.Raise 32003, "cls票据信息", "对应业务不能为空"
    End If
    If Mint最大项数 <= 0 Then
        Err.Raise 32004, "cls票据信息", "最大项数不能小于等于0"
    End If
    
    func数据校验 = True
    Exit Function
errhandle:
    func数据校验 = False
    Call sfsub错误处理("收费数据对象", "cls票据信息", "func数据校验", Err.Number, Err.Description, True)
End Function

'修改：2002-6-25（杨春）退费收据要求打印“负”字，所以增加参数“para退费”。
Private Function func打印票据分页(parastr票据格式文件 As String, _
                                paraCol As Collection, _
                                paraint打折比率 As Single, _
                                paraint打印笔数 As Integer, _
                                Optional para退费 As Boolean = False, _
                                Optional paraModal As Boolean = False) As Boolean
On Error GoTo errhandle
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim lstr As String
    Dim lcur小写金额 As Currency
    Dim lstr大写金额 As String
   
    Dim lcol费用项目值 As New Collection  '费用项目值
    Dim lcol费用项目 As New Collection    '费用项目
    Dim lcol单价 As New Collection
    Dim lcol数量 As New Collection
    Dim lcol计量单位 As New Collection
    Dim lobj As Object
    Dim lobjRec As Object
    Dim lbln As Boolean
    
    
    'If paracol.Count > paraint打印笔数 Then
        'lbln = sffuncMsg("设置的最大打印项数超过了票据格式所容纳的最大项数" & _
                   ' vbCrLf & "请在票据格式里添加相应的变量", sf警告)
        'Exit Function
   ' End If
    func打印票据分页 = False
    
    Dim capp As New CRAXDRT.Application
    Dim cr As CRAXDRT.Report
    Dim Item As CRAXDRT.FormulaFieldDefinition
    
    Set cr = capp.OpenReport(parastr票据格式文件)
    
    'i = 100
    '初始化票据格式文件
    'With frm票据预览.CrystalReport1
'        For i = 1 To 100
'            .Formulas(i) = ""
'        Next i
'
'        .WindowParentHandle = frm票据预览.hWnd
'        .ReportFileName = "" & parastr票据格式文件 & ""
'
'        '判断是否预览
'        If 是否预览 = True Then
'            .Destination = 0
'        Else
'            .Destination = 1
'        End If
        
        For i = 1 To paraCol.Count
            lcol费用项目值.Add CCur(paraCol(i)("金额") * paraCol(i)("打折比率"))
        Next i
        
        For i = 1 To paraCol.Count
            lcol单价.Add CCur(paraCol(i)("单价"))
        Next i
        For i = 1 To paraCol.Count
            lcol数量.Add CStr(paraCol(i)("数量"))
        Next i
        For i = 1 To paraCol.Count
            lcol计量单位.Add paraCol(i)("计量单位")
        Next i
        For i = 1 To paraCol.Count
            lcol费用项目.Add paraCol(i)("收费项目名称")
        Next i
        '给票据格式文件的变量赋值
        For i = 1 To lcol费用项目值.Count
'            .Formulas(i) = "值" & i & "=" & CCur(lcol费用项目值(i))
            Set Item = cr.FormulaFields.GetItemByName("值" & i)
            Item.Text = CCur(lcol费用项目值(i))
        Next i
        
        For j = 1 To lcol费用项目.Count
'            .Formulas(j + i) = "项目" & j & "='" & lcol费用项目(j) & "'"
            Set Item = cr.FormulaFields.GetItemByName("项目" & j)
            Item.Text = "'" & lcol费用项目(j) & "'"
        Next j
        For j = 1 To lcol单价.Count
'            .Formulas(j + 2 * i) = "单价" & j & "=" & CCur(lcol单价(j))
            Set Item = cr.FormulaFields.GetItemByName("单价" & j)
            Item.Text = CCur(lcol单价(j))
        Next j
        For j = 1 To lcol数量.Count
'            .Formulas(j + 3 * i) = "数量" & j & "='" & lcol数量(j) & "'"
            Set Item = cr.FormulaFields.GetItemByName("数量" & j)
            Item.Text = "'" & lcol数量(j) & "'"
        Next j
        For j = 1 To lcol单价.Count
'            .Formulas(j + 4 * i) = "计量单位" & j & "='" & lcol计量单位(j) & "'"
            Set Item = cr.FormulaFields.GetItemByName("计量单位" & j)
            Item.Text = "'" & lcol计量单位(j) & "'"
        Next j
    'End With
    '计算出小写金额
    For k = 1 To paraCol.Count
        lcur小写金额 = lcur小写金额 + CCur(paraCol(k)("金额") * paraCol(k)("打折比率"))
    Next k
    '转化为大写金额
    lstr大写金额 = FuncConvertToCapsStr(lcur小写金额)
    
    On Error Resume Next
'    With frm票据预览.CrystalReport1
        '修改：2002-6-25（杨春）退费收据要求打印“负”字。
        
'        .Formulas(j + 4 * i + 1) = "小写金额='" & lcur小写金额 & "'"
'        .Formulas(j + 4 * i + 2) = "大写金额='" & lstr大写金额 & "'"
        Set Item = cr.FormulaFields.GetItemByName("小写金额")
        Item.Text = "'" & Format(lcur小写金额, "######0.00") & "'"
        Set Item = cr.FormulaFields.GetItemByName("大写金额")
        Item.Text = "'" & lstr大写金额 & "'"
        
'        .Formulas(j + 4 * i + 3) = "收费批号='" & IIf(IsNull(paraCol(1)("收费批号")), "", paraCol(1)("收费批号")) & "'"
        Set Item = cr.FormulaFields.GetItemByName("收费批号")
        Item.Text = "'" & IIf(IsNull(paraCol(1)("收费批号")), "", paraCol(1)("收费批号")) & "'"
'        .Formulas(j + 4 * i + 4) = "交费人='" & IIf(IsNull(paraCol(1)("交费人")), "", paraCol(1)("交费人")) & "'"
        Set Item = cr.FormulaFields.GetItemByName("交费人")
        Item.Text = "'" & IIf(IsNull(paraCol(1)("交费人")), "", paraCol(1)("交费人")) & "'"
'        .Formulas(j + 4 * i + 5) = "交费单位='" & IIf(IsNull(paraCol(1)("交费单位名称")), "", paraCol(1)("交费单位名称")) & "'"
        Set Item = cr.FormulaFields.GetItemByName("交费单位")
        Item.Text = "'" & IIf(IsNull(paraCol(1)("交费单位名称")), "", paraCol(1)("交费单位名称")) & "'"
'        .Formulas(j + 4 * i + 6) = "主管科室='" & IIf(IsNull(paraCol(1)("主管科室名称")), "", paraCol(1)("主管科室名称")) & "'"
        Set Item = cr.FormulaFields.GetItemByName("主管科室")
        Item.Text = "'" & IIf(IsNull(paraCol(1)("主管科室名称")), "", paraCol(1)("主管科室名称")) & "'"
'        .Formulas(j + 4 * i + 7) = "住院号='" & paraCol(1)("住院号") & "'"
        Set Item = cr.FormulaFields.GetItemByName("住院号")
        Item.Text = "'" & paraCol(1)("住院号") & "'"
'        .Formulas(j + 4 * i + 8) = "年龄='" & paraCol(1)("年龄") & "'"
        Set Item = cr.FormulaFields.GetItemByName("年龄")
        Item.Text = "'" & paraCol(1)("年龄") & "'"
'        .Formulas(j + 4 * i + 9) = "性别='" & paraCol(1)("性别") & "'"
        Set Item = cr.FormulaFields.GetItemByName("性别")
        Item.Text = "'" & paraCol(1)("性别") & "'"
        
        '修改：2002-8-16（杨春）在病种列放入交费方式。
        '.Formulas(j + 4 * i + 10) = "病种='" & paracol(1)("病种") & "'"
        Set lobjRec = dafuncGetData("select 名称 from 收费管理_交费方式字典表 where 编号='" & paraCol(1)("交费方式") & "'")
        Set Item = cr.FormulaFields.GetItemByName("病种")
        If lobjRec.recordcount > 0 Then
            '.Formulas(j + 4 * i + 10) = "病种='" & lobjRec("名称") & "'"
            Item.Text = "'" & lobjRec("名称") & "'"
        Else
            '.Formulas(j + 4 * i + 10) = "病种='现金'"
            Item.Text = "'现金'"
        End If
        
        '.Formulas(j + 4 * i + 11) = "入院日期='" & IIf(IsNull(paraCol(1)("入院日期")), "", paraCol(1)("入院日期")) & "'"
        '.Formulas(j + 4 * i + 12) = "出院日期='" & IIf(IsNull(paraCol(1)("出院日期")), "", paraCol(1)("出院日期")) & "'"
        '.Formulas(j + 4 * i + 13) = "入院操作员='" & IIf(IsNull(paraCol(1)("入院操作员")), "", paraCol(1)("入院操作员")) & "'"
        '.Formulas(j + 4 * i + 14) = "经治医生='" & IIf(IsNull(paraCol(1)("经治医生")), "", paraCol(1)("经治医生")) & "'"
'        .Formulas(j + 4 * i + 11) = "收费人='" & IIf(IsNull(paraCol(1)("收费人姓名")), "", paraCol(1)("收费人姓名")) & "'"
        Set Item = cr.FormulaFields.GetItemByName("收费人")
        Item.Text = "'" & IIf(IsNull(paraCol(1)("收费人姓名")), "", paraCol(1)("收费人姓名")) & "'"
        '.Formulas(j + 4 * i + 16) = "入院操作员='" & IIf(IsNull(paraCol(1)("入院操作员")), "", paraCol(1)("入院操作员")) & "'"
'        .Formulas(j + 4 * i + 12) = "打印日期='" & Date & "'"
        Set Item = cr.FormulaFields.GetItemByName("打印日期")
        Item.Text = "'" & Format(Date, "yyyy-mm-dd") & "'"
        
        '修改：2002-6-25（杨春）增加收据号。
'        .Formulas(j + 4 * i + 13) = "收据号='" & paraCol(1)("收据号") & "'"
        Set Item = cr.FormulaFields.GetItemByName("收据号")
        Item.Text = "'" & paraCol(1)("收据号") & "'"
        
        If paraint打折比率 = 1 Then
'            .Formulas(j + 4 * i + 14) = "打折比率=" & paraCol(1)("打折比率") & ""
            Set Item = cr.FormulaFields.GetItemByName("打折比率")
            Item.Text = "'" & paraCol(1)("打折比率") & "'"
        End If
'    End With
    
    '修改：2001-11-14（杨春）可以执行扩展存储过程获取其他打印项目。
    '约定：扩展存储过程名："收费管理_获取" & 收据名称
    Dim lstrName As String
    
    '获取格式文件名，不带路径和后缀。
    lstrName = parastr票据格式文件
    '去掉路径。
    Do While InStr(lstrName, "\") > 0
        lstrName = Right(lstrName, Len(lstrName) - InStr(lstrName, "\"))
    Loop
    '去掉后缀。
    lstrName = Left(lstrName, InStr(lstrName, ".") - 1)
    Err.Clear
    Set lobjRec = dafuncGetData("收费管理_获取" & lstrName)
    If Err.Number = 0 Then
        '扩展存储过程存在。
        '根据扩展过程的返回结果设置公式。
        If lobjRec.recordcount > 0 Then
            For k = 0 To lobjRec.Fields.Count - 1
'                frm票据预览.CrystalReport1.Formulas(j + 4 * i + 19 + k) = lobjRec(k).Name & "='" & IIf(IsNull(lobjRec(k).Value), "", lobjRec(k).Value) & "'"
                Set Item = cr.FormulaFields.GetItemByName(lobjRec(k).Name)
                Item.Text = "'" & IIf(IsNull(lobjRec(k).Value), "", lobjRec(k).Value) & "'"
            Next
        End If
    End If
    Err.Clear
    On Error GoTo errhandle
    
    '修改：2002-10-17（杨春）原打印预览程序有错误，被修改。
'    frm票据预览.CRViewer91.ReportSource = cr
    
'    If 是否预览 Then
        '预览。
'        frm票据预览.CRViewer91.ViewReport
'        frm票据预览.Show 1
'        With frm票据预览.CrystalReport1
'
'            .WindowState = crptMaximized
'            .WindowControlBox = True
'            .WindowLeft = 0
'            .WindowTop = 0
'            .Destination = crptToWindow
'
'            If paraModal Then
'                .WindowParentHandle = frm票据预览.hWnd
'            Else
'                .WindowParentHandle = 0
'            End If
'
'            .Action = 1
'
'            If paraModal Then
'                frm票据预览.Caption = parastr票据格式文件
'                frm票据预览.Show 1
'            End If
'        End With
'    Else
        cr.PrintOut False
'        '直接打印。
'        With frm票据预览.CrystalReport1
'            .WindowParentHandle = 0
'            .Destination = crptToPrinter
'            .Action = 1
'        End With
'    End If
    
    
    func打印票据分页 = True
    Err.Clear
    Exit Function
errhandle:
    If Err.Number = 20504 Then
        lbln = sffuncMsg("请检查打印格式文件的正确性", sf警告)
        Exit Function
    End If
    If Err.Number = 20507 Then
        lbln = sffuncMsg("请检查打印格式文件的正确性", sf警告)
        Exit Function
    End If
    If Err.Number = 20500 Then
        
        Exit Function
    End If
    Call sfsub错误处理("收费数据对象", "cls票据信息", "func打印票据分页", Err.Number, Err.Description, True)
    'sffuncMsg Err.Description, sf警告
    Exit Function
    Resume
End Function



