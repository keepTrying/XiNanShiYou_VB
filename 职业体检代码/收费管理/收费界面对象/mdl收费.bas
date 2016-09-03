Attribute VB_Name = "mdl收费"
'Public pint打折控制 As Integer '保持业务配置的打折控制信息
'Public pint科目级数 As Integer '保持业务配置的科目级数

Public pobj业务设置 As Object
Public pobj收费管理 As Object
Public pobj单位定位 As Object  '单位档案接口


Public Sub Main()
    Set pobj收费管理 = CreateObject("收费对象部件.cls收费管理")
    Set pobj单位定位 = CreateObject("单位档案业务.ClsUnitInterface")

End Sub


Public Function func录入票据号() As String
    Dim lstr票据号 As String
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select 当前值 from 系统管理_系统编号生成记录表 where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'")
    If lobjRec.RecordCount = 0 Then
        dafuncGetData "insert into 系统管理_系统编号生成记录表(业务名称,编号名称,数据类型,当前值,长度,是否按年重编,当前年号) values('收费管理" & um用户编号 & "','收据号','C',0,9,'否',2008)"
        Set lobjRec = dafuncGetData("select 当前值 from 系统管理_系统编号生成记录表 where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'")
    End If
    lstr票据号 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    lLen = Len(lstr票据号)
    lstr票据号 = Format(Val(lstr票据号) + 1, String(lLen, "0"))
    lstr票据号 = InputBox("请输入当前票据号：", "系统询问", lstr票据号)
    
    Do While lstr票据号 <> ""
        lLen = Len(lstr票据号)
        Set lobjRec = dafuncGetData("select 收据号 from 收费管理_费用信息表 where 收据号='" & lstr票据号 & "'")
        If lobjRec.RecordCount Then
            MsgBox "该票据号已经使用了，不能再次使用。", vbCritical, "系统提示"
            lstr票据号 = InputBox("请输入当前票据号：", "系统询问", lstr票据号)
        Else
            lstr票据号 = Format(Val(lstr票据号) - 1, String(lLen, "0"))
            dafuncGetData "update 系统管理_系统编号生成记录表 set 当前值='" & lstr票据号 & "' where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'"
            Exit Do
        End If
    Loop
    
    func录入票据号 = lstr票据号
End Function


'功能: 将给定金额转换为人民币的大写字符串
'输入: money       金额
'输出: FuncConvertToCapsStr     转换的大写字符串
'最后修改时间: 96.6.11
'--------------------------------------------------
Public Function FuncConvertToCapsStr(Money As Currency) As String
On Error GoTo errhandle
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
    Exit Function
errhandle:
    sfsub错误处理 "收费界面部件", "mdl收费", " FuncConvertToCapsStr()", Err.Number, Err.Description, True
End Function



Public Function gfuncKeyNum(paraKey As Integer) As Integer
    On Error Resume Next
    Select Case paraKey
        Case 8, 13, 46, 48 To 57
        ' 接收退格键(ASCII码为8)、回车键(ASCII码为13)、小数点和数字键(ASCII码为48～57)
            gfuncKeyNum = paraKey

        Case Else
            gfuncKeyNum = 0 ' 其他键不接收，用ASCII码为0的字符表示不接收输入信息
    End Select
End Function
