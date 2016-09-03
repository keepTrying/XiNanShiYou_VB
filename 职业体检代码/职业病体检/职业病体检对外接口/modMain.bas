Attribute VB_Name = "modMain"
Option Explicit

'功能：从paraRange[数据范围名，数据范围值] 集合的值分析出导入或导出的条件,并拼成一个SQL的条件语句.
'输入：paraRange As Collection [数据范围名，数据范围值]，数据范围名：起始日期，开始日期，单位名称集，从系统编号,到系统编号。
'       paraInput (true 导入/false 导出)
'返回：以and 开始的string.  如"and 体检日期 >= convert(datetime,'2001-3-1') and 单位名称 in ('单位A','单位b')
'作者：刘浩。
Public Function funcFilter(ByVal paraRange As Collection, Optional paraInput As Boolean = True) As String
    Dim lstrSQL As String
    Dim lstrUnit As String      '存放单位编号集字串,用于拼SQL语句
    Dim lintUnit As Integer     '半角逗号在字串中的位置,用于拼SQL语句的单位编号(每个带单引号的编号用逗号分隔)
    
    On Err GoTo errHandler
    
    If Not paraRange Is Nothing Then
        '开始日期。
        If sffunc判断集合键值是否存在(paraRange, "开始日期") Then
            If paraRange("开始日期")("数据范围值") <> "" Then
                If paraInput Then
                    lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.体检日期 >= #" & paraRange("开始日期")("数据范围值") & "#"
                Else
                    lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.体检日期 >= '" & paraRange("开始日期")("数据范围值") & "'"
                End If
            End If
        End If
        '结束日期。
        If sffunc判断集合键值是否存在(paraRange, "结束日期") Then
            If paraRange("结束日期")("数据范围值") <> "" Then
                If paraInput Then
                    lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.体检日期 <= #" & paraRange("结束日期")("数据范围值") & "#"
                Else
                    lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.体检日期 <= '" & paraRange("结束日期")("数据范围值") & "'"
                End If
            End If
        End If
        '单位名称集。
        If sffunc判断集合键值是否存在(paraRange, "单位名称集") Then
            lstrUnit = Trim(paraRange("单位名称集")("数据范围值"))
            If lstrUnit <> "" Then
                lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.单位名称 in ("
                lstrUnit = lstrUnit & IIf(Right(lstrUnit, 1) <> ",", ",", "")
                While Len(lstrUnit) > 0
                    lintUnit = InStr(1, lstrUnit, ",")
                    lstrSQL = lstrSQL & "'" & Left(lstrUnit, lintUnit - 1) & "',"
                    lstrUnit = Right(lstrUnit, Len(lstrUnit) - lintUnit)
                Wend
                lstrSQL = Left(lstrSQL, Len(lstrSQL) - 1) & ")"
            End If
        End If
        '系统编号范围。
        If sffunc判断集合键值是否存在(paraRange, "从系统编号") Then
            If paraRange("从系统编号")("数据范围值") <> "" Then
                lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.系统编号 >= '" & paraRange("从系统编号")("数据范围值") & "'"
            End If
        End If
        If sffunc判断集合键值是否存在(paraRange, "到系统编号") Then
            If paraRange("到系统编号")("数据范围值") <> "" Then
                lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.系统编号 <= '" & paraRange("到系统编号")("数据范围值") & "'"
            End If
        End If
        '体检对象(体检表名称)
        If sffunc判断集合键值是否存在(paraRange, "体检对象") Then
            lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.体检表名称 = '" & paraRange("体检对象")("数据范围值") & "'"
        End If
        '已体检完毕
        If sffunc判断集合键值是否存在(paraRange, "已体检完毕") Then
            If paraRange("已体检完毕")("数据范围值") Then
                lstrSQL = lstrSQL & " and 体检管理_体检基本数据库.体检状态 = " & P_ENDED_STATUS
            End If
        End If
        
    End If
    funcFilter = lstrSQL
    Exit Function
errHandler:
    sfsub错误处理 "体检对外接口部件", "ClsManngeTransmission", "funcFilter", Err.Number, Err.Description, True
End Function

