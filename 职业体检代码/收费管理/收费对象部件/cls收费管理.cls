VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls收费管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobj票据信息 As cls票据信息

Public Function func收费管理界面查询(ByVal para收费批号 As String, ByVal para收据号 As String, _
                                    ByVal para单位名称 As String, ByVal para交费人 As String, _
                                    ByVal para开始日期 As String, ByVal para截止日期 As String, _
                                    ByVal para业务分类 As String, para次数 As Object) As Object
                                    
    Set func收费管理界面查询 = dafuncGetData("exec 收费管理_返回收费信息 '" & para收费批号 & "','" & para收据号 & "','" & para交费人 & "','" & para单位名称 & "','" & para开始日期 & "','" & para截止日期 & "','" & para业务分类 & "','" & um用户编号 & "'")
    '默认只能查询自己的收费信息
    Set para次数 = dafuncGetData("exec 收费管理_返回费用次数 '" & para收费批号 & "','" & para收据号 & "','" & para交费人 & "','" & para单位名称 & "','" & para开始日期 & "','" & para截止日期 & "','" & para业务分类 & "','" & um用户编号 & "'")
                                    
End Function

Public Function func财务监管界面查询(ByVal para收据号 As String, ByVal para收费员 As String, ByVal para开始日期 As String, ByVal para截止日期 As String) As Object
                                    
    Set func财务监管界面查询 = dafuncGetData("exec 收费管理_返回收费信息 '','" & para收据号 & "','','','" & para开始日期 & "','" & para截止日期 & "','','" & para收费员 & "'")
                                    
End Function

Public Sub sub退费(ByVal para收费批号 As String, ByVal para退费人 As String, ByVal para退费日期 As String)
    Dim lobjRec As Object
    '判断收费状态.
    Set lobjRec = dafuncGetData("select 收费状态 from 收费管理_费用信息表  where 收费编号='" & para收费批号 & "'")
    If lobjRec.recordcount > 0 Then
        If IIf(IsNull(lobjRec!收费状态), 0, lobjRec!收费状态) = 2 Then
            Err.Raise 6666, , "该记录已退费，请刷新记录！"
        End If
        dafuncGetData "update 收费管理_费用信息表 set 收费状态=2,退费人='" & para退费人 & "',退费日期='" & para退费日期 & "' where 收费编号='" & para收费批号 & "'"
    End If
End Sub


Public Sub sub打印退费票据(ByVal para收费批号 As String, ByVal para预览 As Boolean)
    Dim lcol费用打印信息集 As Collection    '存放费用信息的集合
    Dim lcol费用信息 As Collection         '将费用信息所有字段信息写入集合中
    Dim lstr格式文件名 As String           '记录打印格式的文件名
    Dim lrec格式文件名对象 As Object           '记录打印格式的文件名对象
    Dim lrec费用信息 As Object             '记录详的费用信息
    Dim lrec费用票据信息 As Object         ' 记录与票据有关的信息
    Dim k As Long                         '循环变量
    Dim lstr交费人 As String               '记录交费人姓名
    Dim lstr交费单位 As String             ' 记录交费单位姓名
    Dim lsge打折比率 As Single            '记录打折比率
    Dim lsge金额 As Single                '记录打金额
    
    Dim lobj汇总记录 As Object
    
    On Error GoTo errHanler
    
    '获取与费用信息相关的票据信息
    Set lrec费用票据信息 = dafuncGetData("select b.票据类型编号,c.对应业务 from 收费管理_收费项目字典表 b, 收费管理_费用信息表 c " & _
                            "Where b.收费项目编号 = c.收费项目编号 and c.收费编号 ='" & _
                            para收费批号 & "' group by b.票据类型编号,对应业务")
    '校检与费用信息相关的票据信息
    If (lrec费用票据信息.EOF) Then
        Err.Raise 6666, , "未检索到收费项目的票据类型信息,无法进行打印！"
    End If
    
    lrec费用票据信息.MoveFirst
    
    
    '按票据类型取出费用信息
    Do While Not lrec费用票据信息.EOF
        '获取打印费用信息
        Set lrec费用信息 = dafuncGetData("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & "  and 对应业务='" & lrec费用票据信息("对应业务") & "' and 收费编号='" & para收费批号 & "'")
        '校检费用信息
        If (lrec费用信息.EOF) Then
            Err.Raise 6666, , "无可打印信息！"
        End If
        
        '处理费用信息中交费人和交费单位为空值的情况
        If IIf(IsNull(lrec费用信息("交费单位名称").Value), "", lrec费用信息("交费单位名称")) <> "" Then
            lstr交费单位 = lrec费用信息("交费单位名称").Value
        Else
            lstr交费单位 = ""
        End If
        If IIf(IsNull(lrec费用信息("交费人").Value), "", lrec费用信息("交费人")) <> "" Then
            lstr交费人 = lrec费用信息("交费人").Value
        Else
            lstr交费人 = ""
        End If
        '初始化打折比率值
        lsge打折比率 = 1
        Set lcol费用打印信息集 = New Collection
        
        '修改：2002-9-29（杨春）合并打印。
        Set lobj汇总记录 = dafuncGetData("select 收费项目编号,单价=avg(单价),数量=sum(数量),金额=sum(金额) from 收费管理_打印费用信息 " _
                        & "where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费编号='" & para收费批号 _
                        & "' group by 收费批号,收费项目编号")
        
        '将费用信息加入到合对象中
        Do While Not lobj汇总记录.EOF
            '修改：2002-9-29（杨春）获取当前项目的详细信息。
            Set lrec费用信息 = dafuncGetData("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费编号='" & para收费批号 & "' and 收费项目编号='" & lobj汇总记录("收费项目编号") & "'")
            
            Set lcol费用信息 = New Collection
            For k = 0 To lrec费用信息.Fields.Count - 1
                If lrec费用信息.Fields(k).Name = "交费单位名称" Or lrec费用信息.Fields(k).Name = "交费人" Or lrec费用信息.Fields(k).Name = "打折比率" Or lrec费用信息.Fields(k).Name = "金额" Then
                    If lrec费用信息.Fields(k).Name = "交费单位名称" Then lcol费用信息.Add lstr交费单位, "交费单位名称"
                    If lrec费用信息.Fields(k).Name = "交费人" Then lcol费用信息.Add lstr交费人, "交费人"
                    If lrec费用信息.Fields(k).Name = "打折比率" Then
                        lsge打折比率 = lrec费用信息(k).Value
                        lcol费用信息.Add lsge打折比率, "打折比率"
                    End If
                ElseIf lrec费用信息.Fields(k).Name <> "单价" And lrec费用信息.Fields(k).Name <> "数量" And lrec费用信息.Fields(k).Name <> "金额" Then
                    '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
                    lcol费用信息.Add lrec费用信息(k).Value, lrec费用信息.Fields(k).Name
                End If
            Next k
            
            '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
            lcol费用信息.Add Format(lobj汇总记录("单价").Value, "0.00"), "单价"
            lcol费用信息.Add lobj汇总记录("数量").Value, "数量"
            lcol费用信息.Add Format(0 - lobj汇总记录("金额").Value, "0.00"), "金额"
            
            lcol费用信息.Add "年龄值", "年龄"
            lcol费用信息.Add "性别值", "性别"
            lcol费用信息.Add "住院号值", "住院号"
            lcol费用信息.Add "病种值", "病种"
            lcol费用信息.Add "2002", "入院日期"
            lcol费用信息.Add "2002", "出院日期"
            lcol费用信息.Add "入院操作员值", "入院操作员"
            lcol费用信息.Add "经治医生值", "经治医生"
            
            lcol费用打印信息集.Add lcol费用信息
            
            lobj汇总记录.movenext
        Loop
        
        '获取格式文件名
        Set lrec格式文件名对象 = dafuncGetData("select * from 收费管理_票据设置信息表 where 票据类型编号='" & lrec费用票据信息("票据类型编号") & "'  and 对应业务='" & lrec费用票据信息("对应业务") & "'")
        If lrec格式文件名对象.EOF Then
            Err.Raise 6666, , "未查找到对应的票据设置信息！"
        End If
        
        lstr格式文件名 = lrec格式文件名对象("票据格式文件名称")
        
        Set mobj票据信息 = New cls票据信息
        mobj票据信息.项目汇总 = IIf(IsNull(lrec格式文件名对象!项目汇总), "否", lrec格式文件名对象!项目汇总)
        mobj票据信息.是否预览 = para预览
        mobj票据信息.func打印票据 App.Path & "\" & lstr格式文件名, lcol费用打印信息集, lsge打折比率, lrec格式文件名对象("最大项数").Value, True, True
        
        lrec费用票据信息.movenext
    Loop

    Exit Sub
errHanler:
    Call sfsub错误处理("收费对象部件", "cls收费管理", "sub打印退费票据", Err.Number, Err.Description, True)
    Exit Sub
    Resume
End Sub



Public Sub sub打印票据(ByVal para收费批号 As String, ByVal para预览 As Boolean, _
                        Optional paraParentHwnd As Long = 0)
                        
    Dim lcol费用信息 As Collection         '将费用信息所有字段信息写入集合中
    Dim lcol费用打印信息集 As Collection   '存放费用信息的集合
    Dim lrec查找记录 As Object             '存放查询出的费用信息
    Dim lstr格式文件名 As String           '记录打印格式的文件名
    Dim lrec格式文件名对象 As Object           '记录打印格式的文件名对象
    Dim lrec费用信息 As Object             '记录详的费用信息
    Dim lrec费用票据信息 As Object         ' 记录与票据有关的信息
    Dim i As Long                         '循环变量
    Dim j As Long                         '循环变量
    Dim k As Long                         '循环变量
    Dim lstr交费人 As String               '记录交费人姓名
    Dim lstr交费单位 As String             ' 记录交费单位姓名
    Dim lsge打折比率 As Single            '记录打折比率
    Dim lobj汇总记录 As Object
    Dim lstr收据号 As String
    Dim lobjRec As Object
    On Error GoTo errhandle
    
    '获取与费用信息相关的票据信息
    Set lrec费用票据信息 = dafuncGetData("select b.票据类型编号,c.对应业务 from 收费管理_收费项目字典表 b, 收费管理_费用信息表 c " & _
                            "Where b.收费项目编号 = c.收费项目编号 and c.收费批号 ='" & _
                            para收费批号 & "' group by b.票据类型编号,c.对应业务")
    '校检与费用信息相关的票据信息
    If (lrec费用票据信息.EOF) Then
        Err.Raise 6666, , "未检索到收费项目的票据类型信息，无法进行打印！"
        
    End If
    
'    '生成起始收据号。
'    Set lobjRec = dafuncGetData("收费管理_产生收据号")
'    If lobjRec.recordcount > 0 Then
'        lstr收据号 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
'    End If
'    dafuncGetData "update 收费管理_费用信息表 set 收据号='" & lstr收据号 & "',票据张数=1 where 收费批号='" & para收费批号 & "'"
    
    lrec费用票据信息.MoveFirst
    
    '按票据类型取出费用信息
    For i = 0 To lrec费用票据信息.recordcount - 1

        '获取打印费用信息
        Set lrec费用信息 = dafuncGetData("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 对应业务='" & lrec费用票据信息("对应业务") & "' and 收费批号='" & para收费批号 & "'")
        '校检费用信息
        If lrec费用信息.EOF Then
            Err.Raise 6666, , "无可打印信息！"
        End If
        
'        '累计票据号。
'        If i > 0 Then
'            Call dafuncGetData("收费管理_产生收据号")
'            dafuncGetData "update 收费管理_费用信息表 set 票据张数=票据张数+1 where 收费批号='" & para收费批号 & "'"
'        End If
        
        '处理费用信息中交费人和交费单位为空值的情况
        If IIf(IsNull(lrec费用信息("交费单位名称").Value), "", lrec费用信息("交费单位名称")) <> "" Then
            lstr交费单位 = lrec费用信息("交费单位名称").Value
        Else
            lstr交费单位 = ""
        End If
        If IIf(IsNull(lrec费用信息("交费人").Value), "", lrec费用信息("交费人")) <> "" Then
            lstr交费人 = lrec费用信息("交费人").Value
        Else
            lstr交费人 = ""
        End If
        '初始化打折比率值
        lsge打折比率 = 1
        Set lcol费用打印信息集 = New Collection
        
        '修改：2002-9-29（杨春）合并打印。
        Set lobj汇总记录 = dafuncGetData("select 收费项目编号,单价=avg(单价),数量=sum(数量),金额=sum(金额) from 收费管理_打印费用信息 " _
                            & "where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & para收费批号 _
                            & "' group by 收费批号,收费项目编号")
        
        '将费用信息加入到合对象中
        For j = 0 To lobj汇总记录.recordcount - 1
            '修改：2002-9-29（杨春）获取当前项目的详细信息。
            Set lrec费用信息 = dafuncGetData("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & para收费批号 & "' and 收费项目编号='" & lobj汇总记录("收费项目编号") & "'")
            
            Set lcol费用信息 = New Collection
            For k = 0 To lrec费用信息.Fields.Count - 1
                If lrec费用信息.Fields(k).Name = "交费单位名称" Or lrec费用信息.Fields(k).Name = "交费人" Or lrec费用信息.Fields(k).Name = "打折比率" Then
                    If lrec费用信息.Fields(k).Name = "交费单位名称" Then lcol费用信息.Add lstr交费单位, "交费单位名称"
                    If lrec费用信息.Fields(k).Name = "交费人" Then lcol费用信息.Add lstr交费人, "交费人"
                    If lrec费用信息.Fields(k).Name = "打折比率" Then
                        lsge打折比率 = lrec费用信息(k).Value
                        lcol费用信息.Add lsge打折比率, "打折比率"
                    End If
                ElseIf lrec费用信息.Fields(k).Name <> "单价" And lrec费用信息.Fields(k).Name <> "数量" And lrec费用信息.Fields(k).Name <> "金额" Then
                    '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
                    lcol费用信息.Add lrec费用信息(k).Value, lrec费用信息.Fields(k).Name
                End If
            Next k
            '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
            lcol费用信息.Add Format(lobj汇总记录("单价").Value, "0.00"), "单价"
            lcol费用信息.Add lobj汇总记录("数量").Value, "数量"
            lcol费用信息.Add Format(lobj汇总记录("金额").Value, "0.00"), "金额"
            
           
            lcol费用打印信息集.Add lcol费用信息
            
            'If Not lrec费用信息.EOF Then lrec费用信息.MoveNext
            If Not lobj汇总记录.EOF Then lobj汇总记录.movenext
        Next j
        '获取格式文件名
        Set lrec格式文件名对象 = dafuncGetData("select * from 收费管理_票据设置信息表 where 票据类型编号='" & lrec费用票据信息("票据类型编号") & "' and 对应业务='" & lrec费用票据信息("对应业务") & "'")
        If lrec格式文件名对象.EOF Then
            Err.Raise 6666, , "未查找到票据格式设置信息！"
        End If
        lstr格式文件名 = lrec格式文件名对象("票据格式文件名称")
'Debug.Print lstr格式文件名

        Set mobj票据信息 = New cls票据信息
        mobj票据信息.项目汇总 = IIf(IsNull(lrec格式文件名对象!项目汇总), "否", lrec格式文件名对象!项目汇总)
        mobj票据信息.是否预览 = para预览
        mobj票据信息.func打印票据 App.Path & "\" & lstr格式文件名, lcol费用打印信息集, lsge打折比率, lrec格式文件名对象("最大项数").Value, False, True
        
        lrec费用票据信息.movenext
    Next i
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费对象部件", "cls收费管理", "sub打印票据", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

Public Property Let 业务设置(ByVal para项目名称 As String, ByVal para设置值 As String)
    dafuncGetData ("update 收费管理_业务配置表 set 设置值='" & para设置值 & "' where 设置项目='" & para项目名称 & "'")
End Property

Public Property Get 业务设置(ByVal para项目名称 As String) As String
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    业务设置 = ""
    Set lobjRec = dafuncGetData("select * from 收费管理_业务配置表 where 设置项目='" & para项目名称 & "'")
    If lobjRec.recordcount > 0 Then
        业务设置 = IIf(IsNull(lobjRec!设置值), "", lobjRec!设置值)
    End If

    Exit Property
errHandler:

End Property


'根据收费批号删除费用信息.
Public Sub sub报废费用信息(ByVal para收费批号 As String)
    On Error GoTo errHandler
    Dim lstrSql As String           '定义变量记录SQL语句
    dafuncGetData ("update 收费管理_费用信息表 set 收费状态=3 where 收费编号='" & para收费批号 & "'")
    Exit Sub
errHandler:
    sfsub错误处理 "收费对象部件", "cls收费管理", "sub报废费用信息", Err.Number, Err.Description, True

End Sub

'根据收费批号删除费用信息.
Public Sub sub取消报废费用信息(ByVal para收费批号 As String)
    On Error GoTo errHandler
    Dim lstrSql As String           '定义变量记录SQL语句
    dafuncGetData ("update 收费管理_费用信息表 set 收费状态=1 where 收费编号='" & para收费批号 & "'")
    Exit Sub
errHandler:
    sfsub错误处理 "收费对象部件", "cls收费管理", "sub取消报废费用信息", Err.Number, Err.Description, True

End Sub

Public Sub sub删除(ByVal para收费批号 As String)
    dafuncGetData ("delete from 收费管理_费用信息表 where 收费编号='" & para收费批号 & "'")

End Sub


'返回：收费编号。
Public Function func划价保存(ByVal para费用信息 As Collection, _
                            Optional para收费编号 As String = "", _
                            Optional para业务分类 As String = "") As String

    Dim lstr收费编号 As String
    Dim lobjTemp As Object
    Dim i As Long
    Dim lstrSql As String
    
    On Error GoTo errhandle
    func划价保存 = ""
    
    '校验参数
    If para费用信息 Is Nothing Then Exit Function
    
    '开启事务
    dasubBeginTran
    
    '判断操作是否为"修改",如果是修改就取出要修改的费用信息,并将之删除,以便在以后添加新的费用信息
    If para收费编号 <> "" Then
        '修改时，判断是否已收费。
        lstr收费编号 = para收费编号
        
        '取出满足条件"收费编号=lstr收费编号"的记录
        Set lobjTemp = dafuncGetData("select * from 收费管理_费用信息表 where 收费编号='" & lstr收费编号 & "'")
        If Not lobjTemp.EOF Then
            '判断该批费用是否已经交费,如已交费则不允许修改
            If lobjTemp("收费状态") = 1 Then
                Err.Raise 6666, , "该批费用已经交费，不能修改！"
                Exit Function
            Else
                '删除原来的费用信息
                dafuncGetData "delete from 收费管理_费用信息表 where 收费编号='" & lstr收费编号 & "'"
            End If
        End If
    Else
        '生成收费编号。
retry:        Set lobjTemp = dafuncGetData("收费管理_产生收费批号 '" & um服务器代号 & "'")
        lstr收费编号 = lobjTemp(0)
    End If

    '向数据库中新增费用信息
    For i = 1 To para费用信息.Count
        lstrSql = "insert into 收费管理_费用信息表(收费批号,收费编号,交费人,交费单位编号,交费单位名称,主管科室经手人," _
                & "主管科室编号,收费项目编号,收费项目名称,单价,数量,金额,打折比率,交费方式,业务分类,备注1,备注2,收费状态) values("
        lstrSql = lstrSql & "'" & lstr收费编号 & "','" & lstr收费编号 & "',"
        lstrSql = lstrSql & "'" & para费用信息(i)("交费人") & "',"
        lstrSql = lstrSql & "'" & para费用信息(i)("交费单位编号") & "',"
        lstrSql = lstrSql & "'" & para费用信息(i)("交费单位名称") & "',"
        lstrSql = lstrSql & "'" & para费用信息(i)("主管科室经手人") & "',"
        lstrSql = lstrSql & "'" & para费用信息(i)("主管科室编号") & "',"
        lstrSql = lstrSql & "'" & para费用信息(i)("收费项目编号") & "',"
        
        If Not sffunc判断集合键值是否存在(para费用信息(i), "收费项目名称") Then
            lstrSql = lstrSql & "'" & func获取收费项目名称(para费用信息(i)("收费项目编号")) & "',"
        Else
            lstrSql = lstrSql & "'" & para费用信息(i)("收费项目名称") & "',"
        End If
        lstrSql = lstrSql & para费用信息(i)("单价") & ","
        lstrSql = lstrSql & para费用信息(i)("数量") & ","
        lstrSql = lstrSql & Format(Val(para费用信息(i)("数量")) * Val(para费用信息(i)("单价")), "0.0000") & ","
        lstrSql = lstrSql & "1,0,'" & para业务分类 & "',"
        If sffunc判断集合键值是否存在(para费用信息(i), "备注1") Then
            lstrSql = lstrSql & "'" & para费用信息(i)("备注1") & "',"
        Else
            lstrSql = lstrSql & "'',"
        End If
        If sffunc判断集合键值是否存在(para费用信息(i), "备注2") Then
            lstrSql = lstrSql & "'" & para费用信息(i)("备注2") & "',"
        Else
            lstrSql = lstrSql & "'',"
        End If
        lstrSql = lstrSql & "0)"
        
        dafuncGetData lstrSql
    Next
    
    '提交事务
    dasubCommitTran
        
    '返回结果
    func划价保存 = lstr收费编号
    
    Exit Function
    
errhandle:
    If Err.Number = -2147217873 Then
        '收费编号重复.
        Err.Clear
        GoTo retry
    End If
    sfsub错误处理 "收费对象部件", "cls收费管理", "func划价保存", Err.Number, Err.Description, True
    Exit Function
    Resume
End Function
Private Function func获取收费项目名称(ByVal para编号 As String) As String
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select 收费项目名称 from 收费管理_收费项目字典表 where 收费项目编号='" & para编号 & "'")
    If lobjRec.recordcount > 0 Then
        func获取收费项目名称 = lobjRec(0)
    Else
        func获取收费项目名称 = ""
    End If
    
End Function

'输入:para收费编号集
'     para确认信息：打折比率，收费方式，交费日期，收费人
Public Sub sub收费确认(ByVal para收费编号集 As Collection, ByVal para确认信息 As Collection)
    Dim lobjRec As Object
    Dim lstr收据号 As String
    Dim i As Long                   '循环变量
    
    On Error GoTo errhandle
    
    dasubBeginTran

    '保存收费确认信息。
    For i = 1 To para收费编号集.Count
        dafuncGetData "update 收费管理_费用信息表 set 收费批号='" & para收费编号集(1) & "', " _
                & "打折比率=" & para确认信息("打折比率") & ",交费方式=" & para确认信息("收费方式") & "," _
                & "交费日期='" & para确认信息("交费日期") & "',收费人='" & para确认信息("收费人") & "'," _
                & "开户银行='" & para确认信息("开户银行") & "',对应业务='" & para确认信息("对应业务") & "'," _
                & "收费状态=1 where 收费编号='" & para收费编号集(i) & "'"
    Next
    
    dasubCommitTran
    Exit Sub

errhandle:
    sfsub错误处理 "收费对象部件", "cls收费管理", "func收费", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

Public Sub sub删除收费项目(ByVal para编号 As String)
     dafuncGetData ("exec 收费管理_删除收费项目 '" & para编号 & "'")

End Sub

Public Sub sub删除打折信息(ByVal para单位编号 As String)
    dafuncGetData "exec 收费管理_删除打折信息 '" & para单位编号 & "'"
End Sub

Public Sub sub保存打折信息(ByVal para单位编号 As String, ByVal para打折比率 As Single)
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select * from 收费管理_打折信息表 where 单位编号='" & para单位编号 & "'")
    If lobjRec.recordcount > 0 Then
        dafuncGetData "update 收费管理_打折信息表 set 打折比率=" & para打折比率 & " where 单位编号='" & para单位编号 & "'"
    Else
        dafuncGetData "insert into 收费管理_打折信息表(单位编号,打折比率) values('" & para单位编号 & "'," & para打折比率 & ")"
    End If
    
End Sub


Public Sub sub保存票据设置(ByVal para编号 As String, ByVal para票据名称 As String, _
                            ByVal para格式文件 As String, ByVal para票据类型编号 As Integer, _
                            ByVal para对应业务 As String, ByVal para最大项数 As Integer, _
                            Optional para项目汇总 As String = "否")
    Dim lobjRec As Object
    If para编号 <> "" Then
        '修改。
        dafuncGetData "update 收费管理_票据设置信息表 set 票据名称='" & para票据名称 & "',票据类型编号=" & para票据类型编号 _
                    & ",票据格式文件名称='" & para格式文件 & "',对应业务='" & para对应业务 & "',最大项数=" & para最大项数 _
                    & ",项目汇总='" & para项目汇总 & "' where 编号='" & para编号 & "'"
    Else
        '插入，生成编号。
        Set lobjRec = dafuncGetData("select max(编号) from 收费管理_票据设置信息表")
        para编号 = Format(Val(IIf(IsNull(lobjRec(0)), "0", lobjRec(0))) + 1, "00")
        dafuncGetData "insert into 收费管理_票据设置信息表(编号,票据类型编号,票据名称,票据格式文件名称,对应业务,最大项数,项目汇总) values('" _
                & para编号 & "'," & para票据类型编号 & ",'" & para票据名称 & "','" & para格式文件 & "','" & para对应业务 & "'," & para最大项数 & ",'" & para项目汇总 & "')"
    End If
End Sub


Public Function func划价_收费标准(ByVal para收费标准 As String, ByVal para附加信息 As Collection, Optional para收费编号 As String = "", Optional para业务分类 As String = "")
    Dim lcol费用信息 As Collection
    Dim lrds费用信息 As Object
    Dim lstr收费编号 As String
    Dim lcur总金额 As Currency
    
    '开启错误捕获
    On Error GoTo errhandle
    
    func划价_收费标准 = ""
    
    '根据收费标准名称获取费用项目
    Set lrds费用信息 = dafuncGetData("select a.*,b.收费项目名称 from 收费管理_收费标准信息表 a,收费管理_收费项目字典表 b where a.收费项目编号=b.收费项目编号 and a.收费标准名称= '" & para收费标准 & "'")
    If lrds费用信息.EOF Then
        Err.Raise 6666, , "该标准不存在！"
    End If
    
    Set lcol费用信息 = New Collection
    Dim lcolItem As Collection
    lrds费用信息.MoveFirst
    Do While Not lrds费用信息.EOF
    
        Set lcolItem = New Collection
        With lcolItem
            .Add IIf(para附加信息("交费人") = "", "", para附加信息("交费人")), "交费人"
            If sffunc判断集合键值是否存在(para附加信息, "交费单位编号") Then
                .Add IIf(para附加信息("交费单位编号") = "", "", para附加信息("交费单位编号")), "交费单位编号"
            Else
                .Add "", "交费单位编号"
            End If
            
            .Add IIf(para附加信息("交费单位名称") = "", "", para附加信息("交费单位名称")), "交费单位名称"
            .Add para附加信息("主管科室经手人"), "主管科室经手人"
            .Add para附加信息("主管科室编号"), "主管科室编号"
            
            '基本信息
            .Add lrds费用信息("收费项目编号").Value, "收费项目编号"
            .Add lrds费用信息("收费项目名称").Value, "收费项目名称"
            .Add lrds费用信息("单价").Value, "单价"
            .Add lrds费用信息("数量").Value, "数量"
            
            If sffunc判断集合键值是否存在(para附加信息, "备注1") Then
                .Add para附加信息("备注1"), "备注1"
            End If
            If sffunc判断集合键值是否存在(para附加信息, "备注2") Then
                .Add para附加信息("备注2"), "备注2"
            End If
            
        End With
        lcol费用信息.Add lcolItem
        
        lrds费用信息.movenext
    Loop
    
    lstr收费编号 = func划价保存(lcol费用信息, para收费编号, para业务分类)
    
    func划价_收费标准 = lstr收费编号
    
    Exit Function
    
errhandle:
    
    sfsub错误处理 "收费对象部件", "cls收费管理", "func划价_收费标准", Err.Number, Err.Description, True
    
End Function

Public Function func查询收费项目(ByVal para级数 As Integer) As Object
    Dim lint科目级数 As Integer
    
    Set func查询收费项目 = dafuncGetData("select * from 收费管理_收费项目字典表 where Len(收费项目编号) =" & para级数 * 3 & " order by 收费项目编号")

End Function

Public Function func查询费用明细(ByVal para收费批号 As String) As Object
    Set func查询费用明细 = dafuncGetData("exec 收费管理_获取费用明细 '" & para收费批号 & "'")
    
End Function
