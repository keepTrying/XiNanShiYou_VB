VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls对外接口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'返回：收费编号。
'修改：2002-6-23（增加参数：para申请编号,para业务分类）。
Public Function func划价_数据集合(para费用信息 As Collection, _
                                Optional ByVal para收费编号 As String, _
                                Optional ByVal para显示界面 As Boolean = False, _
                                Optional ByVal para申请编号 As String = "", _
                                Optional ByVal para业务分类 As String = "", _
                                Optional ByVal paraTop As Long = 350) As String
    Dim lobj使用对象 As Object
    Dim lobjRec As Object
    Dim lstr收费编号 As String
    
    If para显示界面 Then
        Set lobj使用对象 = CreateObject("收费界面部件.cls界面管理")
        
        '修改：2002-6-23（可以修改划价，但是要判断状态）。
        lstr收费编号 = ""
        If para收费编号 <> "" Then
            Set lobjRec = dafuncGetData("select * from 收费管理_费用信息表 where 收费编号='" & para收费编号 & "'")
            If lobjRec.RecordCount > 0 Then
                Select Case lobjRec("收费状态")
                Case 0
                    '未收费
                    lstr收费编号 = para收费编号
                Case 1
                    '已收费。不允许修改。
                    Err.Raise 6666, , "收费编号为 “" & para收费编号 & "” 的收费记录已收费，不允许修改。" & Chr(13) & Chr(10) & "若要修改，请收费科室先作退费处理。"
                Case 2
                    '已退费。新划价。
                End Select
            End If
        End If
        
        'Call lobj使用对象.funcstart("收费管理_划价", lstr收费编号)
        func划价_数据集合 = lobj使用对象.func划价(lstr收费编号, para申请编号, para业务分类, paraTop)
        
    Else
        Set lobj使用对象 = CreateObject("收费对象部件.cls收费管理")
        func划价_数据集合 = lobj使用对象.func划价保存(para费用信息, para收费编号, para业务分类)
    End If
    
    Set lobj使用对象 = Nothing
End Function

'修改：2002-6-23（杨春）增加参数“para业务分类”。
Public Function func划价_收费标准(para收费标准 As String, _
                                para附加信息 As Collection, _
                                Optional parastr收费编号 As String, _
                                Optional para业务分类 As String = "") As String
    Dim lobj使用对象 As Object
    
    Set lobj使用对象 = CreateObject("收费对象部件.cls收费管理")
    func划价_收费标准 = lobj使用对象.func划价_收费标准(para收费标准, para附加信息, parastr收费编号, para业务分类)
    Set lobj使用对象 = Nothing
End Function

Public Function func查询费用信息(Optional ByVal para条件 As String) As Object
    
    Set func查询费用信息 = dafuncGetData("select * from 收费管理_费用信息表 " & IIf(para条件 = "", "", " where " & para条件))
    
End Function

'功能：查询收费详细信息。
'创建：2002-6-24（杨春）。
Public Function func查询费用详细信息(Optional ByVal para条件 As String = "") As Object
    Set func查询费用详细信息 = dafuncGetData("select 收费批号,收费编号,收费项目编号,收费项目名称,单价,数量,计量单位,金额,交费单位名称,交费人,交费日期,主管科室名称,主管科室经手人,收费人姓名,打折比率,交费方式=case 交费方式 when 1 then '现金' when 2 then '支票' else '' end ,收费状态=case 收费状态 when 0 then '未收费' when 1 then '已收费' when 2 then '已退费' else 收费状态 end ,收据号 from 收费管理_费用详细信息视图 where " & IIf(para条件 = "", "1=1", para条件))
End Function

Public Function func查询收费项目(Optional ByVal para条件 As String) As Object
    Set func查询收费项目 = dafuncGetData("select * from 收费管理_收费项目字典表" & IIf(para条件 = "", "", " where " & para条件))
End Function

Public Function func查询收费标准(Optional ByVal para条件 As String) As Object
    Set func查询收费标准 = dafuncGetData("select * from 收费管理_收费标准信息视图" & IIf(para条件 = "", "", " where " & para条件) & " order by 收费标准名称")
End Function

'功能：根据名称、单价添加指定收费项目的下级收费项目，并自动生成收费项目编号、
'      计量单位='项'、票据类型(=上级的)、最小单价（=单价）、最大单价（=单价）。
'假设：每级收费项目的长度=3。
'输入：para上级项目编号 String  '上级收费项目的编号。
'      para收费项目名称 String  '要生成的收费项目的名称。
'      para单价         Single  '单价。
'返回：生成的收费项目编号。
'创建：2002-2-20（杨春）。
Public Function func添加收费项目(ByVal para上级项目编号 As String, ByVal para收费项目名称 As String, ByVal para单价 As Single) As String
    Dim lobjRec As Object
    Dim lstr票据类型 As String '编号。
    Dim lstr收费编号 As String
    Dim llngMax As Long
    
    On Error GoTo errHandler
    func添加收费项目 = ""
    
    '获取上级项目的票据类型。
    Set lobjRec = dafuncGetData("select 票据类型编号 from 收费管理_收费项目字典表 where 收费项目编号='" & para上级项目编号 & "'")
    If lobjRec.RecordCount = 0 Then
        Err.Raise 6666, , "该上级项目不存在，无法添加其下级收费项目。"
    End If
    lstr票据类型 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
        
    '判断是否有同名的收费项目。
    Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号 like '" & para上级项目编号 & "%' and len(收费项目编号)=" & Len(para上级项目编号) + 3 & " and 收费项目名称='" & para收费项目名称 & "'")
    If lobjRec.RecordCount > 0 Then
        '存在。
        func添加收费项目 = lobjRec("收费项目编号")
    Else
        '不存在，自动生成收费编号。
        
        '若最大编号已>999：
        'a00代表1000，a01代表1001，...a99代表1099；
        'b00代表1100，b01代表1101，...b99代表1199；
        'j00代表1900，i01代表1901，...i99代表1999；
        'j00代表2000，...
        
        '获取已有最大编号。
        Set lobjRec = dafuncGetData("select max(收费项目编号) from 收费管理_收费项目字典表 where 收费项目编号 like '" & para上级项目编号 & "%' and len(收费项目编号)=" & Len(para上级项目编号) + 3)
        If IIf(IsNull(lobjRec(0)), "", lobjRec(0)) = "" Then
            llngMax = 1
        Else
            If Left(Right(lobjRec(0), 3), 1) >= "0" And Left(Right(lobjRec(0), 3), 1) <= "9" Then
                llngMax = Val(Right(lobjRec(0), 3)) + 1
            Else '>999。
                llngMax = (Asc(Left(Right(lobjRec(0), 3), 1)) - Asc("a")) * 100 + 1000 + Val(Right(lobjRec(0), 2)) + 1
            End If
        End If
        If llngMax > 999 Then
            lstr收费编号 = para上级项目编号 & Chr(Asc("a") + (llngMax - 999 - 1) \ 100) & Format(Right(llngMax, 2), "00")
        Else
            lstr收费编号 = para上级项目编号 & Format(llngMax, "000")
        End If
        
        '生成助记符。
        Dim lstr助记符 As String
        lstr助记符 = guf_GetFirstLetter(para收费项目名称)
        lstr助记符 = Left(lstr助记符, 20)
        
        '新增记录（计量单位='项'，最小单价=最大单价=单价，票据类型=上级的）。
        dafuncGetData "insert into 收费管理_收费项目字典表(收费项目编号,收费项目名称,单价,计量单位,票据类型编号,最小单价,最大单价,助记符) values(" _
            & "'" & lstr收费编号 & "','" & para收费项目名称 & "'," & para单价 & ",'项','" & lstr票据类型 & "'," & para单价 & "," & para单价 & ",'" & lstr助记符 & "')"
        
        '返回新增项目编号。
        func添加收费项目 = lstr收费编号
    End If
    Exit Function
errHandler:
    Err.Raise Err.Number, , Err.Description
End Function

'创建：2002-2-20。
Public Sub sub修改收费项目(ByVal para收费项目编号 As String, ByVal para收费项目名称 As String)
    dafuncGetData "update 收费管理_收费项目字典表 set 收费项目名称='" & para收费项目名称 & "' where 收费项目编号='" & para收费项目编号 & "'"
End Sub

'创建：2003-9-23。
'输入：para收费确认信息：收费批号,打折比率,收费方式,交费日期,收费人,退费日期
Public Sub sub收费(ByVal para收费编号 As String, ByVal para交费日期 As String, para收费人 As String)
    Dim lobjRec As Object
    
    Dim lstr收据号  As String
    Dim lint收费方式 As Integer
    
    '获取现金收费方式的编号。
    Set lobjRec = dafuncGetData("select * from 收费管理_交费方式字典表 where 名称='现金'")
    If lobjRec.RecordCount > 0 Then
        lint收费方式 = Val(lobjRec!编号)
    Else
        lint收费方式 = 1
    End If
    
    Dim lcol收费确认信息 As Collection
    Dim lcol收费编号集 As Collection
    
    Set lcol收费编号集 = New Collection
    lcol收费编号集.Add para收费编号
    
    Set lcol收费确认信息 = New Collection
    With lcol收费确认信息
        .Add para收费编号, "收费批号"
        .Add 1, "打折比率"
        .Add lint收费方式, "收费方式"  '现金。
        .Add CDate(para交费日期), "交费日期"
        .Add para收费人, "收费人"
    End With
    
    Dim lobj使用对象 As Object
    Set lobj使用对象 = CreateObject("收费对象部件.cls收费管理")
    Call lobj使用对象.sub收费确认(lcol收费编号集, lcol收费确认信息)
    
End Sub

'创建：2003-9-23.
Public Sub sub修改交费人基本信息(ByVal para收费编号 As String, ByVal para交费人 As String, ByVal para交费单位名称 As String, ByVal para交费单位编号 As String)
    dafuncGetData "update 收费管理_费用信息表 set 交费人='" & para交费人 & "',交费单位名称='" & para交费单位名称 & "',交费单位编号='" & para交费单位编号 & "' where 收费编号='" & para收费编号 & "'"
End Sub
