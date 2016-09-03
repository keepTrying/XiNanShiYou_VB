VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrintSeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlng版式 As Long                 '定义变量记录打印的版式

Public 纸张类型 As Long           '定义变量记录张类型
Public 顶端起始 As Long           '定义变量记录顶端起始
Public 左边起始 As Long           '定义变量记录左边起始
Public 单张横间距 As Long         '定义变量记录单张横间距
Public 单张纵间距 As Long         '定义变量记录单张纵间距
Public 行间距 As Long             '定义变量记录行间距
Public 字体 As String             '定义变量记录字体
Public 字体大小 As Long           '定义变量记录字体大小

Public 标题字体 As String
Public 标题字体大小 As Long           '定义变量记录字体大小

Public 健康证标题x As Long
Public 健康证标题y As Long
Public 健康证号 As Long           '定义变量记录健康证号
Public 姓名 As Long               '定义变量记录姓名
Public 性别 As Long               '定义变量记录性别
Public 工种 As Long               '定义变量记录工种
Public 年龄 As Long               '定义变量记录年龄
Public 体检结果 As Long           '定义变量记录体检结果
Public 体检日期 As Long           '定义变量记录体检日期

Public 照片框x1 As Long           '定义变量记录照片框x1
Public 照片框y1 As Long           '定义变量记录照片框y1
Public 照片框x2 As Long           '定义变量记录照片框x2
Public 照片框y2 As Long           '定义变量记录照片框y2
Public 照片x As Long              '定义变量记录照片x
Public 照片y As Long              '定义变量记录照片y
Public 照片高 As Long             '定义变量记录照片高
Public 照片宽 As Long             '定义变量记录照片宽

Public 培训证标题x As Long
Public 培训证标题y As Long
Public 培训结论x As Long
Public 培训结论y As Long
Public 培训日期x As Long
Public 培训日期y As Long

Public 发证单位x As Long           '定义变量记录发证单位
Public 发证单位y As Long
Public 发证单位标题x As Long
Public 发证单位标题y As Long




'功能：从数据中获取打印设置值的信息
Private Sub sub获取设置值()
On Error GoTo errhandler
    Dim lstrSql As String           '定义变量记录SQL语句
    Dim lobjRectemp As Object       '定义变量记录临时记录集
    
    lstrSql = "select * from 健康证_打印格式设置表 where 打印版式=" & mlng版式
    Set lobjRectemp = dafuncGetData(lstrSql)
    If lobjRectemp.RecordCount > 0 Then
        纸张类型 = IIf(IsNull(lobjRectemp("纸张类型")), 0, lobjRectemp("纸张类型"))
        左边起始 = IIf(IsNull(lobjRectemp("左边起始")), 0, lobjRectemp("左边起始"))
        顶端起始 = IIf(IsNull(lobjRectemp("顶端起始")), 0, lobjRectemp("顶端起始"))
        单张横间距 = IIf(IsNull(lobjRectemp("单张横间距")), 0, lobjRectemp("单张横间距"))
        单张纵间距 = IIf(IsNull(lobjRectemp("单张纵间距")), 0, lobjRectemp("单张纵间距"))
        行间距 = IIf(IsNull(lobjRectemp("行间距")), 0, lobjRectemp("行间距"))
        字体 = IIf(IsNull(lobjRectemp("字体")), "宋体", lobjRectemp("字体"))
        字体大小 = IIf(IsNull(lobjRectemp("字体大小")), 0, lobjRectemp("字体大小"))
        标题字体 = IIf(IsNull(lobjRectemp("标题字体")), "宋体", lobjRectemp("标题字体"))
        标题字体大小 = IIf(IsNull(lobjRectemp("标题字体大小")), 0, lobjRectemp("标题字体大小"))
        
        健康证标题x = IIf(IsNull(lobjRectemp("健康证标题x")), 0, lobjRectemp("健康证标题x"))
        健康证标题y = IIf(IsNull(lobjRectemp("健康证标题y")), 0, lobjRectemp("健康证标题y"))
        
        健康证号 = IIf(IsNull(lobjRectemp("健康证号")), 0, lobjRectemp("健康证号"))
        姓名 = IIf(IsNull(lobjRectemp("姓名")), 0, lobjRectemp("姓名"))
        性别 = IIf(IsNull(lobjRectemp("性别")), 0, lobjRectemp("性别"))
        工种 = IIf(IsNull(lobjRectemp("工种")), 0, lobjRectemp("工种"))
        年龄 = IIf(IsNull(lobjRectemp("年龄")), 0, lobjRectemp("年龄"))
        体检结果 = IIf(IsNull(lobjRectemp("体检结果")), 0, lobjRectemp("体检结果"))
        体检日期 = IIf(IsNull(lobjRectemp("体检日期")), 0, lobjRectemp("体检日期"))
        
        发证单位x = IIf(IsNull(lobjRectemp("发证单位x")), 0, lobjRectemp("发证单位x"))
        发证单位y = IIf(IsNull(lobjRectemp("发证单位y")), 0, lobjRectemp("发证单位y"))
        发证单位标题x = IIf(IsNull(lobjRectemp("发证单位标题x")), 0, lobjRectemp("发证单位标题x"))
        发证单位标题y = IIf(IsNull(lobjRectemp("发证单位标题y")), 0, lobjRectemp("发证单位标题y"))
        
        照片框x1 = IIf(IsNull(lobjRectemp("照片框x1")), 0, lobjRectemp("照片框x1"))
        照片框y1 = IIf(IsNull(lobjRectemp("照片框y1")), 0, lobjRectemp("照片框y1"))
        照片框x2 = IIf(IsNull(lobjRectemp("照片框x2")), 0, lobjRectemp("照片框x2"))
        照片框y2 = IIf(IsNull(lobjRectemp("照片框y2")), 0, lobjRectemp("照片框y2"))
        照片x = IIf(IsNull(lobjRectemp("照片x")), 0, lobjRectemp("照片x"))
        照片y = IIf(IsNull(lobjRectemp("照片y")), 0, lobjRectemp("照片y"))
        照片高 = IIf(IsNull(lobjRectemp("照片高")), 0, lobjRectemp("照片高"))
        照片宽 = IIf(IsNull(lobjRectemp("照片宽")), 0, lobjRectemp("照片宽"))
        
        培训证标题x = IIf(IsNull(lobjRectemp("培训证标题x")), 0, lobjRectemp("培训证标题x"))
        培训证标题y = IIf(IsNull(lobjRectemp("培训证标题y")), 0, lobjRectemp("培训证标题y"))
        培训结论x = IIf(IsNull(lobjRectemp("培训结论x")), 0, lobjRectemp("培训结论x"))
        培训结论y = IIf(IsNull(lobjRectemp("培训结论y")), 0, lobjRectemp("培训结论y"))
        培训日期x = IIf(IsNull(lobjRectemp("培训日期x")), 0, lobjRectemp("培训日期x"))
        培训日期y = IIf(IsNull(lobjRectemp("培训日期y")), 0, lobjRectemp("培训日期y"))
        
    End If
Exit Sub
errhandler:
        
End Sub

Property Let 版式(ByVal paraValue As Long)
    mlng版式 = paraValue
    sub获取设置值
End Property
Property Get 版式() As Long
    版式 = mlng版式
End Property


Property Get 属性值(ByVal para属性名 As String) As Variant
    Select Case para属性名
    Case "左边起始"
        属性值 = 左边起始
    Case "顶端起始"
        属性值 = 顶端起始
    Case "单张横间距"
        属性值 = 单张横间距
    Case "单张纵间距"
        属性值 = 单张纵间距
    Case "行间距"
        属性值 = 行间距
    Case "字体"
        属性值 = 字体
    Case "字体大小"
        属性值 = 字体大小
    Case "标题字体"
        属性值 = 标题字体
    Case "标题字体大小"
        属性值 = 标题字体大小
    Case "健康证标题x"
        属性值 = 健康证标题x
    Case "健康证标题y"
        属性值 = 健康证标题y
    Case "健康证号"
        属性值 = 健康证号
    Case "姓名"
        属性值 = 姓名
    Case "性别"
        属性值 = 性别
    Case "工种"
        属性值 = 工种
    Case "年龄"
        属性值 = 年龄
    Case "体检结果"
        属性值 = 体检结果
    Case "体检日期"
        属性值 = 体检日期
    Case "发证单位x"
        属性值 = 发证单位x
    Case "发证单位y"
        属性值 = 发证单位y
    Case "发证单位标题x"
        属性值 = 发证单位标题x
    Case "发证单位标题y"
        属性值 = 发证单位标题y
    Case "照片框x1"
        属性值 = 照片框x1
    Case "照片框y1"
        属性值 = 照片框y1
    Case "照片框x2"
        属性值 = 照片框x2
    Case "照片框y2"
        属性值 = 照片框y2
    Case "照片x"
        属性值 = 照片x
    Case "照片y"
        属性值 = 照片y
    Case "照片高"
        属性值 = 照片高
    Case "照片宽"
        属性值 = 照片宽
    Case "培训证标题x"
        属性值 = 培训证标题x
    Case "培训证标题y"
        属性值 = 培训证标题y
    Case "培训结论x"
        属性值 = 培训结论x
    Case "培训结论y"
        属性值 = 培训结论y
    Case "培训日期x"
        属性值 = 培训日期x
    Case "培训日期y"
        属性值 = 培训日期y
    Case Else
        属性值 = ""
    End Select
End Property

Property Let 属性值(ByVal para属性名 As String, ByVal paraValue As Variant)
    Select Case para属性名
    Case "左边起始"
        左边起始 = paraValue
    Case "顶端起始"
        顶端起始 = paraValue
    Case "单张横间距"
        单张横间距 = paraValue
    Case "单张纵间距"
        单张纵间距 = paraValue
    Case "行间距"
        行间距 = paraValue
    Case "字体"
        字体 = paraValue
    Case "字体大小"
        字体大小 = paraValue
    Case "标题字体"
        标题字体 = paraValue
    Case "标题字体大小"
        标题字体大小 = paraValue
    Case "健康证标题x"
        健康证标题x = paraValue
    Case "健康证标题y"
        健康证标题y = paraValue
    Case "健康证号"
        健康证号 = paraValue
    Case "姓名"
        姓名 = paraValue
    Case "性别"
        性别 = paraValue
    Case "工种"
        工种 = paraValue
    Case "年龄"
        年龄 = paraValue
    Case "体检结果"
        体检结果 = paraValue
    Case "体检日期"
        体检日期 = paraValue
    Case "发证单位x"
        发证单位x = paraValue
    Case "发证单位y"
        发证单位y = paraValue
    Case "发证单位标题x"
        发证单位标题x = paraValue
    Case "发证单位标题y"
        发证单位标题y = paraValue
    Case "照片框x1"
        照片框x1 = paraValue
    Case "照片框y1"
        照片框y1 = paraValue
    Case "照片框x2"
        照片框x2 = paraValue
    Case "照片框y2"
        照片框y2 = paraValue
    Case "照片x"
        照片x = paraValue
    Case "照片y"
        照片y = paraValue
    Case "照片高"
        照片高 = paraValue
    Case "照片宽"
        照片宽 = paraValue
    Case "培训证标题x"
        培训证标题x = paraValue
    Case "培训证标题y"
        培训证标题y = paraValue
    Case "培训结论x"
        培训结论x = paraValue
    Case "培训结论y"
        培训结论y = paraValue
    Case "培训日期x"
        培训日期x = paraValue
    Case "培训日期y"
        培训日期y = paraValue
    
    End Select
End Property

Public Function Clone(ByVal para版式名 As String) As ClsPrintSeting
    On Error GoTo errhandler
    
    Dim lobjSet As New ClsPrintSeting
    
    lobjSet.版式 = para版式名
    lobjSet.属性值("左边起始") = 左边起始
    lobjSet.属性值("顶端起始") = 顶端起始
    lobjSet.属性值("单张横间距") = 单张横间距
    lobjSet.属性值("单张纵间距") = 单张纵间距
    lobjSet.属性值("行间距") = 行间距
    lobjSet.属性值("字体") = 字体
    lobjSet.属性值("字体大小") = 字体大小
    lobjSet.属性值("标题字体") = 标题字体
    lobjSet.属性值("标题字体大小") = 标题字体大小
    lobjSet.属性值("健康证标题x") = 健康证标题x
    lobjSet.属性值("健康证标题y") = 健康证标题y
    lobjSet.属性值("健康证号") = 健康证号
    lobjSet.属性值("姓名") = 姓名
    lobjSet.属性值("性别") = 性别
    lobjSet.属性值("工种") = 工种
    lobjSet.属性值("年龄") = 年龄
    lobjSet.属性值("体检结果") = 体检结果
    lobjSet.属性值("体检日期") = 体检日期
    lobjSet.属性值("发证单位x") = 发证单位x
    lobjSet.属性值("发证单位y") = 发证单位y
    lobjSet.属性值("发证单位标题x") = 发证单位标题x
    lobjSet.属性值("发证单位标题y") = 发证单位标题y
    lobjSet.属性值("照片框x1") = 照片框x1
    lobjSet.属性值("照片框y1") = 照片框y1
    lobjSet.属性值("照片框x2") = 照片框x2
    lobjSet.属性值("照片框y2") = 照片框y2
    lobjSet.属性值("照片x") = 照片x
    lobjSet.属性值("照片y") = 照片y
    lobjSet.属性值("照片高") = 照片高
    lobjSet.属性值("照片宽") = 照片宽
    lobjSet.属性值("培训证标题x") = 培训证标题x
    lobjSet.属性值("培训证标题y") = 培训证标题y
    lobjSet.属性值("培训结论x") = 培训结论x
    lobjSet.属性值("培训结论y") = 培训结论y
    lobjSet.属性值("培训日期x") = 培训日期x
    lobjSet.属性值("培训日期y") = 培训日期y

    Set Clone = lobjSet
    
    Exit Function
errhandler:
    
    sfsub错误处理 "健康证打印格式设置A", "clsPrintSeting", "Clone", Err.Number, Err.Description, True
    
End Function


Public Sub sub保存()
    On Error GoTo errhandler
    Dim lstrSql As String
    
    dasubBeginTran
    lstrSql = "delete from 健康证_打印格式设置表 where 打印版式=" & mlng版式
    Call dafuncGetData(lstrSql)
    
    lstrSql = "insert into 健康证_打印格式设置表(打印版式,纸张类型,顶端起始,左边起始,单张横间距,单张纵间距,行间距," _
                & "字体大小,健康证号,体检结果,体检日期,发证单位x,发证单位y,性别," _
                & "照片框x1,照片框y1,照片框x2,照片框y2,照片x,照片y,照片高,照片宽)" _
                & "values(" & mlng版式 & "," & 纸张类型 & "," & 顶端起始 & "," & 左边起始 & "," & 单张横间距 & "," & 单张纵间距 & "," & 行间距 & "," _
                & 字体大小 & "," & 健康证号 & "," & 体检结果 & "," & 体检日期 & "," & 发证单位x & "," & 发证单位y & "," & 性别 & "," _
                & 照片框x1 & "," & 照片框y1 & "," & 照片框x2 & "," & 照片框y2 & "," & 照片x & "," & 照片y & "," & 照片高 & "," & 照片宽 & ")"
    Call dafuncGetData(lstrSql)
    dasubCommitTran
    
    Exit Sub
errhandler:
    dasubRollBack
    sfsub错误处理 "健康证打印格式设置", "clsPrintSeting", "sub保存", Err.Number, Err.Description, True
End Sub

Public Sub sub测试打印()
    Dim lcolCards As Collection
    Dim lobj体检 As cls体检
    Dim i As Long
    Dim llngCount As Long
    
    Set lcolCards = New Collection
    Select Case mlng版式
    Case 0 '2*5
        llngCount = 10
    Case 1 '1*5
        llngCount = 5
    Case 2 '单张
        llngCount = 1
    End Select
    
    For i = 1 To llngCount
        Set lobj体检 = New cls体检
        With lobj体检
            .姓名 = "张三" & i
            .性别 = "男"
            .年龄 = 27
            .种类 = "食品卫生"
            .检出病种 = "无从业禁忌症"
            .健康证号 = "000000" & i
            .体检日期 = Format(Date, "yyyy-mm-dd")
            .体检结论 = "合格"
            .发证单位 = "某某疾控中心"
            Set .照片 = LoadPicture(App.Path & "\Photo\测试.bmp")
        End With
        lcolCards.Add lobj体检
    Next
    
    Dim lobjPrint As New cls体检管理
    
    Printer.Font = 字体
    Printer.PaperSize = 纸张类型
    Printer.FontSize = 字体大小
    lobjPrint.sub打印健康证 lcolCards
    
End Sub

