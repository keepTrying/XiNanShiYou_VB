VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls体检管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B38515603AD"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'功能：为健康体检管理主界面提供查询处理。
Public Function func健康体检查询(ByVal para体检号 As String, _
                        ByVal para姓名 As String, _
                        ByVal para单位名称 As String, _
                        ByVal para体检日期从 As String, _
                        ByVal para体检日期到 As String, _
                        ByVal para种类 As String, _
                        ByVal para状态条件 As String, _
                        ByVal para发证单位 As String, _
                        Optional para体检系统编号 As String = "") As Object
    
    Dim lstrSql As String
    Dim lstrFilter As String
    lstrFilter = " where 1=1"
    If para体检号 <> "" Then lstrFilter = lstrFilter & " and 体检号='" & para体检号 & "'"
    If para姓名 <> "" Then lstrFilter = lstrFilter & " and 姓名 like '" & para姓名 & "%'"
    If para单位名称 <> "" Then lstrFilter = lstrFilter & " and 单位名称 like '" & para单位名称 & "%'"
    If para体检日期从 <> "" Then lstrFilter = lstrFilter & " and 体检日期>= '" & para体检日期从 & "'"
    If para体检日期到 <> "" Then lstrFilter = lstrFilter & " and 体检日期<= '" & para体检日期到 & "'"
    If para种类 <> "" Then
        lstrFilter = lstrFilter & " and 种类='" & para种类 & "'"
    Else
        lstrFilter = lstrFilter & " and 种类!='食品卫生'"
    End If
    If para状态条件 <> "" Then lstrFilter = lstrFilter & IIf(para状态条件 = "", "", " and (" & para状态条件 & ")")
    If para发证单位 <> "" Then lstrFilter = lstrFilter & IIf(para发证单位 = "", "", " and 发证单位='" & para发证单位 & "'")
    If para体检系统编号 <> "" Then lstrFilter = lstrFilter & " and (体检系统编号='" & para体检系统编号 & "' or 体检号='" & para体检系统编号 & "' or 姓名 like '" & para体检系统编号 & "%')"
    
    lstrSql = "select  a.系统编号,体检号, 姓名,性别,年龄,单位名称,种类,项目值 血号,convert(varchar(10),体检日期,20) as 体检日期,体检结论,处置,健康证号,身份证号 食品证号,convert(varchar(10),发证日期,20) as 发证日期,发证单位,体检系统编号,状态 from 健康证管理_办证申请信息表 a,体检管理_体检附加信息表 b " & lstrFilter & " and a.体检系统编号=b.系统编号 and b.附加项目='血号' order by 血号"
    
    Set func健康体检查询 = dafuncGetData(lstrSql)
    
End Function

Public Sub sub打印健康证(ByVal paraCards As Collection)

    
    Dim i As Long
    Dim lintPrintModel As Integer
    Dim llngRealCount As Long
    Dim lModelNumber As Integer
    Dim lobjRec As Object
    
    On Error GoTo errordeal
    '预览。
    Set frmPrintCard.Cards = paraCards
    frmPrintCard.Show 1
    Set paraCards = frmPrintCard.Cards
    
    '判断打印板子类型。
    If frmPrintCard.opt模板类型(0).Value = True Then
        '2*5
        lintPrintModel = 0
        lModelNumber = 10
    Else
        If frmPrintCard.opt模板类型(1).Value = True Then
            '1*5
            lintPrintModel = 1
            lModelNumber = 5
        Else
            '单张。
            lintPrintModel = 2
            lModelNumber = 1
        End If
    End If
    
    
    If frmPrintCard.pblnPrint = True Then
           
        On Error GoTo errordeal
        
        '判断卡号是否已在库中存在
        For i = 1 To paraCards.Count
            Set lobjRec = dafuncGetData("select * from 健康证管理_办证申请信息表 where 系统编号<>'" & paraCards(i).系统编号 & "' and  健康证号='" & paraCards(i).健康证号 & "'")
            If lobjRec.RecordCount > 3 Then
                Err.Raise 6666, , "该健康证号(" & paraCards(i).健康证号 & ")已存在，请换一张卡！"
            End If
        Next
    
        '用户选择要打印。
        i = 0
        While i < paraCards.Count
            llngRealCount = lModelNumber
            If funcSetupPrint(paraCards, i, lintPrintModel) = True Then
                '修改：2002-5-23（设置打印的起始张）。
                If i = 0 And frmPrintCard.cframPos.Enabled Then
                    funcPrintEMC paraCards, lintPrintModel, i, frmPrintCard.ccmbIndex.ListIndex, frmPrintCard.ccmbSide.ListIndex
                    
                    '计算第一页实际打印的张数。
                    If lModelNumber > ((4 - frmPrintCard.ccmbIndex.ListIndex + 1) * 2 - frmPrintCard.ccmbSide.ListIndex) Then
                        '实际张数小于模板数。
                        llngRealCount = ((4 - frmPrintCard.ccmbIndex.ListIndex + 1) * 2 - frmPrintCard.ccmbSide.ListIndex)
                    End If
                    
                Else
                    funcPrintEMC paraCards, lintPrintModel, i
                End If
            
                
            End If
            i = i + llngRealCount
        Wend
        
        '打印成功，更新体检信息的状态。
        For i = 1 To paraCards.Count
            dafuncGetData "Update 健康证管理_办证申请信息表  Set 状态='已打印',健康证号 ='" & paraCards(i).健康证号 & "',身份证号 ='" & paraCards(i).身份证号 & "',发证日期='" & paraCards(i).发证日期 & "',有效期至='" & paraCards(i).有效期至 & "', 发证单位='" & paraCards(i).发证单位 & "' where 系统编号 ='" & paraCards(i).系统编号 & "'"
        Next
        '如果是新版卡，在公司服务器上更新当前已打印的最大卡号
        Dim lobjCheck As New clsCheck
        If Len(paraCards(i - 1).健康证号) = 10 Then lobjCheck.subSetCurNo paraCards(i - 1).健康证号
    End If
    Unload frmPrintCard
    
    Exit Sub

errordeal:
    MsgBox "更新数据库成功，但是打印失败，请选择“已打印”选项，重新打印！错误原因为：" & vbCr & Err.Description, vbInformation, "系统提示"
    Exit Sub
    On Error Resume Next
    
errDeal2:
    dasubRollBack
    MsgBox "更新数据库失败！请重新打印。" & Err.Description, vbInformation, "系统提示"
    Exit Sub
    On Error Resume Next

End Sub


'设置该版健康证
Private Function funcSetupPrint(paraCards As Collection, ByVal Index As Long, ByVal paraPrintModel As Integer) As Boolean
    On Error GoTo errordeal
    
    Dim i As Long
    Dim llngCount As Long
    
    llngCount = paraCards.Count - Index
    
    Dim con As Long
    Dim lModelNumber As Integer
    If paraPrintModel = 0 Then
        lModelNumber = 10
    Else
        If paraPrintModel = 1 Then
            lModelNumber = 5
        Else
            lModelNumber = 1
        End If
    End If
    con = lModelNumber - llngCount
    
    If llngCount < lModelNumber Then
        If MsgBox("该次打印健康证最后一版张数小于" & CStr(lModelNumber) & "张，将要浪费" & con & "张模版，是否还要打印该版健康证？", vbOKCancel, "健康证系统") = vbOK Then
            funcSetupPrint = True
        Else
            For i = 1 To llngCount
                paraCards.Remove (Index + 1)
            Next i
            funcSetupPrint = False
        End If
    Else
        funcSetupPrint = True
    End If
    

Exit Function

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next


End Function

'打印该版健康证
'修改：2002-2-23（增加参数：起始行号paraRow，起始列号ParaCol)
'      paraRow:0~4
'      paraCol:0~1,0代表左边列，1代表右边列。
Private Function funcPrintEMC(ByVal paraCards As Collection, ByVal paraPrintModel As Integer, Optional paraStart As Long = 0, Optional paraRow As Long = 0, Optional paraCol As Long = 0) As Boolean
    Dim lintX, lintY
    Dim Y As Long
    Dim k As Long
    Dim i As Long
    Dim j As Long
    Dim lobjPrintSeting As New ClsPrintSeting '创建打印对象,获取打印设置信息
    
    On Error Resume Next
    
    Printer.ScaleMode = 3
    
    Select Case paraPrintModel
    Case 0 '2*5  从左到右，从上到下。
        
        '设置对象打印版式为0，表示：2*5
        lobjPrintSeting.版式 = 0
        
        '设置纸张大小和字体。
        Printer.Font = lobjPrintSeting.字体
        Printer.PaperSize = lobjPrintSeting.纸张类型
        Printer.FontSize = lobjPrintSeting.字体大小
        
        k = 1
        For i = paraRow To 4
            lintY = lobjPrintSeting.顶端起始 + lobjPrintSeting.单张纵间距 * i
            
            For j = IIf(i = paraRow, paraCol, 0) To 1
                '修改：2002-5-23（没有证号不能打印）。
                If paraCards(k + paraStart).健康证号 = "" Then Exit For
                
                lintX = lobjPrintSeting.左边起始 + lobjPrintSeting.单张横间距 * j
                
                sub打印单张 paraCards, lintX, lintY, k + paraStart, lobjPrintSeting
                
                k = k + 1
            Next
        Next
        Printer.NewPage
        Printer.EndDoc
        
    Case 1 '1*5   从上到下。
        '设置打印版式
        lobjPrintSeting.版式 = 1
        Printer.Font = lobjPrintSeting.字体
        Printer.PaperSize = lobjPrintSeting.纸张类型
        Printer.FontSize = lobjPrintSeting.字体大小
        
        For i = 0 To IIf(paraCards.Count > 5, 4, paraCards.Count - 1)
            k = i + 1
            '修改：2002-5-23（没有证号不能打印）。
            If paraCards(k + paraStart).健康证号 = "" Then Exit For
            
            lintY = lobjPrintSeting.顶端起始 + lobjPrintSeting.单张纵间距 * i
            lintX = lobjPrintSeting.左边起始
            
            sub打印单张 paraCards, lintX, lintY, k + paraStart, lobjPrintSeting
        Next
        
        Printer.NewPage
        Printer.EndDoc
    
    Case 2 '1
    
        '设置打印板式为２表示是单张打印
        lobjPrintSeting.版式 = 2
        Printer.Font = lobjPrintSeting.字体
        Printer.PaperSize = lobjPrintSeting.纸张类型
        Printer.FontSize = lobjPrintSeting.字体大小
        
        lintY = lobjPrintSeting.顶端起始
        lintX = lobjPrintSeting.左边起始
        
        sub打印单张 paraCards, lintX, lintY, 1 + paraStart, lobjPrintSeting
        
        Printer.NewPage
        Printer.EndDoc
    End Select
End Function

'功能：在指定起始位置打印指定单张健康证。
'输入：paraTopX,paraTopy 该张健康证的起始位置。
'      paraIndex         该张健康证在frmPrintCard上的序号。
'      paraPrintSeting   打印设置对象。
Private Sub sub打印单张(ByVal paraCards As Collection, ByVal paraTopX As Long, ByVal paraTopY As Long, ByVal paraIndex As Long, ByVal paraPrintSeting As ClsPrintSeting)

    Printer.FontBold = False
    If paraCards(paraIndex).种类 = "食品卫生" Then
        '打印编号信息
        'If IsNull(paraPrintSeting.姓名) Then paraPrintSeting.姓名 = 380
        Printer.FontSize = 8
        Printer.FontName = "仿宋"
        '打印姓名和性别信息
        paraTopY = paraTopY + 60
        
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY
        Printer.Print "姓名：____________ 性别：____ 年龄：____"
        Printer.CurrentX = paraTopX + 200
        Printer.CurrentY = paraTopY - 10
        Printer.Print paraCards(paraIndex).姓名
        'Printer.Print "        " & paraCards(paraIndex).姓名
        Printer.CurrentX = paraTopX + 870    '+ paraPrintSeting.性别
        Printer.CurrentY = paraTopY - 10
        Printer.Print paraCards(paraIndex).性别
        Printer.CurrentX = paraTopX + 1230    '+ paraPrintSeting.性别
        Printer.CurrentY = paraTopY - 10
        Printer.Print paraCards(paraIndex).年龄
        
        '打印发证日期
        Dim lDate As Date
        
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + paraPrintSeting.行间距
        Printer.Print "有效期限：______年____月____日"
        Printer.CurrentX = paraTopX + 390
        Printer.CurrentY = paraTopY + paraPrintSeting.行间距 - 10
        Printer.Print Format(Date, "yyyy")
        Printer.CurrentX = paraTopX + 630
        Printer.CurrentY = paraTopY + paraPrintSeting.行间距 - 10
        Printer.Print Format(Date, "mm")
        Printer.CurrentX = paraTopX + 830
        Printer.CurrentY = paraTopY + paraPrintSeting.行间距 - 10
        Printer.Print Format(Date, "dd")
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + 2 * paraPrintSeting.行间距
        Printer.Print "　　　　至______年____月____日止"
        lDate = DateAdd("yyyy", 1, Date)
        Printer.CurrentX = paraTopX + 390
        Printer.CurrentY = paraTopY + 2 * paraPrintSeting.行间距 - 10
        Printer.Print Format(lDate, "yyyy")
        Printer.CurrentX = paraTopX + 630
        Printer.CurrentY = paraTopY + 2 * paraPrintSeting.行间距 - 10
        Printer.Print Format(lDate, "mm")
        Printer.CurrentX = paraTopX + 830
        Printer.CurrentY = paraTopY + 2 * paraPrintSeting.行间距 - 10
        Printer.Print Format(lDate, "dd")
        
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + 3 * paraPrintSeting.行间距
        
        Printer.Print "编　号：川(" & Format(Date, "yyyy") & ")5100001-" & paraCards(paraIndex).身份证号
            
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + 4 * paraPrintSeting.行间距
        Printer.Print "发证机构：_____________________________________"
        Printer.CurrentX = paraTopX + 300
        Printer.CurrentY = paraTopY + 4 * paraPrintSeting.行间距 - 10
        Printer.Print "四川省疾病预防控制中心预防医学门诊部"
        
        '照片框、照片。
        Printer.FontSize = paraPrintSeting.字体大小
        Printer.DrawWidth = 1
        Printer.Line (paraTopX + paraPrintSeting.照片x + 150 - 0.2, paraTopY + paraPrintSeting.照片y - 0.2 - 100)-(paraTopX + paraPrintSeting.照片x + 150 + paraPrintSeting.照片宽 - 30 + 0.2, paraTopY + paraPrintSeting.照片y + paraPrintSeting.照片高 - 30 + 0.2 - 100), &HE38B5B, B
        If paraPrintSeting.照片宽 > 0 And paraPrintSeting.照片高 > 0 Then
            If Not paraCards(paraIndex).照片 Is Nothing Then Printer.PaintPicture paraCards(paraIndex).照片, paraTopX + paraPrintSeting.照片x + 150, paraTopY + paraPrintSeting.照片y - 100, paraPrintSeting.照片宽 - 30, paraPrintSeting.照片高 - 30
        End If
        Printer.FontSize = 8
        Printer.CurrentX = paraTopX + paraPrintSeting.照片x + 430 - Len(paraCards(paraIndex).民族) * 20
        Printer.CurrentY = paraTopY + paraPrintSeting.照片y + paraPrintSeting.照片高 - 30 - 100 + 30
        Printer.Print paraCards(paraIndex).民族      '血号
    Else
        '打印编号信息
        If IsNull(paraPrintSeting.姓名) Then paraPrintSeting.姓名 = 380
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY
        Printer.Print "编　号：_______________________"
        Printer.CurrentX = paraTopX + paraPrintSeting.姓名
        Printer.CurrentY = paraTopY - 10
        Printer.Print paraCards(paraIndex).健康证号 & "        " & paraCards(paraIndex).民族      '血号
        
        '打印姓名和性别信息
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + paraPrintSeting.行间距
        Printer.Print "姓　名：__________ 性　别：____"
        Printer.CurrentX = paraTopX + paraPrintSeting.姓名
        Printer.CurrentY = paraTopY + paraPrintSeting.行间距 - 10
        Printer.Print paraCards(paraIndex).姓名
        Printer.CurrentX = paraTopX + paraPrintSeting.性别
        Printer.CurrentY = paraTopY + paraPrintSeting.行间距 - 10
        Printer.Print paraCards(paraIndex).性别
        
        '打印工种和年龄信息
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + 2 * paraPrintSeting.行间距
        Printer.Print "工　种：__________ 年　龄：____"
        Printer.CurrentX = paraTopX + paraPrintSeting.姓名
        Printer.CurrentY = paraTopY + 2 * paraPrintSeting.行间距 - 10
        Printer.Print paraCards(paraIndex).种类
        Printer.CurrentX = paraTopX + paraPrintSeting.性别
        Printer.CurrentY = paraTopY + 2 * paraPrintSeting.行间距 - 10
        Printer.Print paraCards(paraIndex).年龄
        
        '打印体检结果信息
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + 3 * paraPrintSeting.行间距
        Printer.Print "体　检：_______________________"
        Printer.CurrentX = paraTopX + paraPrintSeting.姓名
        Printer.CurrentY = paraTopY + 3 * paraPrintSeting.行间距 - 10
        Printer.Print IIf(paraCards(paraIndex).检出病种 = "无", "无从业禁忌症", paraCards(paraIndex).检出病种)
            
        '打印发证日期
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + 4 * paraPrintSeting.行间距
        Printer.Print "体检日期：_____________________"
        Printer.CurrentX = paraTopX + paraPrintSeting.姓名 + 140
        Printer.CurrentY = paraTopY + 4 * paraPrintSeting.行间距 - 10
        Printer.Print Format(paraCards(paraIndex).体检日期, "yyyy年mm月dd日")
        
        Printer.CurrentX = paraTopX
        Printer.CurrentY = paraTopY + 5 * paraPrintSeting.行间距
        Printer.Print "健康检查单位：_________________"
        Printer.CurrentX = paraTopX + paraPrintSeting.姓名 + 260
        Printer.CurrentY = paraTopY + 5 * paraPrintSeting.行间距 - 10
        Printer.Print "四川省疾控中心"
        
        
        '照片框、照片。
        Printer.FontSize = paraPrintSeting.字体大小
        Printer.DrawWidth = 1
        Printer.Line (paraTopX + paraPrintSeting.照片x - 0.2, paraTopY + paraPrintSeting.照片y - 0.2)-(paraTopX + paraPrintSeting.照片x + paraPrintSeting.照片宽 + 0.2, paraTopY + paraPrintSeting.照片y + paraPrintSeting.照片高 + 0.2), &HE38B5B, B
        If paraPrintSeting.照片宽 > 0 And paraPrintSeting.照片高 > 0 Then
            If Not paraCards(paraIndex).照片 Is Nothing Then Printer.PaintPicture paraCards(paraIndex).照片, paraTopX + paraPrintSeting.照片x, paraTopY + paraPrintSeting.照片y, paraPrintSeting.照片宽, paraPrintSeting.照片高
        End If
    
    End If

End Sub

Public Property Get 业务设置(ByVal para设置项名 As String) As String
    Dim lobjRec As Object
    On Error GoTo errhandler
    '健康证带条码
    Set lobjRec = dafuncGetData("select * from 健康证管理_业务设置信息表 where 设置项目='" & para设置项名 & "'")
    If lobjRec.RecordCount > 0 Then
        业务设置 = IIf(IsNull(lobjRec!设置值), "", lobjRec!设置值)
    End If
    Exit Property
errhandler:
    业务设置 = ""
End Property

Public Property Let 业务设置(ByVal para设置项名 As String, ByVal paraValue As String)
    Dim lobjRec As Object
    On Error GoTo errhandler
    
    Set lobjRec = dafuncGetData("select * from 健康证管理_业务设置信息表 where 设置项目='" & para设置项名 & "'")
    If lobjRec.RecordCount = 0 Then
        dafuncGetData "insert into 健康证管理_业务设置信息表(设置项目,设置值) values('" & para设置项名 & "','" & paraValue & "')"
    Else
        dafuncGetData "update 健康证管理_业务设置信息表 set 设置值='" & paraValue & "' where 设置项目='" & para设置项名 & "'"
    End If

    Exit Property
errhandler:
    sfsub错误处理 "健康证管理部件", "cls体检管理", "Property Let 业务设置", Err.Number, Err.Description, True
End Property



'规则：yy+流水号(5)
Public Function func生成系统编号(ByVal para卫生种类 As String) As String
    Dim lobjRec As Object
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData(" exec 健康证管理_生成系统编号 '" & para卫生种类 & "'")
    
    func生成系统编号 = lobjRec(0)
    
    Exit Function
errhandler:
    sfsub错误处理 "健康证管理部件", "cls体检管理", "func生成系统编号", Err.Number, Err.Description, True
    
End Function

'规则：yy+流水号(5)
Public Function func生成健康证号() As String
    Dim lobjRec As Object
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData(" exec 系统管理_生成系统编号 '健康证管理','健康证号'")
    
    func生成健康证号 = lobjRec(0)

    Exit Function
errhandler:
    sfsub错误处理 "健康证管理部件", "cls体检管理", "func生成健康证号", Err.Number, Err.Description, True
End Function

'规则：yy+流水号(4)
'Public Function func生成调离编号(ByVal para系统编号 As String) As String
'    Dim lobjRec As Object
'    On Error GoTo errhandler
'    Set lobjRec = dafuncGetData(" exec 健康证管理_生成系调离编号 '" & para系统编号 & "'")
'
'    func生成调离编号 = lobjRec(0)
'
'    Exit Function
'errhandler:
'    sfsub错误处理 "健康证管理部件", "cls体检管理", "func生成调离编号", Err.Number, Err.Description, True
'End Function

'Public Function func获取调离单位(ByVal para状态条件 As String) As Object
'    Set func获取调离单位 = dafuncGetData("select distinct 单位名称 from 健康证管理_办证申请信息表 where 处置='调离'" & IIf(para状态条件 = "", "", " and " & para状态条件))
'
'End Function
'Public Function func获取调离人员(ByVal para调离单位 As String) As Object
'    Set func获取调离人员 = dafuncGetData("select 系统编号,调离编号,姓名,性别,年龄,单位名称,种类,调离日期,调离期限,备注,状态 from 健康证管理_办证申请信息表 where 单位名称='" & para调离单位 & "' and 处置='调离' order by 调离编号 desc")
'
'End Function
'
'Public Sub sub打印调离通知(ByVal para调离编号 As String, _
'                            Optional para是否预览 As Boolean = True)
'    Dim i As Integer
'    Dim j As Long
'    Dim lstrItem As String
'    Dim lstr调离编号 As String
'    Dim lobjRec As Object
'
'    On Error GoTo errhandler
'    Dim lstrData As String
'    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
'
'    '设置报表文件。
'    frm打印水晶报表.cRepPrint.ReportFileName = App.Path & "\健康证_调离通知书.rpt"
'    frm打印水晶报表.cRepPrint.Connect = "odbc;dsn=WSFY2001;uid=user26;PWD=welcome;Database=" & lstrData
'
'
'    '清空多余的参数
'    j = 0
'    While frm打印水晶报表.cRepPrint.Formulas(j) <> ""
'        frm打印水晶报表.cRepPrint.Formulas(j) = ""
'        j = j + 1
'    Wend
'    '设置打印参数
'    frm打印水晶报表.cRepPrint.Formulas(0) = "调离编号='" & para调离编号 & "'"
'
'    '显示报表
'    With frm打印水晶报表
'        If para是否预览 = True Then
'
'            .cRepPrint.WindowState = crptMaximized
'            .cRepPrint.WindowControlBox = True
'            .cRepPrint.WindowLeft = 0
'            .cRepPrint.WindowParentHandle = 0
'            .cRepPrint.WindowTop = 0
'            .cRepPrint.Destination = crptToWindow
'            .cRepPrint.Action = 1
'
'            '.Visible = True
'        Else
'            .cRepPrint.WindowParentHandle = 0
'            .cRepPrint.Destination = crptToPrinter
'            .cRepPrint.Action = 1
'
'        End If
'    End With
'
'    Exit Sub
'
'errhandler:
'    Dim llngErr As Long
'    Dim lstrError As String
'    llngErr = Err.Number
'    lstrError = Err.Description
'    If llngErr = 20599 Then
'        lstrError = "连接数据库失败，请找系统管理员解决。"
'    Else
'        lstrError = func错误处理(llngErr, lstrError)
'    End If
'    lstrError = "打印水晶文书调离通知书时时失败。原因是：" & lstrError
'    Err.Raise 6666, lstrError
'    Exit Sub
'    Resume
'End Sub



