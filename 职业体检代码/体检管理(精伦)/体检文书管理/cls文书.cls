VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls文书"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mblnPrint As Boolean
'作者：杨春。

'功能：打印体检管理的当前系统编号的指定文书。
'输入：para文书名称
'      paraRec as Recordset 文书内容。
'作者：杨春。
Public Sub Sub打印文书(ByVal para文书名称 As String, _
                    ByVal para文书内容 As Object, _
                    ByVal para是否预览 As Boolean, _
                    Optional para是否只预览 As Boolean = False)
    Dim lcol文书内容 As Collection
    Dim lcolItem As Collection
    Dim i As Long
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    If para文书内容 Is Nothing Then
        Err.Raise 6666, , "文书内容记录为空，不能打印文书。"
    End If
    If para文书内容.recordcount = 0 Then
        Err.Raise 6666, , "文书内容记录为空，不能打印文书。"
    End If
    
    Select Case para文书名称
    Case "体检登记表"
        '套打。
        With frm打印体检表
            Set .pobj文书内容 = para文书内容
            If para是否只预览 Or para是否预览 Then
                .Show 1
            End If
            If Not para是否只预览 Then
                .PrintForm
            End If
            Set .pobj文书内容 = Nothing
        End With
        
    Case "体检表" '修改：2003-4-16（杨春）增加体检表。
        With frm打印体检表
            Set .pobj文书内容 = para文书内容
            If para是否只预览 Or para是否预览 Then
                .Show 1
            End If
            If Not para是否只预览 Then
                .PrintForm
            End If
            Set .pobj文书内容 = Nothing
        End With
    Case "阳性体检单"
        Set lcol文书内容 = New Collection
        Set lcolItem = New Collection
        lcolItem.Add para文书内容.Fields(i).Name, "名称"
        lcolItem.Add IIf(IsNull(para文书内容!系统编号), "", para文书内容.Fields("系统编号").Value), "值"
        lcol文书内容.Add lcolItem, lcolItem("名称")

        '获取体检单耗。
        Set lobjRec = dafuncGetData("select 项目值 from 体检管理_体检附加信息表 where 系统编号='" & para文书内容!系统编号 & "' and 附加项目='体检表号'")
        Set lcolItem = New Collection
        lcolItem.Add "体检表号", "名称"
        If lobjRec.recordcount > 0 Then
            lcolItem.Add IIf(IsNull(lobjRec(0)), "", lobjRec(0).Value), "值"
        Else
            lcolItem.Add "", "值"
        End If
        lcol文书内容.Add lcolItem, lcolItem("名称")
        
        Set lcolItem = New Collection
        lcolItem.Add "体检单位", "名称"
        lcolItem.Add um防疫站名, "值"
        lcol文书内容.Add lcolItem, lcolItem("名称")
        
        '获取照片。
        Dim lpicPhoto As StdPicture
        Dim lobjSys As Object
        Set lobjSys = CreateObject("Scripting.FileSystemObject")
        '先拷贝空白签名的图片。
        lobjSys.CopyFile App.Path & "\空白照片.bmp", "c:\体检照片.bmp"

        Set lpicPhoto = pmfunc获取图片(para文书内容!健康档案编号, "体检管理")
        If Not lpicPhoto Is Nothing Then
            SavePicture lpicPhoto, "c:\体检照片.bmp"
        End If
        
        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
    
    Case "体检结果单"
        Set lcol文书内容 = New Collection
        For i = 0 To para文书内容.Fields.Count - 1
            Set lcolItem = New Collection
            lcolItem.Add para文书内容.Fields(i).Name, "名称"
            lcolItem.Add IIf(IsNull(para文书内容.Fields(i).Value), "", para文书内容.Fields(i).Value), "值"
            lcol文书内容.Add lcolItem
        Next
        
        Set lcolItem = New Collection
        lcolItem.Add "体检单位", "名称"
        lcolItem.Add um防疫站名, "值"
        lcol文书内容.Add lcolItem, lcolItem("名称")
        
        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
        
    Case "体检结果通知单"
        '把文书内容按水晶报表打印的要求存放在collection中。
        Set lcol文书内容 = New Collection
        For i = 0 To para文书内容.Fields.Count - 1
            Set lcolItem = New Collection
            lcolItem.Add para文书内容.Fields(i).Name, "名称"
            lcolItem.Add IIf(IsNull(para文书内容.Fields(i).Value), "", para文书内容.Fields(i).Value), "值"
            lcol文书内容.Add lcolItem
        Next
        
        Set lcolItem = New Collection
        lcolItem.Add "体检单位", "名称"
        lcolItem.Add um防疫站名, "值"
        lcol文书内容.Add lcolItem, "tjdw"
        
        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
        
    Case "B超体检结果单"
        Set lcol文书内容 = New Collection
        Set lcolItem = New Collection
        lcolItem.Add "系统编号", "名称"
        lcolItem.Add para文书内容("系统编号").Value, "值"
        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
        
'    Case "化验单"
'        Dim lobj化验单 As Object
'        '创建化验单对象。
'        Set lobj化验单 = CreateObject("化验单对象部件.clsBigCheckSheet")
'        lobj化验单.化验单编号 = para文书内容!化验单编号
'
'        '启动化验单打印界面。
'        Set frm打印化验单.pobj化验单对象 = lobj化验单
'
'        '先预览。
'        If para是否预览 Or para是否只预览 Then
'            frm打印化验单.Move 1330, 400
'            frm打印化验单.Show 1
'        End If
'        If Not para是否只预览 Then
'            frm打印化验单.c化验单.subPrint
'        End If
'        Set frm打印化验单.pobj化验单对象 = Nothing
        
    Case Else
        '体检单，要打印条形码和照片，所以采用窗体打印。
        With frm打印体检单
            '设置那些条码需要打印。
            Select Case para文书名称
            Case "尿常规体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = False
                .pbln是否有验血 = False
                
            Case "大便培养体检单"
                .pbln是否有尿常规 = False
                .pbln是否有大便培养 = True
                .pbln是否有验血 = False
            
            Case "验血体检单"
                .pbln是否有尿常规 = False
                .pbln是否有大便培养 = False
                .pbln是否有验血 = True
            
            Case "尿常规大便培养体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = True
                .pbln是否有验血 = False
                
            Case "尿常规验血体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = True
                .pbln是否有验血 = True
            
            Case "大便培养验血体检单"
                .pbln是否有尿常规 = False
                .pbln是否有大便培养 = True
                .pbln是否有验血 = True
            
            Case Else '"尿常规大便培养验血体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = True
                .pbln是否有验血 = True
            End Select
            
            '设置文书内容。
            
            Set .pobj文书内容 = para文书内容
            
            If para是否只预览 Or para是否预览 Then
                .Show 1
            End If
            
            If Not para是否只预览 Then
                '打印窗体。
                .PrintForm
            End If
            
            Set .pobj文书内容 = Nothing
            
        End With
    End Select
    Exit Sub
errhandler:
    sfsub错误处理 "体检文书管理", "cls文书", "Sub打印文书", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：获取体检单类型。
'作者：杨春。
Public Function func获取所有体检单类型() As Collection
    On Error GoTo errhandler
    Set func获取所有体检单类型 = New Collection
    With func获取所有体检单类型
        .Add "尿常规体检单", "尿常规体检单"
        .Add "大便培养体检单", "大便培养体检单"
        .Add "验血体检单", "验血体检单"
        
        .Add "尿常规大便培养体检单", "尿常规大便培养体检单"
        .Add "尿常规验血体检单", "尿常规验血体检单"
        .Add "大便培养验血体检单", "大便培养验血体检单"
        
        .Add "尿常规大便培养验血体检单", "尿常规大便培养验血体检单"
        
    End With

    Exit Function
errhandler:
    sfsub错误处理 "体检文书管理", "cls文书", "func获取所有体检单类型", Err.Number, Err.Description, True
End Function
'
'Public Function funcPrintReport(ByVal para文书名称 As String, ByVal para文书内容 As Collection, ByVal para是否预览 As Boolean) As Boolean
'    Dim i As Integer
'
'    On Error GoTo errHandler
'    '获取配置信息中的服务器名，数据库名。
'    Dim lstrServer  As String
'    Dim lstrDatabase  As String
''    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
''    lstrDatabase = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
'
'    If InStr(dapstrDSN, "Initial Catalog=") > 0 Then
'        lstrDatabase = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Initial Catalog=") - Len("Initial Catalog=") + 1)
'        If InStr(lstrDatabase, ";") > 0 Then
'            lstrDatabase = Left(lstrDatabase, InStr(lstrDatabase, ";") - 1)
'        End If
'    Else
'        lstrDatabase = "防疫2001"
'    End If
'
'    If InStr(dapstrDSN, "Data Source=") > 0 Then
'        lstrServer = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Data Source=") - Len("Data Source=") + 1)
'        If InStr(lstrServer, ";") > 0 Then
'            lstrServer = Left(lstrServer, InStr(lstrServer, ";") - 1)
'        End If
'    Else
'        lstrServer = "local"
'    End If
'
'
'    frmPrint.cRepPrint.Connect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=user26;PWD = welcome;Initial Catalog=" & lstrDatabase & ";Data Source=" & lstrServer
'
'
'    frmPrint.cRepPrint.ReportFileName = App.Path & "\" & para文书名称 & ".rpt"
'
'    '清空多余的参数
'    i = 0
'    While frmPrint.cRepPrint.Formulas(i) <> ""
'        If frmPrint.cRepPrint.Formulas(i) <> "" Then frmPrint.cRepPrint.Formulas(i) = ""
'        i = i + 1
'    Wend
'    '设置打印参数
'    For i = 1 To para文书内容.Count
'        frmPrint.cRepPrint.Formulas(i - 1) = Trim(para文书内容(i)("名称")) & "='" & Trim(para文书内容(i)("值")) & "'"
'    Next i
'
'    '显示报表
'    With frmPrint
'        If para是否预览 = True Then
'            .cRepPrint.WindowState = crptMaximized
'            .cRepPrint.WindowLeft = 0
'            .cRepPrint.WindowTop = 0
'            .cRepPrint.Destination = crptToWindow
'            .cRepPrint.Action = 1
'            .Visible = True
'
'        Else
'            .cRepPrint.Destination = crptToPrinter
'            .cRepPrint.Action = 1
'        End If
'    End With
'    funcPrintReport = True
'    Exit Function
'
'errHandler:
'    Dim llngErr As Long
'    Dim lstrError As String
'    llngErr = Err.Number
'    lstrError = Err.Description
'    Select Case llngErr
'    Case 20504
'        lstrError = "未找到打印文件" & para文书名称 & "。"
'    Case 20510
'        lstrError = "水晶报表公式与获取文书内容的存储过程不一致。"
'    Case 20526, 20545
'        lstrError = "存在妨碍打印的问题。此错误产生的原因及解决方法如下： " & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
'                    & "  (1) 没有从 Windows 控制面板中安装打印机？" & Chr(13) & Chr(10) _
'                    & "      解决：打开控制面板，双击“打印机”图标，选择“添加打印机”以装入打印机。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
'                    & "  (2) 打印机没在线？" & Chr(13) & Chr(10) _
'                    & "      解决：检查打印机与计算机的连接是否正常。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
'                    & "  (3)  打印机阻塞或缺纸？" & Chr(13) & Chr(10) _
'                    & "      解决：解决这些问题。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
'                    & "  (4) 试图在只能接受文本的打印机上打印窗体？" & Chr(13) & Chr(10) _
'                    & "      解决：切换到一台能打印图形的打印机。"
'    End Select
'    sfsub错误处理 "体检文书管理", "cls文书", "funcPrintReport", 6666, lstrError, True
'
'    funcPrintReport = False
'    Exit Function
'    Resume
'End Function

Public Function funcPrintReport(ByVal para文书名称 As String, ByVal para文书内容 As Collection, ByVal para是否预览 As Boolean) As Boolean
    Dim i As Integer
    Dim lstrItem As String
    
    On Error GoTo errhandler
    
    Dim capp As New CRAXDRT.Application
    Dim cr As CRAXDRT.Report
    Dim lstrDatabase As String
    Dim lstrServer As String
    
'    lstrDatabase = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    If InStr(dapstrDSN, "Initial Catalog=") > 0 Then
        lstrDatabase = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Initial Catalog=") - Len("Initial Catalog=") + 1)
        If InStr(lstrDatabase, ";") > 0 Then
            lstrDatabase = Left(lstrDatabase, InStr(lstrDatabase, ";") - 1)
        End If
    ElseIf InStr(dapstrDSN, "Database=") > 0 Then
        lstrDatabase = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Database=") - Len("Database=") + 1)
        If InStr(lstrDatabase, ";") > 0 Then
            lstrDatabase = Left(lstrDatabase, InStr(lstrDatabase, ";") - 1)
        End If
    Else
        lstrDatabase = "防疫2001"
    End If
    If InStr(dapstrDSN, "Data Source=") > 0 Then
        lstrServer = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Data Source=") - Len("Data Source=") + 1)
        If InStr(lstrServer, ";") > 0 Then
            lstrServer = Left(lstrServer, InStr(lstrServer, ";") - 1)
        End If
    ElseIf InStr(dapstrDSN, "Server=") > 0 Then
        lstrServer = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Server=") - Len("Server=") + 1)
        If InStr(lstrServer, ";") > 0 Then
            lstrServer = Left(lstrServer, InStr(lstrServer, ";") - 1)
        End If
    Else
        lstrServer = "local"
    End If
    
    
    '设置报表文件。
    Set cr = capp.OpenReport(App.Path & "\" & para文书名称 & ".rpt")
    
    '连接到数据库。
    On Error Resume Next
    For i = 1 To cr.Database.Tables.Count
        cr.Database.Tables.Item(i).ConnectionProperties("User ID").Value = "user26"
        cr.Database.Tables.Item(i).ConnectionProperties("Password").Value = "welcome"
        cr.Database.Tables.Item(i).ConnectionProperties("Database").Value = lstrDatabase
        If InStr(cr.Database.Tables.Item(i).Location, ".dbo.") > 0 Then
            cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & Right(cr.Database.Tables.Item(i).Location, Len(cr.Database.Tables.Item(i).Location) - InStr(cr.Database.Tables.Item(i).Location, ".dbo.") + 1)
        Else
            cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & cr.Database.Tables.Item(i).Location
        End If
        If Err.Number <> 0 Then Exit For
    Next
    
    
    '使用oledb连接
    If Err.Number <> 0 Then
        Err.Clear

        For i = 1 To cr.Database.Tables.Count
            cr.Database.Tables(i).SetLogOnInfo lstrServer, lstrDatabase, "user26", "welcome"
        Next
        cr.Database.LogOnServerEx "p2soledb.dll", lstrServer, "", "user26", "welcome", "OLE DB", "Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrDatabase & ";Data Source=" & lstrServer

        For i = 1 To cr.Database.Tables.Count
            If InStr(cr.Database.Tables.Item(i).Location, ".dbo.") > 0 Then
                cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & Right(cr.Database.Tables.Item(i).Location, Len(cr.Database.Tables.Item(i).Location) - InStr(cr.Database.Tables.Item(i).Location, ".dbo.") + 1)
            Else
                cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & cr.Database.Tables.Item(i).Location
            End If
        Next
    End If
    
'    If Err.Number <> 0 Then
'        Err.Clear
        
'        For i = 1 To cr.Database.Tables.Count
'            Dim dbTable As Object
'            Dim cpProperties As Object
'            Dim lstrLocation As String
'            Set dbTable = cr.Database.Tables(i)
'            lstrLocation = dbTable.Location
'            Set cpProperties = dbTable.ConnectionProperties
'            dbTable.DllName = "crdb_ado.dll"
'            cpProperties.DeleteAll
'            cpProperties.Add "Provider", "SQLOLEDB"
'            cpProperties.Add "Data Source", lstrServer
'            cpProperties.Add "Initial Catalog", lstrDatabase
'            cpProperties.Add "User ID", "user26"
'            cpProperties.Add "Password", "welcome"
'            dbTable.Location = lstrLocation
'        Next
'    End If
    
    '设置打印参数
    Dim Item As CRAXDRT.FormulaFieldDefinition
    For i = 1 To para文书内容.Count
        lstrItem = Trim(para文书内容(i)("值"))
        '把回车符替换掉。
        If InStr(1, lstrItem, Chr(13) & Chr(10)) > 0 Then
            lstrItem = Replace(lstrItem, Chr(13) & Chr(10), " ")
        End If
                
        On Error Resume Next

        Set Item = cr.FormulaFields.GetItemByName(para文书内容(i)("名称"))

        Item.Text = "'" & lstrItem & "'"
        On Error GoTo errhandler
    Next i
    
    cr.DiscardSavedData
    
    '显示报表
    Dim lobjFrm As New frmPrint
    With lobjFrm.cRepPrint
        If para是否预览 = True Then
            If InStr(para文书名称, "_") > 0 Then
                lobjFrm.Caption = Right(para文书名称, Len(para文书名称) - InStr(para文书名称, "_"))
            Else
                lobjFrm.Caption = para文书名称
            End If
            .ReportSource = cr
            .ViewReport
            lobjFrm.Show
        Else
            cr.PrintOut False
        End If
    End With
    
    funcPrintReport = True
    Exit Function
    
errhandler:
    Dim llngErr As Long
    Dim lstrError As String
    llngErr = Err.Number
    lstrError = Err.Description
    If llngErr = 20599 Then
        '连接数据库失败，重试。
        i = i + 1
        If i < 3 Then
            Resume
        Else
            lstrError = "连接数据库失败，可能是网络阻塞引起。请找系统管理员解决。"
        End If
    ElseIf llngErr = -2147206461 Then
        lstrError = "文书模板“" & App.Path & "\" & para文书名称 & ".rpt" & "”不存在！" & Chr(13) & Chr(10) & "请找到这个文件并拷贝到" & App.Path & "下面后再打印。"
    Else
        lstrError = func错误处理(llngErr, lstrError)
    End If
    lstrError = "打印水晶文书“" & para文书名称 & "”时失败。原因是：" & lstrError
    
    sfsub错误处理 "文书管理部件", "cls文书集", "funcPrintReport", 6666, lstrError, True
    funcPrintReport = False
    Exit Function
    Resume
End Function


Private Function func错误处理(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case -2147217833
        func错误处理 = "输入数据过长（或过大），已超过系统规定长度（或大小）。"
    Case 6
        func错误处理 = "输入数据过大，已超过系统规定大小。"
    Case -2147217913
        func错误处理 = "日期格式非法！"
    Case 94 '无效使用Null。
        func错误处理 = "使用的字典项被人通过字典管理操作删除了，系统无法再继续正常处理。请找系统管理员恢复字典内容。请注意，不要随便删除字典项！"
    Case 336, 337, 338, 429, 430
        func错误处理 = "系统部件已损坏（或已丢失），系统无法再正常运行。请退出系统，并重新安装系统。"
    Case 440 '外部对象错误：类自动错误。
        func错误处理 = "系统部件不正常终止运行。请退出系统，再重新启动系统。"
    Case 91 '对象没有初始化成功。
        func错误处理 = "因为网络阻塞，系统启动功能时无法完成正常的初始化。请退出功能界面，再重新进入功能界面。"
    Case 20504, 20507
        func错误处理 = "在系统路径下未找到要打印的文件模板文件。解决办法：" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                     & " (1) 请检查“文书管理部件.dll”系统文件所在路径下是否存在该水晶文书模板。" & Chr(13) & Chr(10) _
                     & " (2) 请进入“业务设置”的“文书格式设置”，在文书的“说明”信息中准确设置要打印的文书模板文件名（不带后缀）。"
    Case 20510
        func错误处理 = "水晶报表公式与获取文书内容的存储过程不一致（存储过程返回字段在水晶报表上不存在）。"
    Case 20526, 20545
        func错误处理 = "存在妨碍打印的问题。此错误产生的原因及解决方法如下： " & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (1) 没有从 Windows 控制面板中安装打印机？" & Chr(13) & Chr(10) _
                    & "      解决：打开控制面板，双击“打印机”图标，选择“添加打印机”以装入打印机。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (2) 打印机没在线？" & Chr(13) & Chr(10) _
                    & "      解决：检查打印机与计算机的连接是否正常。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (3)  打印机阻塞或缺纸？" & Chr(13) & Chr(10) _
                    & "      解决：解决这些问题。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (4) 试图在只能接受文本的打印机上打印窗体？" & Chr(13) & Chr(10) _
                    & "      解决：切换到一台能打印图形的打印机。"
        
    Case Else
        func错误处理 = paraErrDes
    End Select
End Function


