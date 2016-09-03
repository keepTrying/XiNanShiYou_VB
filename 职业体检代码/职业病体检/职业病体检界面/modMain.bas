Attribute VB_Name = "modMain"

'名称：职业病史职业病对象启动
'函数：
'功能：刷二代身份证函数声名
'      错误处理
'作者：Yunle Liu
'时间：2012.03
Public mstr系统编号 As String  '2015-10-22
Public InputFlag As String
Public InputFlagNo As String
Option Explicit
'*********************************************************************
'二代身份证读卡器 函数声明
'Public Declare Function CVR_InitComm Lib "termb.dll" (ByVal Port As Long) As Integer
'Public Declare Function CVR_CloseComm Lib "termb.dll" () As Integer
'Public Declare Function CVR_Authenticate Lib "termb.dll" () As Integer
'Public Declare Function CVR_Read_Content Lib "termb.dll" (ByVal Active As Long) As Integer
'Public Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer
'
'Public Declare Function GetPeopleName Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
'Public Declare Function GetPeopleAddress Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
'Public Declare Function GetPeopleIDCode Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
'Public Declare Function GetPeopleNation Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Declare Function InitComm Lib "Sdtapi.dll" (ByVal iPort As Integer) As Integer

Declare Function Authenticate Lib "Sdtapi.dll" () As Integer

Declare Function ReadBaseInfos Lib "Sdtapi.dll" (ByVal iname As String, ByVal isex As String, ByVal folk As String, ByVal birthday As String, ByVal code As String, ByVal addr As String, ByVal agency As String, ByVal startdate As String, ByVal enddate As String) As Integer

Declare Function CloseComm Lib "Sdtapi.dll" () As Integer

Declare Function ReadBaseMsgW Lib "Sdtapi.dll" (ByVal pMsg As String, ByRef LenT As Integer) As Integer


Declare Function ReadBaseMsg Lib "Sdtapi.dll" (ByVal pMsg As String, ByRef LenT As Integer) As Integer
Declare Function ReadIINSNDN Lib "Sdtapi.dll" (ByVal pIINSNDN As String) As Integer
Declare Function GetSAMIDToStr Lib "Sdtapi.dll" (ByVal pcSAMID As String) As Integer
Global Comm As Boolean
Public Declare Function SendMessage Lib "user32" _
            Alias "SendMessageA" (ByVal hwnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long
'**********************************************************************

Public pobjDict As Object
Public pstrFilename As String
Public pstrWordname As String
Public pobjFileToDatabase As Object
Public pstr工作站代号 As String
Public pstrPhoto As String

Public 访问标志 As Integer
Public pobj业务对象 As Object '体检管理业务对象clsManageMedicalExam。

Public mstrQuery As String




Public Sub Main()
    On Error Resume Next
     '创建字典对象。
    Set pobjDict = CreateObject("字典管理.clsDictionary")
    Err.Clear
    '创建业务对象。
    Set pobj业务对象 = CreateObject("职业病对象.clsManageMedicalExam")
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "无法创建职业病史录入界面。请重新注册“职业病史录入界面.dll”。"
    End If
   
    
    
    Dim lstrServer As String
    Dim lstrData As String
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
'    lstrServer = "KAMA-AA251EA62C"
'    lstrData = "BJB-SJK2012"
    
     '创建写文件对象。
    Set pobjFileToDatabase = CreateObject("FileToDatabase.clsFileToDatabase")
    '与库建立连接。
    With pobjFileToDatabase
        .pstrConnectString = "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
        .subConnect
    End With
  Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "modmain", "Main", Err.Number, Err.Description, False
End Sub

'功能：填录入板内容。
'作者：杨春。
Public Sub sub填录入板值(ByVal para录入板 As Control, _
                        ByVal paraGUI As cls界面通用对象, _
                        ByVal paraInfo As Collection)
    Dim lstrItem As String
    Dim lstrItemText  As String
    Dim i As Integer
    Dim lint卫生种类 As Integer
    Dim lint行业类别 As Integer
    Dim j As Integer
    
    On Error GoTo errHandler
    
    
    para录入板.pblnTemp = True
    lint卫生种类 = 0
    
    For i = 1 To para录入板.InfoCollection.Count
        '录入项目名称。
        lstrItem = para录入板.InfoCollection(i).Title
        
        If sffunc判断集合键值是否存在(paraInfo, lstrItem) Then
            '设置TrueText。
            para录入板.ItemTrueText(i - 1) = paraInfo(lstrItem)("项目值编号")
            '设置Text。
            para录入板.ItemText(i - 1) = paraInfo(para录入板.InfoCollection(i).Title)("项目值")
            
            If lstrItem = "卫生种类" Then
                lint卫生种类 = i
            ElseIf lstrItem = "行业类别" Then
                lint行业类别 = i
            End If
        Else
            para录入板.ItemTrueText(i - 1) = ""
            para录入板.ItemText(i - 1) = ""
        End If
    Next i
    
    Dim lobjRec As Object
    Dim lstrItemTrueText As String
    '设置行业类别录入框的字典内容的条件。
    If lint卫生种类 > 0 And lint行业类别 > 0 Then

        '获取卫生种类编号。
        lstrItemTrueText = para录入板.ItemTrueText(lint卫生种类 - 1)

        '设置行业类别录入框的字典。
        If lstrItemTrueText <> "" And Not para录入板.InfoCollection(lint卫生种类).DictRecordSet Is Nothing Then
            Set lobjRec = para录入板.InfoCollection(lint卫生种类).DictRecordSet
            If Not lobjRec.EOF Then
                paraGUI.sub初始化字典表 lint行业类别, "Parent=" & lobjRec("InnerId")
            End If
        End If
    End If
  
    para录入板.pblnTemp = False
    Exit Sub
errHandler:
    para录入板.pblnTemp = False
    sfsub错误处理 "职业病界面部件", "modMain", "sub填录入板值", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

Public Sub sub显示单位属性(ByVal ciptBase As Control, _
            ByVal para单位申请编号 As String, _
            ByVal paraGUI As cls界面通用对象)
    Dim i As Long
    Dim lcolInfo As Collection
    
    If para单位申请编号 <> "" Then
    
        '获取单位属性。
        On Error Resume Next
        '获取单位属性。
        Set lcolInfo = pobj业务对象.func获取单位属性(para单位申请编号)
        
        ciptBase.pblnTemp = True
        
        ciptBase.Box1("卫生种类").TrueText = ""
        ciptBase.Box1("行业类别").TrueText = ""
        ciptBase.Box1("片区").TrueText = ""
        ciptBase.Box1("经济性质").TrueText = ""
        
        ciptBase.Box1("卫生种类").TrueText = lcolInfo("卫生种类")
        ciptBase.Box1("行业类别").TrueText = lcolInfo("行业类别")
        ciptBase.Box1("片区").TrueText = lcolInfo("片区")
        ciptBase.Box1("经济性质").TrueText = lcolInfo("经济性质")
        
        ciptBase.Box1("卫生种类").Text = lcolInfo("卫生种类名称")
        ciptBase.Box1("行业类别").Text = lcolInfo("行业类别名称")
        ciptBase.Box1("片区").Text = lcolInfo("片区名称")
        ciptBase.Box1("经济性质").Text = lcolInfo("经济性质名称")
        ciptBase.Box1("单位地址").Text = lcolInfo("单位地址")
        
        
        
        Dim lstrItem As String
        Dim lint卫生种类 As Integer
        Dim lint行业类别  As Integer
        
        Err.Clear
        
        '判断是否有卫生种类。
        For i = 1 To ciptBase.InfoCollection.Count
            '录入项目名称。
            lstrItem = ciptBase.InfoCollection(i).Title
            
            If lstrItem = "卫生种类" Then
                lint卫生种类 = i
            ElseIf lstrItem = "行业类别" Then
                lint行业类别 = i
            End If
            If Err <> 0 Then Exit For
        Next i
        
        '设置行业类别录入框的字典内容的条件。
        Dim lstrItemTrueText As String
        Dim lobjRec As Object
        If lint卫生种类 > 0 And lint行业类别 > 0 Then
            '获取卫生种类编号。
            lstrItemTrueText = ciptBase.ItemTrueText(lint卫生种类 - 1)
            '设置行业类别录入框的字典。
            If lstrItemTrueText <> "" And Not ciptBase.InfoCollection(lint卫生种类).DictRecordSet Is Nothing Then
                Set lobjRec = ciptBase.InfoCollection(lint卫生种类).DictRecordSet
                If Not lobjRec.EOF Then
                    paraGUI.sub初始化字典表 lint行业类别, "Parent=" & lobjRec("InnerId")
                End If
            End If
        End If
        
        ciptBase.pblnTemp = False
    End If

End Sub


'------打印文书部分重新整理，要与其它独立出来
'------(系统原来要求是：没有个人信息，不给打印表。但这里只是打条码)
'------但打印出来是什么样子?不清楚。现在还不知道省疾控那边希望打印成啥样~
Public Function sub打印单个体检条码号(ByVal para体检条码号 As String)
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病文书.cls文书")
    lobjTmp.sub打印条码 "仅打印条码", para体检条码号, False
End Function

'2012-04-05 陶露
'打印多个体检条码号
Public Function sub打印多个体检条码号(ByRef para体检条码号 As Collection)
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病文书.cls文书")
    lobjTmp.Sub打印多个文书 "仅打印条码", Nothing, False, False, para体检条码号
End Function
'2012-04-05 陶露

'2012-08-20 于登淼
'编辑word文档主函数
Public Function sub编辑word文档(paraParent As Object, ByVal paraSysNo As String, ByVal mstr体检表名称, ByVal paraReadOnly As Boolean)
    Dim objWord As Object                      'Word.Application
    Dim objWordDocument As Object       'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer                    '已有Word报告的ID
    Dim lstr体检日期 As String
    
    On Error GoTo errHandler
    
    '启动word
    On Error Resume Next
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            Err.Raise 6666, , "你没有安装Word，无法编辑报告！请先安装MS Office 2000以上版本。"
             Unload frmPrintPaper
        End If
    End If
    
    On Error GoTo errHandler
    
    Dim lstrNewDoc As String
    Dim lstrDotFile As String           '模板文件
    Dim i As Integer, j As Integer
    
    
    '添加文书。
    '取出模板文件
    '修改人：罗李奎 时间：2013-1-8 ↓
    '说明：根据体检表名取出需要的word模版
'    If mstr窗体名称 = "体检报告" Then
'         '判断文书是否已存在。
'    Set lobjRec = dafuncGetData("select 编号,文件类型 from 职业病体检_体检报告信息表 where 报告编号='" & paraSysNo & "'")
'    If lobjRec.RecordCount = 0 Then
''        If paraReadOnly Then        '不允许修改，表明为查看操作，不是新增操作
'            Err.Raise 6666, , "该样品没有录入Word报告，您此时不能再为其添加Word报告！", "系统提示"
''        End If
'    End If
'         pstrWordname = para体检表名称
'         mstr窗体名称 = ""
'    Else
'        pstrWordname = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.SelectedRow(0), 7)
'    End If
'     pstrWordname = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.SelectedRow(0), 7)
        pstrWordname = mstr体检表名称

    If pstrWordname = "" Then
        Err.Raise 6666, , "体检类型为空"
         Unload frmPrintPaper
    End If
    sub取出word模版
    '修改人：罗李奎 时间：2013-1-8 ↑
    
'     frm选择Word模板.pstrWordname = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.SelectedRow(0), 6)
'    frm选择Word模板.Show 1
'    lstrDotFile = frm选择Word模板.pstrFilename

    lstrDotFile = pstrFilename
    If lstrDotFile = "" Then Exit Function
    lstrNewDoc = App.Path & "\temp\" & paraSysNo & "_" & Format(Now, "yyyy-mm-dd") & ".doc"
    '打开模板，生成新文档（暂时为临时文件）
    Set objWordDocument = objWord.Documents.Open(FileName:=App.Path & "\" & lstrDotFile, ReadOnly:=False)
    objWordDocument.ActiveWindow.Caption = lstrNewDoc
    
    
    '获取书签内容。
    If Right(lstrDotFile, 4) = ".dot" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 4)
    ElseIf Right(lstrDotFile, 4) = "dotx" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 5)
    End If
    'Set lobjrec1 = dafuncGetData("exec 职业病体检_获取word报告_空白报告 '" & paraSysNo & "','" & um用户编号 & "','" & lstrDotFile & "'")
    dasubSetQueryTimeout 600
    Set lobjRec1 = dafuncGetData("select * from  职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'")

    'objWord.Visible = True
    
    '填充书签内容
    Dim lstrValue As String
    Dim myRange As Object, myTable As Object

    If lobjRec1.RecordCount > 0 Then
        Set myRange = objWordDocument.Content

        '处理表体的体检结果表格2012-10-25 罗李奎
        If objWordDocument.Tables.Count > 1 Then
            If InStr(lstrDotFile, "放射工作人员") > 0 Then                           '放射工作等
                sub加载放射工作人员健康word信息 objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "涉核性工作人员") > 0 Then                       '涉核工作等
                sub加载涉核工作人员职业健康word信息 objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "职业健康") > 0 Then         '职业健康
                sub加载职业健康word信息 objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "普通性工作人员") > 0 Then          '和普通体检等
                sub加载普通性工作人员健康word信息 objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "8023") Or InStr(lstrDotFile, "放射性工作人员") > 0 Then                       '8023和放射性工作等
                sub加载8023和放射性工作人员word信息 objWordDocument, myRange, paraSysNo
            End If
            
        End If

        '更新文档正文（不包括页眉、页脚）中的域对象，主要是页码、页数
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
        '保存文件
        objWordDocument.SaveAs lstrNewDoc
        objWordDocument.Saved = False
    End If

    With objWord.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = 0          'wdRevisionsViewFinal
    End With

    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0

    objWord.Visible = True

    On Error GoTo errHandler

    objWordDocument.Activate
    objWord.Activate
     
    '罗李奎  2013年1月10日   ↓
        If Not paraReadOnly Then
            If lintRepID = 0 Then
                 objWord.Run "subStart", paraParent, -1, paraSysNo
            Else
                objWord.Run "subStart", paraParent, lintRepID, paraSysNo
            End If
          End If
     '罗李奎  2013年1月10日   ↑
     
    If paraReadOnly Then
        If objWordDocument.Range.Fields.Count = 0 Then objWordDocument.Protect 3, , "sccdc789"
        objWordDocument.Saved = True
    End If
    Exit Function
    
errHandler:
    If Err = 3001 Then
        MsgBox "没有在数据库中找到该Word报告的具体文件，可能是保存该报告到系统中时网络发生故障，请在系统中删除该报告信息后，重新录入该报告。", vbInformation, "系统提示"
    Else
        sfsub错误处理 "界面部件", "mod检验界面", "sub编辑word文档", Err.Number, Err.Description, True
    End If
    Exit Function
    Resume
End Function

'2012-09-14 翁乔
'打印单位报表
Sub sub编辑单位报表(paralobjRec As Object, paralcol As Collection)
    Dim objWord As Object                      'Word.Application
    Dim objWordDocument As Object       'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer                    '已有Word报告的ID
    '填充书签内容
    Dim lstrValue As String
    Dim myRange As Object, myTable As Object
    Dim lstrNewDoc As String
    Dim lstrDotFile As String           '模板文件
    Dim i As Integer, j As Integer
    
    On Error GoTo errHandler
    
    '启动word
    On Error Resume Next
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            Err.Raise 6666, , "你没有安装Word，无法编辑报告！请先安装MS Office 2000以上版本。"
        End If
    End If

    '添加文书。
    '取出模板文件
    lstrDotFile = "职业病体检_单位报表.dot"
    If lstrDotFile = "" Then Exit Sub
    lstrNewDoc = App.Path & "\temp\" & paralcol("档案编号") & "_" & Format(Now, "yyyy-mm-dd") & ".doc"
    '打开模板，生成新文档（暂时为临时文件）
    Set objWordDocument = objWord.Documents.Open(FileName:=App.Path & "\" & lstrDotFile, ReadOnly:=False)
    objWordDocument.ActiveWindow.Caption = lstrNewDoc
    
    '获取书签内容。
    If Right(lstrDotFile, 4) = ".dot" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 4)
    ElseIf Right(lstrDotFile, 4) = "dotx" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 5)
    End If
    
    If paralobjRec.RecordCount > 0 Then
        Set myRange = objWordDocument.Content
        
            '处理表头，第一节内容。
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【单位名称】", ReplaceWith:=paralcol("单位名称"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【单位地址】", ReplaceWith:=paralcol("单位地址"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【文件编号】", ReplaceWith:=paralcol("档案编号") & Format(Now, "yyyymmdd"), Replace:=2

            '处理正文，第二节内容
            Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
            Set myTable = myRange.Tables(2)
            If myTable.rows.Count < paralobjRec.RecordCount + 1 Then
                j = myTable.rows.Count
                myTable.rows(j).Select
                objWordDocument.ActiveWindow.Selection.InsertRows paralobjRec.RecordCount - j + 1
                For i = 1 To paralobjRec.RecordCount - myTable.rows.Count + 1
                    myTable.rows.Add (myTable.rows(j))
                Next
            End If
            For i = 1 To paralobjRec.RecordCount
                For j = 1 To paralobjRec.Fields.Count
                    If j = paralobjRec.Fields.Count Then
                        myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(paralobjRec(j - 1)), "", Format(paralobjRec(j - 1), "yyyy-mm-dd"))
                    Else
                        myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(paralobjRec(j - 1)), "", paralobjRec(j - 1))
                    End If
                    
                Next
                paralobjRec.MoveNext
            Next
            
        '更新文档正文（不包括页眉、页脚）中的域对象，主要是页码、页数
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
        '保存文件
        objWordDocument.SaveAs lstrNewDoc
        objWordDocument.Saved = False
    End If

    With objWord.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = 0          'wdRevisionsViewFinal
    End With

    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0

    objWord.Visible = True

    On Error GoTo errHandler

    objWordDocument.Activate
    objWord.Activate

    Exit Sub
    
errHandler:
    If Err = 3001 Then
        MsgBox "没有在数据库中找到该Word报告的具体文件，可能是保存该报告到系统中时网络发生故障，请在系统中删除该报告信息后，重新录入该报告。", vbInformation, "系统提示"
    Else
        sfsub错误处理 "界面部件", "mod检验界面", "sub编辑单位报表", Err.Number, Err.Description, True
    End If
    Exit Sub
    Resume
End Sub

'2012-09-22 翁乔
'2012-09-22 翁乔
'功能：输出单位总检报告
Sub sub编辑总检报告(lcolFactor As Collection)
    Dim objWord As Object                      'Word.Application
    Dim objWordDocument As Object       'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer                    '已有Word报告的ID
    '填充书签内容
    Dim lstrValue As String
    Dim myRange As Object, myTable As Object
    Dim lstrNewDoc As String
    Dim lstrDotFile As String           '模板文件
    Dim i As Integer, j As Integer
    Dim lcolInfo As Collection, lcolInfo2 As Collection
    On Error GoTo errHandler
    
    '启动word
    On Error Resume Next
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            Err.Raise 6666, , "你没有安装Word，无法编辑报告！请先安装MS Office 2000以上版本。"
        End If
    End If

    '添加文书。
    '取出模板文件
    lstrDotFile = "单位公司职业健康体检报告.dot"    '2015-10-28
    
'    lstrDotFile = "公司职业健康体检报告.dot"
    If lstrDotFile = "" Then Exit Sub
    lstrNewDoc = App.Path & "\temp\" & Format(Now, "yyyymmddss") & ".doc"
    '打开模板，生成新文档（暂时为临时文件）
    Set objWordDocument = objWord.Documents.Open(FileName:=App.Path & "\" & lstrDotFile, ReadOnly:=False)
    objWordDocument.ActiveWindow.Caption = lstrNewDoc
    
    '获取书签内容。
    If Right(lstrDotFile, 4) = ".dot" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 4)
    ElseIf Right(lstrDotFile, 4) = "dotx" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 5)
    End If
    
'    If paralobjRec.RecordCount > 0 Then
        Set myRange = objWordDocument.Content
'        Set lcolInfo = lcolFactor("危害因素")
'        Set lcolInfo2 = lcolFactor("危害结果")
        
            '增加体检类别（在岗期间等）  2015-11-2↓
            Dim TlobjRec As Object
            Dim Tlstr As String
            Dim Testtype As String
            Set TlobjRec = dafuncGetData("select 体检类别 from 职业病体检_体检基本数据库 where  单位名称 = '" & lcolFactor("单位名称") & "' and (体检日期 >= '" & lcolFactor("开始日期") & "' and 体检日期 <= '" & lcolFactor("截止日期") & "') group by 体检类别")
'            Testtype = ""
            For i = 1 To TlobjRec.RecordCount
            Testtype = TlobjRec("体检类别")
             lcolFactor.Add Testtype, "体检类别" & i
            Testtype = Testtype & "、"
'            Testtype = Testtype & lcolFactor("体检类别" & i) & "、"
            TlobjRec.MoveNext

            Next
            Testtype = Left(Testtype, Len(Testtype) - 1)
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【体检表类别】", ReplaceWith:=Testtype, Replace:=2
            '2015-11-2↑
        
        
        
            '处理表头，第一节内容。
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【单位名称】", ReplaceWith:=lcolFactor("单位名称"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【体检日期】", ReplaceWith:=lcolFactor("体检日期"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【生成时间】", ReplaceWith:=Format(Now, "yyyy年mm月dd日"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【应检人数】", ReplaceWith:=lcolFactor("应检人数"), Replace:=2       '增加应检人数  2015-10-29
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【实检人数】", ReplaceWith:=lcolFactor("实检人数"), Replace:=2
            
            Dim lcolItem As Collection
            Dim lstr As String, ltemp As String, lstr2 As String, lsql As String
            Set lcolItem = lcolFactor("体检项目")
            For i = 1 To lcolItem.Count
                lstr = lstr & "，" & lcolItem(i)
            Next

            objWordDocument.Sections(1).Range.Find.Execute FindText:="【体检大项】", ReplaceWith:=lstr, Replace:=2
            lstr2 = "其中"
            dasubSetQueryTimeout 600
  
                '确定危害因素有几种  2015-10-29
            Dim KlobjRec As Object
            Dim lstr3 As String
            Dim lstr4 As String
            Set KlobjRec = dafuncGetData("select 危害因素 from 职业病体检_体检基本数据库 where  单位名称 = '" & lcolFactor("单位名称") & "' and (体检日期 >= '" & lcolFactor("开始日期") & "' and 体检日期 <= '" & lcolFactor("截止日期") & "') group by 危害因素")

            For i = 1 To KlobjRec.RecordCount
'            For i = 1 To 3
'                    FrmQueryCompany.BigNum (i)    '将小写阿拉伯数字改成大写  2015-11-4)
                '处理每个解除因素的不合格人员
                    Set lobjRec = dafuncGetData("select b.名称,count(*) 人数 from dbo.职业病体检_体检结果视图 a,职业病体检_体检项目设置表 b where a.体检项目=b.编码 and 系统编号 in(" _
                        & " select 系统编号 from 职业病体检_体检基本数据库 where 危害因素 = '" & lcolFactor("危害因素" & i) & "'and 单位名称 = '" & lcolFactor("单位名称") & "'" _
                        & " and (体检日期 >= '" & lcolFactor("开始日期") & "' and 体检日期 <= '" & lcolFactor("截止日期") & "') and 体检状态 in( 7,5)" _
                        & ") and 单项结论 = '不合格' group by b.名称")
'                    If lobjRec Is Nothing Then Exit Sub
                    If lcolFactor("危害因素" & i) <> "" Then
                        If Not lobjRec Is Nothing Then
                        While Not lobjRec.EOF
                            ltemp = ltemp & lobjRec("名称") & "不合格" & lobjRec("人数") & "人，"
                            lobjRec.MoveNext
                        Wend
                        End If

                        ltemp = Left(ltemp, Len(ltemp) - 1)
                        lstr2 = lstr2 & "接触" & lcolFactor("危害因素" & i) & "作业人员" & lcolFactor("人数" & i) & "人，"
                        If lstr3 = "" Then
                        lstr3 = lstr3 & "（" & i & "）接触" & lcolFactor("危害因素" & i) & "人员：" & Chr(13) & Chr(10) & IIf(ltemp = "", "本次体检未查出与职业相关的异常改变。", ltemp)
                        Else
                        lstr3 = lstr3 & Chr(13) & Chr(10) & "（" & i & "）接触" & lcolFactor("危害因素" & i) & "人员：" & Chr(13) & Chr(10) & IIf(ltemp = "", "本次体检未查出与职业相关的异常改变。", ltemp)
                        End If
                        If lstr4 = "" Then
                        lstr4 = lstr4 & "（" & i & "）接触" & lcolFactor("危害因素" & i) & "人员：" & Chr(13) & Chr(10) & IIf(ltemp = "", "本次体检未查出职业禁忌症和与职业相关的健康损害。", ltemp)
                        Else
                        lstr4 = lstr4 & Chr(13) & Chr(10) & "（" & i & "）接触" & lcolFactor("危害因素" & i) & "人员：" & Chr(13) & Chr(10) & IIf(ltemp = "", "本次体检未查出职业禁忌症和与职业相关的健康损害。", ltemp)
                        End If
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="【危害因素及结果】", ReplaceWith:="接触" & lcolFactor("危害因素" & i) & "人员：" & IIf(ltemp = "", "本次体检未查出与职业相关的异常改变。", ltemp), Replace:=2
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="【危害结果及结果】", ReplaceWith:=IIf(ltemp = "", "本次体检未查出与职业相关的异常改变。", ltemp), Replace:=2
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="【危害因素" & i & "】", ReplaceWith:="接触" & lcolFactor("危害因素" & i) & "人员：", Replace:=2
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="【危害结果" & i & "】", ReplaceWith:=IIf(ltemp = "", "本次体检未查出与职业相关的异常改变。", ltemp), Replace:=2

                    End If

                KlobjRec.MoveNext
                ltemp = ""
            Next
            lstr3 = lstr3 & Chr(13) & Chr(10) & "（" & i & "）详细体检结果见《职业健康检查结果一览表》和个人体检报告"
            lstr4 = lstr4 & Chr(13) & Chr(10) & "（" & i & "）本次体检其它检查项目结果异常者，受检者可根据有无临床症状或复查后仍为异常结果者到医院相关科室诊治。详细处理建议可见个人报告"
            lstr2 = Left(lstr2, Len(lstr2) - 1)
'            lstr3 = Left(lstr3, Len(lstr3) - 1)
            
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【危害因素】", ReplaceWith:=lstr2, Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【危害因素及结果】", ReplaceWith:=lstr3, Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="【危害因素及结论】", ReplaceWith:=lstr4, Replace:=2
            
'            For i = 1 To lcolInfo.Count
'                lstr = "select 系统编号,姓名,危害因素,(体检结论 + 诊断及处理意见) from 职业病体检_体检基本数据库 where " _
'                        & "单位名称 = '" & lcolFactor("单位名称") & "' and 危害因素 ='" & lcolInfo(i) & "' and (体检日期 >= '" & lcolFactor("开始日期") & "' and 体检日期 <= '" & lcolFactor("截止日期") & "')"
'                Set lobjRec = dafuncGetData(lstr)
'            Next i

'
'            '处理正文，第二节内容
            
'            Set lobjRec = dafuncGetData("select distinct a.系统编号,a.姓名,a.危害因素,b.体检项目 as " _
'                & "复查内容,a.体检结论+a.诊断和处理意见 as 复查意见 from 职业病体检_体检基本数据库 a,(select 系统编号,体检项目=dbo.z_fc(系统编号) from " _
'                & "职业病体检_体检结果视图 where 单项结论 = '不合格') b where a.系统编号 = b.系统编号 and 单位名称 = '" & lcolFactor("单位名称") & "'" _
'                & " and (体检日期 >= '" & lcolFactor("开始日期") & " 00:00:00' and 体检日期 <='" & lcolFactor("截止日期") & " 23:59:59') and 体检状态 in( 7,6,5)")
'            Set lobjRec = dafuncGetData("select distinct a.系统编号,a.姓名,a.危害因素,a.诊断和处理意见 from 职业病体检_体检基本数据库 a where 单位名称 = '" & lcolFactor("单位名称") & "'" _
                & " and (体检日期 >= '" & lcolFactor("开始日期") & " 00:00:00' and 体检日期 <='" & lcolFactor("截止日期") & " 23:59:59') and 体检状态 in( 7,6,5)")
            
            
          '  查询人员信息填入表格  2015-10-30  by 牟俊
             Set lobjRec = dafuncGetData("select convert(varchar(100),体检日期,111) as 体检日期,系统编号,姓名,性别,年龄,工龄,现工种,危害因素,体检结论,诊断和处理意见 from 职业病体检_体检基本数据库  where 单位名称 = '" & lcolFactor("单位名称") & "'" _
                & " and (体检日期 >= '" & lcolFactor("开始日期") & " 00:00:00' and 体检日期 <='" & lcolFactor("截止日期") & " 23:59:59') and 体检状态 >=1 ")

'             Set lobjRec = dafuncGetData("select 体检日期,系统编号,姓名,性别,年龄,工龄,现工种,危害因素,体检结论,诊断和处理意见 from 职业病体检_体检基本数据库  where 单位名称 = '" & lcolFactor("单位名称") & "'" _
'                & " and (体检日期 >= '" & Format(" & lcolFactor("开始日期") & ", "yyyy-mm-dd") & "' and 体检日期 <='" & Format(" & lcolFactor("截止日期") & ", "yyyy-mm-dd") & "') and 体检状态 >=1 ")


'            Set lobjRec = dafuncGetData("select distinct b.体检日期,a.系统编号,a.姓名,a.性别,a.年龄,a.工龄,a.现工种,a.危害因素,b.体检结论,b.诊断和处理意见 from 职业病体检_体检人员基本信息表 a ,职业病体检_体检基本信息表 b where 单位名称 = '" & lcolFactor("单位名称") & "'" _
'                & " and (体检日期 >= '" & lcolFactor("开始日期") & " 00:00:00' and 体检日期 <='" & lcolFactor("截止日期") & " 23:59:59') and 体检状态 >=1 ")

            Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
            Set myTable = myRange.Tables(1)
            If myTable.rows.Count <= lobjRec.RecordCount Then
                j = myTable.rows.Count
                myTable.rows(j).Select
                objWordDocument.ActiveWindow.Selection.InsertRows lobjRec.RecordCount - j + 1
                For i = 1 To lobjRec.RecordCount - myTable.rows.Count
                    myTable.rows.Add (myTable.rows(j))
                Next
            End If

            For i = 1 To lobjRec.RecordCount
                For j = 1 To lobjRec.Fields.Count
                    myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec(j - 1)), "", lobjRec(j - 1))
                Next
                lobjRec.MoveNext
            Next
            
            '增加复查人员一览表并填入表格  2015-11-13 by 牟俊 ↓
            Dim lobjRec2 As Object
            Dim yijian As String
            Set lobjRec2 = dafuncGetData("select 系统编号,姓名,危害因素,复查原因,诊断和处理意见 from 职业病体检_体检基本数据库  where 单位名称 = '" & lcolFactor("单位名称") & "'" _
                & " and (体检日期 >= '" & lcolFactor("开始日期") & " 00:00:00' and 体检日期 <='" & lcolFactor("截止日期") & " 23:59:59') and 体检状态 >=1 and 诊断和处理意见 like '%再做职业健康评价%'")
                Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
                Set myTable = myRange.Tables(2)
                If myTable.rows.Count <= lobjRec2.RecordCount Then
                    j = myTable.rows.Count
                    myTable.rows(j).Select
                    objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
                    For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                        myTable.rows.Add (myTable.rows(j))
                    Next
                End If

                For i = 1 To lobjRec2.RecordCount
                    For j = 1 To lobjRec2.Fields.Count
                        myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
                    Next
                    lobjRec2.MoveNext
                Next

           '2015-11-13 by 牟俊 ↑
            

        '更新文档正文（不包括页眉、页脚）中的域对象，主要是页码、页数
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
        '保存文件
        objWordDocument.SaveAs lstrNewDoc
        objWordDocument.Saved = False

        With objWord.ActiveWindow.View
            .ShowRevisionsAndComments = False
            .RevisionsView = 0          'wdRevisionsViewFinal
        End With
    
        objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9
        objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0
    
        objWord.Visible = True
    
        On Error GoTo errHandler
    
        objWordDocument.Activate
        objWord.Activate
        
        
        
        
'            '传递对象参数，以便保存文件到数据库。
'        On Error Resume Next
'        If Not paraReadOnly Then
'            If lintRepID = 0 Then
'                objWord.Run "subStart", paraParent, -1, "", ""
'            Else
'                objWord.Run "subStart", paraParent, lintRepID, "", ""
'            End If
'        Else
'            If lintRepID = 0 Then
'                objWord.Run "subStart", Nothing, -1, "", ""
'            Else
'                objWord.Run "subStart", Nothing, lintRepID, "", ""
'            End If
'        End If
'        If Err.Number = 450 Then     '参数不对，说明该模板没有增加样品种类的参数
'            If Not paraReadOnly Then
'                If lintRepID = 0 Then
'                    objWord.Run "subStart", paraParent, -1, ""
'                Else
'                    objWord.Run "subStart", paraParent, lintRepID, ""
'                End If
'            Else
'                If lintRepID = 0 Then
'                    objWord.Run "subStart", Nothing, -1, ""
'                Else
'                    objWord.Run "subStart", Nothing, lintRepID, ""
'                End If
'            End If
'        End If
'        If Err.Number = 438 Then
'            MsgBox "该报告的模板没有按照规定编写宏代码subStart，将导致无法保存到数据库里。", vbOKOnly + vbCritical, "系统提示"
'        End If
        Exit Sub
    
errHandler:
    If Err = 3001 Then
        MsgBox "没有在数据库中找到该Word报告的具体文件，可能是保存该报告到系统中时网络发生故障，请在系统中删除该报告信息后，重新录入该报告。", vbInformation, "系统提示"
    Else
        sfsub错误处理 "界面部件", "mod检验界面", "sub编辑总检报告", Err.Number, Err.Description, True
    End If
    Exit Sub
    Resume
End Sub
Sub sub加载普通性工作人员健康word信息(objWordDocument As Object, myRange As Object, paraSysNo As String)
      Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  '每节一个object,模板文件共5节。
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr体检日期 As String
    Dim lstr最终结论, lstr体检建议 As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '添加和更换体检日期
    strSQL = "select 体检日期 from 职业病体检_体检基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr体检日期 = Format(IIf(IsNull(lobjrec0("体检日期")), Now, lobjrec0("体检日期")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="【体检日期】", ReplaceWith:=lstr体检日期, Replace:=2    'wdReplaceAll

    '处理表头，第一节内容。
    strSQL = "select 系统编号,姓名,单位名称,住址,电话号码,邮编,建档日期 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(1).Headers.Count
'            If lobjrec1.Fields(i).Name = "建档日期" Then
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjrec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", Format(lobjrec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
'            Else
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjrec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", lobjrec1(i)), Replace:=2    'wdReplaceAll
'            End If
'        Next
   
         If lobjRec1.Fields(i).Name = "建档日期" Then
              objWordDocument.Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
         Else
             objWordDocument.Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
         End If
         
    Next
    '1
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    
    '第二节内容，基本信息和职业病史信息。
    strSQL = "select 性别,民族,出生日期,出生地,文化程度 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjrec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next

 
         objWordDocument.Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

     
    Next
    strSQL = "select * from 职业病体检_个人生活史表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
        myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjrec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next
        
        objWordDocument.Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

        Else
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    '2
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents

    '第三节内容，体检结果和体检医师填充。
'    strSQL = "select * from 职业病体检_体检结果视图 where 系统编号='" & paraSysNo & "'"
    strSQL = "select b.编码 as 体检项目,isnull(体检结果,'')as 体检结果,系统编号 from 职业病体检_体检结果视图 a right join 职业病体检_体检项目设置表 b on a.体检项目=b.编码 and 系统编号='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
'        Next

            objWordDocument.Range.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
    
       '3
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    
    strSQL = "select a.科室,b.姓名 医师姓名 from 职业病体检_科室结论表 a, 系统管理_员工基本信息表 b where a.系统编号='" & paraSysNo & "' and a.科室<>'06' and a.医生编号=b.编号"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2   'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2    'wdReplaceAll
'        Next

        objWordDocument.Range.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2   'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
    '4
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
dasubSetQueryTimeout 600
    '第四节内容，最终结果和结论、建议填充。
    strSQL = "select * from 职业病体检_科室结论表 where 系统编号='" & paraSysNo & "' and 科室='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("文字结论"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr最终结论 = lstrTmp(0)
            lstr体检建议 = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="【最终结论】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(4).Headers.Count
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【最终结果】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
'        Next

         objWordDocument.Range.Find.Execute FindText:="【最终结果】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
         objWordDocument.Range.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
    End If
    '5
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    '替换五官科眼睛检查图片
    '先在c盘下复制一份，内容如何无所谓
'    Dim lobjSys As Object
'    Set lobjSys = CreateObject("Scripting.FileSystemObject")
'    lobjSys.copyfile App.Path & "\晶状体环状及正面图.bmp", "c:\晶状体环状及正面图.bmp"
'
'    '从数据库中复制出该体检人员的图片，覆盖c盘相应图片
'    Set lobjSys = CreateObject("职业病体检结果录入.ClsCommon")
'    frmFinalConclusion.libPicture.AutoRedraw = True
'    frmFinalConclusion.libPicture.Picture = lobjSys.func获取结果图片(paraSysNo, "01069", "晶状体环状及正面图.bmp")
'    SavePicture frmFinalConclusion.libPicture.Picture, "c:\晶状体环状及正面图.bmp"
 
    
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
'    Set lobjSys = Nothing
    Exit Sub
errHandler:
End Sub
'2012-08-20 于登淼
'添加填入word报告中数据的代码，涉核部队、8023部队
Sub sub加载放射工作人员健康word信息(objWordDocument As Object, myRange As Object, paraSysNo As String)
    Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  '每节一个object,模板文件共5节。
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr体检日期 As String
    Dim lstr最终结论, lstr体检建议 As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '添加和更换体检日期
    strSQL = "select 体检日期 from 职业病体检_体检基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr体检日期 = Format(IIf(IsNull(lobjrec0("体检日期")), Now, lobjrec0("体检日期")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="【体检日期】", ReplaceWith:=lstr体检日期, Replace:=2    'wdReplaceAll

    '处理表头，第一节内容。
    strSQL = "select 系统编号,姓名,单位名称,住址,电话号码,邮编,建档日期 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(1).Headers.Count
'            If lobjrec1.Fields(i).Name = "建档日期" Then
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjrec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", Format(lobjrec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
'            Else
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjrec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", lobjrec1(i)), Replace:=2    'wdReplaceAll
'            End If
'        Next
   
         If lobjRec1.Fields(i).Name = "建档日期" Then
              objWordDocument.Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
         Else
             objWordDocument.Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
         End If
         
    Next
    
       '1
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '第二节内容，基本信息和职业病史信息。
    strSQL = "select 性别,民族,出生日期,出生地,文化程度 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjrec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next

 
         objWordDocument.Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

     
    Next
    strSQL = "select * from 职业病体检_个人生活史表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
        myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjrec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next
        
        objWordDocument.Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

        Else
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next

           '2
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    '第三节内容，体检结果和体检医师填充。
'    strSQL = "select * from 职业病体检_体检结果视图 where 系统编号='" & paraSysNo & "'"
    strSQL = "select b.编码 as 体检项目,isnull(体检结果,'')as 体检结果,系统编号 from 职业病体检_体检结果视图 a right join 职业病体检_体检项目设置表 b on a.体检项目=b.编码 and 系统编号='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
'        Next

            objWordDocument.Range.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
    strSQL = "select a.科室,b.姓名 医师姓名 from 职业病体检_科室结论表 a, 系统管理_员工基本信息表 b where a.系统编号='" & paraSysNo & "' and a.科室<>'06' and a.医生编号=b.编号"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2   'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2    'wdReplaceAll
'        Next

        objWordDocument.Range.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2   'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
           '3
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
        
        
    '第四节内容，最终结果和结论、建议填充。
    strSQL = "select * from 职业病体检_科室结论表 where 系统编号='" & paraSysNo & "' and 科室='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("文字结论"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr最终结论 = lstrTmp(0)
            lstr体检建议 = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="【最终结论】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(4).Headers.Count
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【最终结果】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
'        Next

         objWordDocument.Range.Find.Execute FindText:="【最终结果】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
         objWordDocument.Range.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
    End If
    
       '4
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    '替换五官科眼睛检查图片
    '先在c盘下复制一份，内容如何无所谓
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    lobjSys.copyfile App.Path & "\晶状体环状及正面图.bmp", "c:\晶状体环状及正面图.bmp"
    
    '从数据库中复制出该体检人员的图片，覆盖c盘相应图片
    Set lobjSys = CreateObject("职业病体检结果录入.ClsCommon")
    frmFinalConclusion.libPicture.AutoRedraw = True
    frmFinalConclusion.libPicture.Picture = lobjSys.func获取结果图片(paraSysNo, "01069", "晶状体环状及正面图.bmp")
    SavePicture frmFinalConclusion.libPicture.Picture, "c:\晶状体环状及正面图.bmp"
    
       '5
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
    
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
    Set lobjSys = Nothing
    Exit Sub
errHandler:
'    MsgBox ("sdgdg")
End Sub
Sub sub加载8023和放射性工作人员word信息(objWordDocument As Object, myRange As Object, paraSysNo As String)
     Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  '每节一个object,模板文件共5节。
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr体检日期 As String
    Dim lstr最终结论, lstr体检建议 As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '添加和更换体检日期
    strSQL = "select 体检日期 from 职业病体检_体检基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr体检日期 = Format(IIf(IsNull(lobjrec0("体检日期")), Now, lobjrec0("体检日期")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="【体检日期】", ReplaceWith:=lstr体检日期, Replace:=2    'wdReplaceAll

    '处理表头，第一节内容。
    strSQL = "select 系统编号,姓名,单位名称,住址,电话号码,邮编,建档日期 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(1).Headers.Count
            If lobjRec1.Fields(i).Name = "建档日期" Then
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
            Else
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
            End If
        Next
    Next
    '1
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    '第二节内容，基本信息和职业病史信息。
    strSQL = "select 性别,民族,出生日期,出生地,文化程度 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(2).Headers.Count
            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        Next
    Next
    strSQL = "select * from 职业病体检_个人生活史表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            For j = 1 To objWordDocument.Sections(2).Headers.Count
                objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            Next
        Else
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    '2
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
dasubSetQueryTimeout 600
    '第三节内容，体检结果和体检医师填充。
    strSQL = "select * from 职业病体检_体检结果视图 where 系统编号='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    '3
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        
    strSQL = "select a.科室,b.姓名 医师姓名 from 职业病体检_科室结论表 a, 系统管理_员工基本信息表 b where a.系统编号='" & paraSysNo & "' and a.科室<>'06' and a.医生编号=b.编号"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2   'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    '4
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        
    '第四节内容，最终结果和结论、建议填充。
    strSQL = "select * from 职业病体检_科室结论表 where 系统编号='" & paraSysNo & "' and 科室='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("文字结论"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr最终结论 = lstrTmp(0)
            lstr体检建议 = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="【最终结论】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(4).Headers.Count
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【最终结果】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
        Next
    End If
    
    '5
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        
    '替换五官科眼睛检查图片
    '先在c盘下复制一份，内容如何无所谓
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    lobjSys.copyfile App.Path & "\晶状体环状及正面图.bmp", "c:\晶状体环状及正面图.bmp"
    
    '从数据库中复制出该体检人员的图片，覆盖c盘相应图片
    Set lobjSys = CreateObject("职业病体检结果录入.ClsCommon")
    frmFinalConclusion.libPicture.AutoRedraw = True
    frmFinalConclusion.libPicture.Picture = lobjSys.func获取结果图片(paraSysNo, "01069", "晶状体环状及正面图.bmp")
    SavePicture frmFinalConclusion.libPicture.Picture, "c:\晶状体环状及正面图.bmp"
    
    '6
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
       DoEvents
        
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
    Set lobjSys = Nothing
    Unload frmProcess
    Exit Sub
errHandler:
End Sub
'添加填入word报告中数据的代码，放射工作类
Sub sub加载涉核工作人员职业健康word信息(objWordDocument As Object, myRange As Object, paraSysNo As String)
     Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  '每节一个object,模板文件共5节。
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr体检日期 As String
    Dim lstr最终结论, lstr体检建议 As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '添加和更换体检日期
    strSQL = "select 体检日期 from 职业病体检_体检基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr体检日期 = Format(IIf(IsNull(lobjrec0("体检日期")), Now, lobjrec0("体检日期")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="【体检日期】", ReplaceWith:=lstr体检日期, Replace:=2    'wdReplaceAll

    '处理表头，第一节内容。
    strSQL = "select 系统编号,姓名,单位名称,住址,电话号码,邮编,建档日期 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(1).Headers.Count
            If lobjRec1.Fields(i).Name = "建档日期" Then
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
            Else
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
            End If
        Next
    Next
    '1
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '第二节内容，基本信息和职业病史信息。
    strSQL = "select 性别,民族,出生日期,出生地,文化程度 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(2).Headers.Count
            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        Next
    Next
    strSQL = "select * from 职业病体检_个人生活史表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            For j = 1 To objWordDocument.Sections(2).Headers.Count
                objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            Next
        Else
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    
       '2
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '第三节内容，体检结果和体检医师填充。
    strSQL = "select * from 职业病体检_体检结果视图 where 系统编号='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    strSQL = "select a.科室,b.姓名 医师姓名 from 职业病体检_科室结论表 a, 系统管理_员工基本信息表 b where a.系统编号='" & paraSysNo & "' and a.科室<>'06' and a.医生编号=b.编号"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2   'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    
       '3
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '第四节内容，最终结果和结论、建议填充。
    strSQL = "select * from 职业病体检_科室结论表 where 系统编号='" & paraSysNo & "' and 科室='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("文字结论"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr最终结论 = lstrTmp(0)
            lstr体检建议 = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="【最终结论】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(4).Headers.Count
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【最终结果】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
        Next
    End If
    
       '4
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
    
    '替换五官科眼睛检查图片
    '先在c盘下复制一份，内容如何无所谓
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    lobjSys.copyfile App.Path & "\晶状体环状及正面图.bmp", "c:\晶状体环状及正面图.bmp"
    
    '从数据库中复制出该体检人员的图片，覆盖c盘相应图片
    Set lobjSys = CreateObject("职业病体检结果录入.ClsCommon")
    frmFinalConclusion.libPicture.AutoRedraw = True
    frmFinalConclusion.libPicture.Picture = lobjSys.func获取结果图片(paraSysNo, "01069", "晶状体环状及正面图.bmp")
    SavePicture frmFinalConclusion.libPicture.Picture, "c:\晶状体环状及正面图.bmp"
    
       '5
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
    
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
    Set lobjSys = Nothing
    Exit Sub
errHandler:
'    MsgBox ("sdgdg")
End Sub

'2012-08-20 于登淼 修改：罗李奎 时间：2013-1-9
'添加填入word报告中数据的代码，职业健康类
Sub sub加载职业健康word信息(objWordDocument As Object, myRange As Object, paraSysNo As String)
    Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  '每节一个object,模板文件共5节。
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr体检日期 As String
    Dim lstr最终结论, lstr体检建议 As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '添加和更换体检日期
    strSQL = "select 体检日期 from 职业病体检_体检基本信息表 where 系统编号='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr体检日期 = Format(IIf(IsNull(lobjrec0("体检日期")), Now, lobjrec0("体检日期")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="【体检日期】", ReplaceWith:=lstr体检日期, Replace:=2    'wdReplaceAll

    '处理表头，第一节内容。
    strSQL = "select 系统编号,姓名,性别,出生日期,文化程度,单位名称,电话号码,公民身份号码,工龄,现工种,职业危害工龄,危害因素,建档日期 from 职业病体检_体检人员基本信息表  where 系统编号='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(1).Headers.Count
            If lobjRec1.Fields(i).Name = "建档日期" Then
                objWordDocument.Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
            Else
                objWordDocument.Range.Find.Execute FindText:="【" & lobjRec1.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
            End If
'        Next
    Next
    '1
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    '第二节内容，基本信息和职业病史信息。
'    strSQL = "select 性别,民族,出生日期,出生地,文化程度 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
'    Set lobjrec2 = dafuncGetData(strSQL)
'    For i = 0 To lobjrec2.Fields.Count - 1
'        myRange.Find.Execute FindText:="【" & lobjrec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
''        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Range.Find.Execute FindText:="【" & lobjrec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
''        Next
'    Next
    
    strSQL = "select 起始时间,工作单位,工种,危害种类,防护措施 from 职业病体检_职业史表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
'        objWordDocument.Sections(2).Range.Find.Execute FindText:="【职业史】", ReplaceWith:="", Replace:=2
        Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
        Set myTable = myRange.Tables(3)
        If myTable.rows.Count < lobjRec2.RecordCount Then
            j = myTable.rows.Count
            myTable.rows(j).Select
            objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
            For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                myTable.rows.Add (myTable.rows(j))
            Next
        End If
        
        For i = 1 To lobjRec2.RecordCount
            For j = 1 To lobjRec2.Fields.Count
                myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
            Next
            lobjRec2.MoveNext
        Next
        
    End If
    
       strSQL = "select * from 职业病体检_个人生活史表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
        myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        Next
        Else
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
         strSQL = "select 出生地 from 职业病体检_体检人员基本信息表  where 系统编号='" & paraSysNo & "'"
         Set lobjRec2 = dafuncGetData(strSQL)
         For i = 0 To lobjRec2.Fields.Count - 1
            If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
                myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
                objWordDocument.Range.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            Else
            myRange.Find.Execute FindText:="【" & lobjRec2.Fields(i).Name & "】", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    
    strSQL = "select 编号,疾病名称,诊断日期,诊断单位,治疗经过,转归 from 职业病体检_既往病史表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
'        objWordDocument.Sections(2).Range.Find.Execute FindText:="【职业史】", ReplaceWith:="", Replace:=2
        Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
        Set myTable = myRange.Tables(4)
        If myTable.rows.Count < lobjRec2.RecordCount Then
            j = myTable.rows.Count
            myTable.rows(j).Select
            objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
            For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                myTable.rows.Add (myTable.rows(j))
            Next
        End If
        
        For i = 1 To lobjRec2.RecordCount
            For j = 1 To lobjRec2.Fields.Count
                myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
            Next
            lobjRec2.MoveNext
        Next
        
    End If
    strSQL = "select * from 职业病体检_自觉症状表 where 系统编号='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
'        objWordDocument.Sections(2).Range.Find.Execute FindText:="【职业史】", ReplaceWith:="", Replace:=2
        Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
        Set myTable = myRange.Tables(6)
        If myTable.rows.Count < lobjRec2.RecordCount Then
            j = myTable.rows.Count
            myTable.rows(j).Select
            objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
            For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                myTable.rows.Add (myTable.rows(j))
            Next
        End If
        
        For i = 1 To lobjRec2.RecordCount
            For j = 1 To lobjRec2.Fields.Count
                myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
            Next
            lobjRec2.MoveNext
        Next
        
    End If
     '2
        frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    
    '第三节内容，体检结果和体检医师填充。
'    strSQL = "select * from 职业病体检_体检结果视图 where 系统编号='" & paraSysNo & "'"
    strSQL = "select b.编码 as 体检项目,isnull(体检结果,'')as 体检结果,系统编号 from 职业病体检_体检结果视图 a right join 职业病体检_体检项目设置表 b on a.体检项目=b.编码 and 系统编号='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="【" & lobjrec3("体检项目") & "】", ReplaceWith:=IIf(IsNull(lobjrec3("体检结果")), "", lobjrec3("体检结果")), Replace:=2    'wdReplaceAll
'        Next
        lobjrec3.MoveNext
    Next
    '3
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    strSQL = "select a.科室,b.姓名 医师姓名 from 职业病体检_科室结论表 a, 系统管理_员工基本信息表 b where a.系统编号='" & paraSysNo & "' and a.科室<>'06' and a.医生编号=b.编号"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2   'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="【" & lobjrec3("科室") & "体检医师】", ReplaceWith:=IIf(IsNull(lobjrec3("科室")), "", lobjrec3("医师姓名")), Replace:=2    'wdReplaceAll
'        Next
        lobjrec3.MoveNext
    Next
    '4
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
         
    '替换没有的科室
    For i = 1 To 17
        If Not i = 16 Then
            myRange.Find.Execute FindText:="【0" & i & "体检医师】", ReplaceWith:="", Replace:=2   'wdReplaceAll
        End If
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="【0" & i & "体检医师】", ReplaceWith:="", Replace:=2    'wdReplaceAll
'        Next
'        lobjrec3.MoveNext
    Next
        
    '第四节内容，最终结果和结论、建议填充。
    strSQL = "select * from 职业病体检_科室结论表 where 系统编号='" & paraSysNo & "' and 科室='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("文字结论"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr最终结论 = lstrTmp(0)
            lstr体检建议 = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="【最终结论】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(4).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="【最终结果】", ReplaceWith:=lstr最终结论, Replace:=2    'wdReplaceAll
            objWordDocument.Range.Find.Execute FindText:="【体检建议】", ReplaceWith:=lstr体检建议, Replace:=2    'wdReplaceAll
'        Next
    End If
    '5
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    '替换五官科眼睛检查图片
    '先在c盘下复制一份，内容如何无所谓
'    Dim lobjSys As Object
'    Set lobjSys = CreateObject("Scripting.FileSystemObject")
'    lobjSys.copyfile App.Path & "\晶状体环状及正面图.bmp", "c:\晶状体环状及正面图.bmp"
'
'    '从数据库中复制出该体检人员的图片，覆盖c盘相应图片
'    Set lobjSys = CreateObject("职业病体检结果录入.ClsCommon")
'    frmFinalConclusion.libPicture.AutoRedraw = True
'    frmFinalConclusion.libPicture.Picture = lobjSys.func获取结果图片(paraSysNo, "01069", "晶状体环状及正面图.bmp")
'    SavePicture frmFinalConclusion.libPicture.Picture, "c:\晶状体环状及正面图.bmp"
'
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
'    Set lobjSys = Nothing

    Exit Sub
   
errHandler:
'    MsgBox ("sdgdg")
End Sub
'作者：罗李奎 时间:2013-1-8 ↓
'取出选择对应信息的word模版
Sub sub取出word模版()
    Dim lstrFile As String
    Dim lobjRec As Object
    
    lstrFile = Dir(App.Path & "\通用_*.dot")
    Do While lstrFile <> ""
'       clstFile.AddItem lstrFile
       lstrFile = Dir
    Loop
    '寻找当前用户所属科室所对应的专用模板名前缀
    Set lobjRec = dafuncGetData("select 描述 from 系统管理_科室字典表 where 编号='" & um用户所属科室编号 & "'")

'        lstrFile = Dir(App.Path & "\" & lobjRec(0) & "_四川省" & Left(pstrWordname, 2) & "*.dot")
        lstrFile = Dir(App.Path & "\职业病体检_四川省" & Left(pstrWordname, 2) & "*.dot")

    If lstrFile = "" Then
        MsgBox "没有找到Word模板文件！", vbInformation, "系统提示"
       Exit Sub
    Else
        pstrFilename = lstrFile
  End If
End Sub
'把文档保存到数据库
'word文档关闭时会触发该方法的调用。para系统编号
Public Sub subSaveDoc(ByVal paraFile As String, ByVal paraNo As Integer, ByVal para系统编号 As String)
    '保存基本信息。
    Dim lobjRec As Object
    Dim lstrFileType As String, lstrNo As String
    Dim i As Integer, lstrType As String
    Dim lstr体检表编号 As String
    Dim lstr体检类型 As String
    Dim lstr体检类别 As String
    
'    If mbln只读 Then Exit Sub
    
    On Error GoTo errHandler
    
    For i = Len(paraFile) To 1 Step -1
        If Mid(paraFile, i, 1) = "." Then Exit For
    Next
    lstrFileType = Mid(paraFile, i + 1)
    '寻找当前用户所能操作的检验类别
'    Set lobjRec = dafuncGetData("select 名称 from 检验管理_检验分工类别视图 where 员工编号='" & um用户编号 & "' order by 编号")
'    If lobjRec.RecordCount > 0 Then lstrType = lobjRec(0)
    
'    Set lobjRec = dafuncGetData("select 编号 from 职业病体检_体检报告信息表 where 编号=" & paraNo & " and 类别='" & lstrType & "'")
    dasubBeginTran
    Set lobjRec = dafuncGetData("select 报告编号 from 职业病体检_体检报告信息表 where 系统编号='" & para系统编号 & "'")
    If lobjRec.RecordCount = 0 Then
        Set lobjRec = dafuncGetData("select 体检表编号,体检类型,体检类别 from dbo.职业病体检_体检基本信息表  where 系统编号='" & para系统编号 & "'")
      If lobjRec.RecordCount <> 0 Then
            lstr体检类型 = lobjRec(1)
            lstr体检类别 = lobjRec(2)
            lstr体检表编号 = lobjRec(0)
    Else
        MsgBox "该信息的基本信息不存在，数据无法保存！", vbInformation, "系统提示！"
        Exit Sub
    End If
        dafuncGetData "insert into 职业病体检_体检报告信息表(系统编号,报告编号,报告类别,文件类型,创建人,体检类型,体检类别,体检表编号,修改日期) values('" & para系统编号 & "', '" & para系统编号 & "' ,'结果','" & UCase(lstrFileType) & "','" & um用户编号 & "', '" & lstr体检类型 & "','" & lstr体检类别 & "','" & lstr体检表编号 & "',getdate() " & ")"
        '寻找新保存的文件的编号
'        Set lobjRec = dafuncGetData("select max(编号) from 职业病体检_体检报告信息表 where 报告编号='" & para报告编号 & "' and 创建人='" & um用户编号 & "'")
        Set lobjRec = dafuncGetData("select max(系统编号) from 职业病体检_体检报告信息表 where 报告编号='" & para系统编号 & "'")
        
        lstrNo = lobjRec(0)
    Else
'        dafuncGetData "update 职业病体检_体检报告信息表 set 文件类型='" & UCase(lstrFileType) & "',修改日期=getdate(),创建人='" & um用户编号 & "' where 编号=" & paraNo
'        lstrNo = CStr(paraNo)

         dafuncGetData "update 职业病体检_体检报告信息表 set 文件类型='" & UCase(lstrFileType) & "',修改日期=getdate(),创建人='" & um用户编号 & "' where 报告编号='" & para系统编号 & "'"
        lstrNo = CStr(para系统编号)
    End If
    
    '保存文档文件到数据库。
     pobjFileToDatabase.subFileToColumn "职业病体检_体检报告信息表", "报告", "报告编号=" & lstrNo, paraFile
'     pobjFileToDatabase.subFileToColumn "系统管理_检测检验报告信息表", "文档", "编号=" & lstrNo, paraFile
    dafuncGetData "update 职业病体检_体检基本信息表 set 体检状态=7 where 系统编号='" & para系统编号 & "'"
   dasubCommitTran
   
    On Error Resume Next
'    oeExamSubSave "检验报告编制", frm编制报告.pstr受理编号, "编制报告"
'    frmFinalConclusion.subRefreshView
    Exit Sub
errHandler:
    sfsub错误处理 "界面部件", "mod检验界面", "subSave", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
''供WORD宏调用，用于保存WORD至数据库
'Public Sub subSave(ByVal paraFile As String, ByVal paraNo As Integer, ByVal para报告编号 As String)
'    subSaveDoc paraFile, paraNo, para报告编号
'End Sub

'功能：为收检信息浏览界面获取指定条件的收检信息。
Public Function func获取收检浏览信息(ByVal paraOffice As String, Optional ByVal paraFilter As String) As Recordset
    
    On Error GoTo errHandle
    
    dasubSetQueryTimeout 600
    
'    Set func获取收检浏览信息 = dafuncGetData("exec 职业病体检_查询体检浏览信息 '" & paraOffice & "','" & paraFilter & "','" & um用户编号 & "'")
     Set func获取收检浏览信息 = dafuncGetData(" select 系统编号,体检表编号,体检类型,体检类别,体检日期,体检状态 from 职业病体检_体检基本信息表 where 1=1  order by 体检状态 ")
    
    Exit Function
    
errHandle:

    sfsub错误处理 "职业病界面", "modMain", "func获取收检浏览信息", Err.Number, Err.Description, True
        
End Function
Public Sub sub读取word文档(paraParent As Object, ByVal para系统编号 As String, ByVal para报告编号 As String, ByVal paraReadOnly As Boolean)

    Dim objWord As Object 'Word.Application
    Dim objWordDocument As Object 'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer        '已有Word报告的ID
    
    On Error GoTo errHandler
    
    '启动word。
    On Error Resume Next
    
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            On Error GoTo errHandler
            Err.Raise 6666, , "你没有安装Word，无法编辑报告！请先安装MS Office 2000以上版本。 "
        End If
    End If
    
    objWord.UserName = um用户名
    objWord.Options.UpdateLinksAtOpen = False       '不让Word提示用户更新签名图片
    objWord.Options.CheckGrammarAsYouType = False   '禁止拼写检查和语法检查
    objWord.Options.CheckSpellingAsYouType = False
    
    On Error GoTo errHandler
    
    Dim lstrNewDoc As String
    Dim lstrDotFile As String '模板文件。
    Dim j As Integer
    

    Dim lpicPhoto As StdPicture
    Dim lobjSys As Object '
    Dim lstr受理编号 As String
    Dim i As Long, lstr样品种类 As String
    
    Set lobjSys = CreateObject("Scripting.FileSystemObject")

    
    Dim lstr生产日期 As String      '个人剂量使用
    Dim lsngHeight As Single
    
    '判断文书是否已存在。
    Set lobjRec = dafuncGetData("select 报告编号,文件类型 from 职业病体检_体检报告信息表 where 报告编号='" & para报告编号 & "' and 系统编号='" & para系统编号 & "'")
    If lobjRec.RecordCount = 0 Then
        If paraReadOnly Then        '不允许修改，表明为查看操作，不是新增操作
            MsgBox "该样品没有录入Word报告，您此时不能再为其添加Word报告！", vbOKOnly + vbInformation, "系统提示"
            Exit Sub
        End If
        
    Else
        lintRepID = CInt(lobjRec(0))
        '编辑已有文书。
        lstrNewDoc = App.Path & "\temp\" & lstr受理编号 & "_" & Format(Now, "yymmddhhmmss") & "." & lobjRec(1)
        '取出文档。
        pobjFileToDatabase.subColumnToFile "职业病体检_体检报告信息表", "文档", "编号=" & lobjRec(0), lstrNewDoc
        
        '直接打开已有的文档
        Set objWordDocument = objWord.Documents.Open(FileName:=lstrNewDoc, ReadOnly:=paraReadOnly)
        
        On Error Resume Next
        '更新域
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
       
         
    End If
    With objWord.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = 0      'wdRevisionsViewFinal
    End With
    
'    If objWord.Version = 11 Then
'    objWordDocument.CommandBars("Reviewing").Controls(11).Enabled = False    '工具栏上的“修订”按钮
'    For i = 1 To objWordDocument.CommandBars("Menu Bar").Controls(6).CommandBar.Controls.Count
'        If Left(objWordDocument.CommandBars("Menu Bar").Controls(6).CommandBar.Controls(i).Caption, 2) = "修订" Then
'            objWordDocument.CommandBars("Menu Bar").Controls(6).CommandBar.Controls(i).Enabled = False  '工具菜单上的“修订”命令
'        End If
'    Next
    '进入页眉编辑状态，然后退出编辑状态，以解决页眉上的横线的显示问题
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9       'wdSeekCurrentPageHeader
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0       'wdSeekMainDocument
'    End
'    Dim X As Object
'    For j = 1 To objWordDocument.CommandBars.Count
'        Set X = objWordDocument.CommandBars(j)
'        For i = 1 To X.Controls.Count
'            If left(X.Controls(i).Caption, 2) = "打印" Then
'                X.Controls(i).Visible = False
'            End If
'        Next
'    Next

    objWord.Visible = True
    
    On Error GoTo errHandler

    objWordDocument.Activate
    objWord.Activate
    
    'objWordDocument.Close
'    objWord.Quit
    
    If paraReadOnly Then
        If objWordDocument.Range.Fields.Count = 0 Then objWordDocument.Protect 3, , "cdc"     '如果文档为只读，保护该文档，不允许用户修改
        objWordDocument.Saved = True
    End If
    
         '显示修订痕迹
        'objWordDocument.ShowRevisions = True
        objWordDocument.TrackRevisions = True
        objWordDocument.Saved = True     '必须在此处设置，否则Saved=false
'    mbln只读 = paraReadOnly
    
    '传递对象参数，以便保存文件到数据库。
    On Error Resume Next
    If Not paraReadOnly Then
        If lintRepID = 0 Then
            objWord.Run "subStart", paraParent, -1, para报告编号, lstr样品种类
        Else
            objWord.Run "subStart", paraParent, lintRepID, para报告编号, lstr样品种类
        End If
    Else
        If lintRepID = 0 Then
            objWord.Run "subStart", Nothing, -1, para报告编号, lstr样品种类
        Else
            objWord.Run "subStart", Nothing, lintRepID, para报告编号, lstr样品种类
        End If
    End If
    If Err.Number = 450 Then     '参数不对，说明该模板没有增加样品种类的参数
        If Not paraReadOnly Then
            If lintRepID = 0 Then
                objWord.Run "subStart", paraParent, -1, para报告编号
            Else
                objWord.Run "subStart", paraParent, lintRepID, para报告编号
            End If
        Else
            If lintRepID = 0 Then
                objWord.Run "subStart", Nothing, -1, para报告编号
            Else
                objWord.Run "subStart", Nothing, lintRepID, para报告编号
            End If
        End If
    End If
    If Err.Number = 438 Then
        MsgBox "该报告的模板没有按照规定编写宏代码subStart，将导致无法保存到数据库里。", vbOKOnly + vbCritical, "系统提示"
    End If
    
    Kill "c:\检验报告*.bmp"
    Exit Sub
errHandler:
    If Err = 3001 Then
        MsgBox "没有在数据库中找到该Word报告的具体文件，可能是保存该报告到系统中时网络发生故障，请在系统中删除该报告信息后，重新录入该报告。", vbInformation, "系统提示"
    Else
        sfsub错误处理 "界面部件", "mod检验界面", "sub编辑word文档", Err.Number, Err.Description, True
    End If
    Exit Sub
    Resume
End Sub


Public Function func错误处理(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case 6
        func错误处理 = "输入数据过大，已超过系统规定大小。"
    Case -2147217833
        func错误处理 = "输入数据过长（或过大），已超过系统规定长度（或大小）。"
    Case -2147217913
        func错误处理 = "日期格式非法！"
    Case -2147217873 '外键不存在。
        func错误处理 = "系统服务继续处理。因为：" & Chr(13) & Chr(10) & "(1) 你正在保存的数据涉及的相关信息已被人删除！" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出本业务操作界面，重新进入。"
    Case 94 '无效使用Null。
        func错误处理 = "使用的字典项被人通过字典管理操作删除了，系统无法再继续正常处理。请找系统管理员恢复字典内容。请注意，不要随便删除字典项！"
    Case 336, 337, 338, 429, 430
        func错误处理 = "系统部件已损坏（或已丢失），系统无法再正常运行。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，并重新安装系统。"
    Case 440 '外部对象错误：类自动错误。
        func错误处理 = "系统部件不正常终止运行。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，重新进入。"
    Case 91 '对象没有初始化成功。
        func错误处理 = "因为网络阻塞，系统启动功能时无法完成正常的初始化。请退出功能界面，再重新进入功能界面。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，重新进入。"
    Case 5
        func错误处理 = "因为网络中断（或阻塞），系统无法正常运行。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请退出系统，重新进入。"
    Case Else
        func错误处理 = paraErrDes
    End Select
End Function


