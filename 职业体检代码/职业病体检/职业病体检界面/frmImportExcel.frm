VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportExcel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "批量导入"
   ClientHeight    =   7230
   ClientLeft      =   570
   ClientTop       =   900
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Left            =   10080
      Top             =   2040
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdDetails 
      Height          =   4935
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "双击单元格可修改内容，自动保存到EXCEL中"
      Top             =   2160
      Width           =   10455
      _cx             =   2088781833
      _cy             =   2088772097
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton ccmd单位定位 
      Caption         =   "单位定位"
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox ctxt单位名称 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox ccmbTemplate 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox ccmb体检人类别 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox ccmb体检人类型 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退 出"
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton ccmdImport 
      Caption         =   "确认导入"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   480
      Width           =   900
   End
   Begin MSComDlg.CommonDialog ccdg 
      Left            =   10200
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox ctxtDataPath 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.CommandButton ccmdSelect 
      Caption         =   "选择文件"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "注：文件格式正确后方可导入；每次只能导入文件一次，如果导入内容有错或文件重复导入，需到体检管理界面中删除。"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   9540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "单位名称："
      Height          =   180
      Left            =   7080
      TabIndex        =   13
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "黄色底为必录项"
      Height          =   180
      Left            =   7080
      TabIndex        =   11
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "体检表："
      Height          =   180
      Left            =   4080
      TabIndex        =   10
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "体检类别："
      Height          =   180
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "体检人员类型："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "导入文件信息： (可以双击修改，但修改后会自动覆盖至原来的Excel文件内容)"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   6375
   End
End
Attribute VB_Name = "frmImportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'2012-02-29 于登淼 添加整个窗体
'1、文件导入，只有体检人员基本信息和体检基本信息（不包括体检项目）的导入。
'2、做的并不完善的是：合法性检查、导入信息保存、单位信息（暂时只保存了单位名称）。
'3、尤其是导入信息保存，缺少体检项目初始化、试管编号等内容，等待其它模块的具体设定，该部分的补充待定。
'4、单位信息的内容是在体检之外的部件，单位档案→档案管理里设定。
'这里可以调用其内容，也可以手动输入。只是，如果调用除单位名称之外的信息，最好不要用手动输入。
'5、体检人类别和体检人类型，完全Copy体检登记部分的设定。那部分更改的话，这里也需要修改。
Private mobj体检, mobj体检表模板 As Object
Private lobj体检类型 As Object
Private lobj体检人类别 As Object
Private indextmp As Integer         '标记“公民身份号码”是哪一列，在保存excel文件时，强制设置该列的格式为字符串
Private lbol已经导入 As Boolean     '标记文件是否已经导入过一次。（因为觉得无法合理地设计重复导入的更新等问题，所以文件格式合法后只能导入一次。出错则返回管理界面修改。）
Private mcol体检项目 As New Collection  '新选择的体检项目
Private mcol收费项目 As New Collection  '新选择的收费项目

'选择体检表模板下拉列表
Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    MousePointer = 11
    subChangeTemplate       '选择体检表
    MousePointer = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmImportExcel", "ccmbTemplate_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


'全称为“体检类别”，字典内容：上岗前、在岗期间、离岗时、应急检查
Private Sub Ccmb体检人类别_Click()
    Dim lobj体检表模板集 As Object
    Dim lcolInfo As Collection
    Dim lcol体检表编号 As Collection
    Dim i As Integer
    On Error GoTo errHandler
    
    '将所有的非复查体检表模板加入到体检表下拉列表框中。再加体检人类别条件
    ccmbTemplate.Clear
    Set lobj体检表模板集 = CreateObject("职业病对象.ClsMedicalExamTemplateSet")
    
    
    '---------------
    '刘晗 2012-04-01 修改句为以下四句（注释两句，添加两句）
    
    'lobj体检表模板集.体检表类型 = 3
    'lobj体检表模板集.体检表类别 = ccmb体检人类别.ItemData(ccmb体检人类别.ListIndex)
    lobj体检表模板集.体检表类型 = Trim(ccmb体检人类型.Text)
    lobj体检表模板集.体检表类别 = Trim(Ccmb体检人类别.Text)
    
    
    
    Set lcolInfo = lobj体检表模板集.元素集
    Set lcol体检表编号 = lobj体检表模板集.体检表编号元素集
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
        ccmbTemplate.ItemData(ccmbTemplate.NewIndex) = lcol体检表编号(i)
    Next
    ccmbTemplate.Text = ccmbTemplate.List(0)
    
    Set lobj体检表模板集 = Nothing
    Set lcolInfo = Nothing
    Call ccmbTemplate_Click
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmImportExcel", "ccmb体检人类别_click", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'全称为“体检人员类型”，字典内容：职业健康、放射工作
Private Sub ccmb体检人类型_Click()
    On Error GoTo errHandler
    Set lobj体检类型 = CreateObject("职业病对象.clsmedicalexam")
    lobj体检类型.体检类型 = ccmb体检人类型.ItemData(ccmb体检人类型.ListIndex)
    'Call Ccmb体检人类别_Click
    'sub填充体检人类型List
    ccmbTemplate.Text = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmImportExcel", "ccmb体检人类型_Click", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

Private Sub ccmdExit_Click()
    Unload Me
    Set mobj体检 = Nothing
    Set mobj体检表模板 = Nothing
    Set lobj体检类型 = Nothing
    Set lobj体检人类别 = Nothing
    Set frmImportExcel = Nothing
    sub更新管理界面查询
End Sub

Private Sub ccmdImport_Click()
    Dim lcolTmp As Collection
    Dim lobjTmp As Object
    Dim i, j As Integer
    Dim lintRows As Integer
    On Error GoTo errHandler
    
    lintRows = cgrdDetails.rows - 2
    If lintRows < 1 Then
        MsgBox "没有可导入的信息，请查看EXCEL！", vbExclamation, "系统提示"
        Exit Sub
    End If
    If Trim(ccmb体检人类型.Text) = "" Then
        MsgBox "体检人员类型不能为空！", vbExclamation, "系统提示"
        Exit Sub
    End If
    If Trim(Ccmb体检人类别.Text) = "" Then
        MsgBox "体检类别不能为空！", vbExclamation, "系统提示"
        Exit Sub
    End If
    If Trim(ccmbTemplate.Text) = "" Then
        MsgBox "体检表不能为空！", vbExclamation, "系统提示"
        Exit Sub
    End If
    If Trim(ctxt单位名称.Text) = "" Then
        MsgBox "单位名称不能为空！", vbExclamation, "系统提示"
        Exit Sub
    End If
    Me.Enabled = False
    MousePointer = 11
    '显示进度。
    frmProcess.proPercent.Max = lintRows
    frmProcess.Label1.Caption = "正在导入，请等待..."
    frmProcess.proPercent.Value = 0
    frmProcess.Show 0, Me
    DoEvents
    
    Set lobjTmp = CreateObject("职业病对象.clsmedicalexam")
    For i = 2 To cgrdDetails.rows - 1
        Set lcolTmp = New Collection
        For j = 0 To cgrdDetails.cols - 1
            If cgrdDetails.TextMatrix(0, j) = "" Then GoTo NEXT_j
            If cgrdDetails.TextMatrix(0, j) = "姓名" And cgrdDetails.TextMatrix(i, j) = "" Then
                GoTo NEXT_i
            End If
            lcolTmp.Add cgrdDetails.TextMatrix(i, j), cgrdDetails.TextMatrix(0, j)
NEXT_j: Next j
        lcolTmp.Add lobjTmp.Func分配职业病体检系统编号 & (ccmb体检人类型.ListIndex + 1), "系统编号"
        sub导入文件内容 lcolTmp
        frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        cgrdDetails.RowHidden(i) = True
NEXT_i: Next i                   '想模仿着c语言里面的continue语句，结果就写成了goto
    Unload frmProcess
    
    
'    Exit Sub
errHandler:
    ccmdImport.Enabled = False
    ccmdSelect.Enabled = False
    Me.Enabled = True
    MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox ("导入内容失败!"), vbExclamation, "系统提示"
    Else
        MsgBox ("导入成功！"), vbInformation, "系统提示"
        lbol已经导入 = True
        frmRegisterManage.sub查询并显示
        Unload Me
        
    End If
End Sub

Private Sub ccmdSelect_Click()
    ccdg.Filter = "All Files (*.*)|*.*|Excel file" & _
            "(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
    ccdg.FileName = ""
    ccdg.ShowOpen
    sub显示导入信息
End Sub

Sub sub显示导入信息()
    Dim i As Integer
    Dim lstrTmp As String
    
    On Error GoTo errHandler
    
    ctxtDataPath.Text = ccdg.FileName
    With cgrdDetails
        .LoadGrid ccdg.FileName, flexFileExcel, 0
        .FormatString = cgrdDetails.TextMatrix(1, 0)
        For i = 1 To .cols - 1
'            .ColHidden(1) = True    '隐藏其它证件 2015-11-16 by 牟俊
            .FormatString = .FormatString & "|" & .TextMatrix(1, i)
        Next i
'        .ColHidden(1) = True    '隐藏其它证件 2015-11-16 by 牟俊
        .RowHidden(1) = True
        .AutoSize 0, .cols - 1, 0, 0
    End With

    '单条信息格式等判断，用不同颜色标出，修改后可改为正常颜色。不合格则不能导入。
    If lobl已经导入 = False Then ccmdImport.Enabled = sub导入信息合法性检查
    
    Exit Sub
errHandler:
    MsgBox ("显示导入信息错误！")
End Sub

Private Sub ccmd单位定位_Click()
    On Error GoTo errHandler

'    Dim lobjRec As Object                       '单位定位返回的结果记录。
'    Set lobjRec = pobj业务对象.func单位定位     '启动单位定位界面。
'
'    '获取定位的单位，显示在“单位名称”录入框中。(暂时只显示“单位名称”)
'    If Not lobjRec Is Nothing Then
'        If lobjRec.RecordCount > 0 Then
'            ctxt单位名称.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
'        End If
'    End If
    
    FrmCompany.Show
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmImportExcel", "ccmd单位定位_Click", 6666, lstrError, False
End Sub

Private Sub cgrdDetails_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '将当前修改保存到excel中，不提示。然后进行合法性检查
    cgrdDetails.ColDataType(indextmp) = flexDTString
    cgrdDetails.SaveGrid ccdg.FileName, flexFileExcel, 0
    cgrdDetails.Editable = flexEDNone
    ccmdImport.Enabled = sub导入信息合法性检查
End Sub

Private Sub cgrdDetails_DblClick()
    If lbol已经导入 = True Then Exit Sub        '已经导入文件后，不允许单元格编辑
    cgrdDetails.Editable = flexEDKbdMouse
    If cgrdDetails.MouseCol < 0 Or cgrdDetails.MouseRow < 0 Then
        Exit Sub
    ElseIf cgrdDetails.MouseCol >= 0 And cgrdDetails.MouseCol < cgrdDetails.cols Then
        sub录入修改内容 cgrdDetails.MouseRow, cgrdDetails.MouseCol
    End If
End Sub

Sub sub录入修改内容(ByVal paraRow As Integer, ByVal paraCol As Integer)
    cgrdDetails.Select paraRow, paraCol
    cgrdDetails.EditCell
End Sub

Private Sub Form_Load()
        
    ccmdImport.Enabled = False
    lobl已经导入 = False
    
    Set mobj体检 = CreateObject("职业病对象.clsMedicalExam")
    Set mobj体检表模板 = CreateObject("职业病对象.clsMedicalExamTemplate")

    sub填充体检人类型List
    sub填充体检人类别List
End Sub

Private Sub subChangeTemplate()
    On Error GoTo errHandler
    
    If mobj体检.体检表.体检表名 <> ccmbTemplate.Text Then
        mobj体检.体检表.体检表名 = ccmbTemplate.Text
        '根据体检表模板获取该体检表所有可用的字母。
        mobj体检表模板.体检表名 = ccmbTemplate.Text
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmImportExcel", "subChangeTemplate", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'全称为“体检人员类型”，字典内容：职业健康、放射工作
'将体检人类型加入组合框中
Sub sub填充体检人类型List()
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    Set lobjRec = pobjDict.FetchEx("体检人类别字典")
    ccmb体检人类型.Clear
    'ccmb体检人类型.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb体检人类型.AddItem lobjRec("名称")
        ccmb体检人类型.ItemData(ccmb体检人类型.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ccmb体检人类型.ListIndex = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmImportExcel", "sub填充体检人类型List", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'全称为“体检类别”，字典内容：上岗前、在岗期间、离岗时、应急检查
'将体检人类别加入组合框中
Sub sub填充体检人类别List()
    Dim lobjRec As Object
    On Error GoTo errHandler

    Set lobjRec = pobjDict.FetchEx("体检类型字典")
    Ccmb体检人类别.Clear
    'Ccmb体检人类别.AddItem ""
    For i = 1 To lobjRec.RecordCount
        Ccmb体检人类别.AddItem lobjRec("名称")
        Ccmb体检人类别.ItemData(Ccmb体检人类别.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    Ccmb体检人类别.ListIndex = 0
    Ccmb体检人类别.Visible = True
    Call Ccmb体检人类别_Click
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmImportExcel", "sub填充体检人类别List", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'合法性检查内容：1、必要信息是否为空(姓名为空，完全忽略该体检人员信息)；
'2、输入数据类型格式和范围是否正确
'3、出错的某行标记浅黄色，出错的某单元格内字体粗体、蓝色。
Function sub导入信息合法性检查() As Boolean
    Dim i, j As Integer
    Dim changeRowColor As Boolean
    
    On Error GoTo errHandler
    
    sub导入信息合法性检查 = True
    For i = 2 To cgrdDetails.rows - 1
        changeRowColor = False
        For j = 0 To cgrdDetails.cols - 1
            '姓名一栏没有内容，认为改行不导入。也不显示。
            If cgrdDetails.TextMatrix(0, j) = "姓名" And cgrdDetails.TextMatrix(i, j) = "" Then
                '2013-02-26 刘云乐
                '改隐藏为删除行，因为隐藏后在循环保存时有效
'                cgrdDetails.RemoveItem (i)
                cgrdDetails.RowHidden(i) = True
                '2013-02-26 刘云乐
                Exit For
            End If
            
            '判断“年龄”的数据合法性
            If cgrdDetails.TextMatrix(0, j) = "年龄" Then
                If IsNumeric(cgrdDetails.TextMatrix(i, j)) = True Then
                    If CLng(cgrdDetails.TextMatrix(i, j)) > 130 Then
                        sub不合法单元格标记并染色 changeRowColor, i, j
                    End If
                Else
                    sub不合法单元格标记并染色 changeRowColor, i, j
                End If
            End If
            
            '判断“工龄”的数据合法性
'            If cgrdDetails.TextMatrix(0, j) = "工龄" Then
'                'If IsNumeric(cgrdDetails.TextMatrix(i, j)) = False Or CLng(cgrdDetails.TextMatrix(i, j)) > 70 Then sub不合法单元格标记并染色 changeRowColor, i, j
'            End If
            
            '判断“出生日期”的数据合法性
            If cgrdDetails.TextMatrix(0, j) = "出生日期" Then
                If IsDate(cgrdDetails.TextMatrix(i, j)) = False Then sub不合法单元格标记并染色 changeRowColor, i, j
            End If
            
            '判断“婚否”的数据合法性（只能写“已婚”、“未婚”、“离异”）
'            If cgrdDetails.TextMatrix(0, j) = "婚否" Then
'                If Not (cgrdDetails.TextMatrix(i, j) = "已婚" Or cgrdDetails.TextMatrix(i, j) = "未婚" Or cgrdDetails.TextMatrix(i, j) = "离异") Then sub不合法单元格标记并染色 changeRowColor, i, j
'            End If
            
'            '判断“身份证号”的数据合法性
'            If cgrdDetails.TextMatrix(0, j) = "公民身份号码" Then
'                indextmp = j                '记下“公民身份号码”是哪一列，保存时用到
'
'                '格子内容标记不同颜色，行颜色变化
'                Dim lstrSex As String
'                Dim lstrBirth As String
'                sub根据公民身份号码获取生日和性别 cgrdDetails.TextMatrix(i, j), lstrBirth, lstrSex
'                If IsDate(lstrBirth) = False Then sub不合法单元格标记并染色 changeRowColor, i, j
'            End If

            '判断“身份证号”的数据合法性
            '增加其它证件后处理身份证和其它证件   2015-11-16 by 牟俊
            If cgrdDetails.TextMatrix(0, j) = "公民身份号码" Then
                indextmp = j                '记下“公民身份号码”是哪一列，保存时用到

                '格子内容标记不同颜色，行颜色变化
                Dim lstrSex As String
                Dim lstrBirth As String
                sub根据公民身份号码获取生日和性别 cgrdDetails.TextMatrix(i, j), lstrBirth, lstrSex
                If cgrdDetails.TextMatrix(i, j) = "" And cgrdDetails.TextMatrix(i, j + 1) <> "" Then
                cgrdDetails.TextMatrix(i, j) = cgrdDetails.TextMatrix(i, j + 1)
                cgrdDetails.TextMatrix(i, j + 1) = ""
'                lstrBirth1 = cgrdDetails.TextMatrix(i, j + 5)
'                dafuncGetData ("insert into 职业病体检_体检人员基本信息表(出生日期) value('" & cgrdDetails.TextMatrix(i, j + 5) & "')")
                Else
                If IsDate(lstrBirth) = False Then sub不合法单元格标记并染色 changeRowColor, i, j
                End If
            End If

        Next j
        
        If changeRowColor = True Then
            For j = 0 To cgrdDetails.cols - 1
                cgrdDetails.Cell(flexcpBackColor, i, j) = &HC0FFFF      '背景色浅黄
            Next j
            sub导入信息合法性检查 = False
        Else
            '取消单元格染色
            For j = 0 To cgrdDetails.cols - 1
                cgrdDetails.Cell(flexcpForeColor, i, j) = vbBlack       '字体颜色黑色
                cgrdDetails.Cell(flexcpFontBold, i, j) = 0              '字体一般粗细
                cgrdDetails.Cell(flexcpBackColor, i, j) = 0             '无色
            Next j
        End If
    Next i
    
    For i = cgrdDetails.rows - 1 To 2 Step -1
        For j = 0 To cgrdDetails.cols - 1
            '姓名一栏没有内容，认为改行不导入。也不显示。
            If cgrdDetails.TextMatrix(0, j) = "姓名" Then
                 If cgrdDetails.TextMatrix(i, j) = "" Then
                    '2013-02-26 刘云乐
                    '改隐藏为删除行，因为隐藏后在循环保存时有效
                    cgrdDetails.RemoveItem (i)
    '                cgrdDetails.RowHidden(i) = True
                    '2013-02-26 刘云乐
                End If
                Exit For
            End If
        Next j
    Next i
    
    Exit Function
errHandler:
    sub导入信息合法性检查 = False
    MsgBox ("数据格式错误！请仔细检查后，双击修改。")
End Function

Sub sub不合法单元格标记并染色(changeRowColor As Boolean, ByVal index_X As Integer, ByVal index_Y As Integer)
    '格子内容染为蓝色，字体加粗，标记行颜色变化
    changeRowColor = True
    cgrdDetails.Cell(flexcpForeColor, index_X, index_Y) = vbBlue        '字体蓝色
    cgrdDetails.Cell(flexcpFontBold, index_X, index_Y) = 1              '字体粗体
End Sub


Sub sub导入文件内容(ByVal paraCol As Collection)
        On Error GoTo errHandler

        Dim lobjRec As Object
        Dim lstrError As String
        
        Set mcol收费项目 = New Collection
        Set mcol体检项目 = New Collection
        
        mobj体检.系统编号 = paraCol("系统编号")
        Set lobj体检表编号 = CreateObject("职业病对象.clsmedicalexamsheet")
        lobj体检表编号.体检表编号 = ccmbTemplate.Text
        With mobj体检
            If .体检表.体检表名 <> ccmbTemplate.Text Then
                .体检表.体检表名 = paraCol("体检表类型") + "-" + paraCol("体检表类别") + "-" + paraCol("危害因素")
            End If
''            If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
''                If .体检表.试管编号字母 <> FrmRegister.clblLetter.Caption Then
''                    .体检表.试管编号字母 = FrmRegister.clblLetter.Caption
''                End If
''            Else
''                .体检表.试管编号字母 = FrmRegister.clblLetter.Caption
''                .试管编号 = FrmRegister.ctxtTubeNo.Text
''            End If
            On Error Resume Next
            .体检人员.系统编号 = paraCol("系统编号")
            .体检人员.姓名 = paraCol("姓名")
            .体检人员.性别 = paraCol("性别")
            .体检人员.单位名称 = ctxt单位名称
'            If ctxt单位名称.Text = "" Then .体检人员.单位名称 = paraCol("单位名称")
            .体检人员.危害因素 = paraCol("危害因素")
            .体检人员.照射源 = paraCol("照射源")
            .体检人员.职业分类 = paraCol("体检分类")
            .体检人员.现工种 = paraCol("现工种")
            .体检人员.职务或职称 = paraCol("职务或职称")
            .体检人员.职业危害工龄 = paraCol("职业危害工龄")
            .体检人员.放射剂量 = paraCol("放射剂量")
            .体检人员.籍贯 = paraCol("籍贯")
            .体检人员.邮编 = paraCol("邮编")
            .体检人员.住址 = paraCol("住址")
            .体检人员.婚否 = paraCol("婚否")
            .体检人员.电话号码 = paraCol("电话号码")
            .体检人员.工龄 = paraCol("工龄")
            .体检人员.负责人 = paraCol("负责人")
            .体检人员.联系电话 = paraCol("联系电话") '跟上面的“电话号码”重复？？？还是负责人的联系电话
            .体检人员.经济性质 = paraCol("经济性质")
            .体检人员.行业类别 = paraCol("行业类别")
            .体检人员.单位地址 = paraCol("单位地址")
            .体检人员.出生地 = paraCol("出生地")
            .体检人员.年龄 = paraCol("年龄")
             If paraCol("出生日期") <> "" Then .体检人员.出生日期 = Format(paraCol("出生日期"), "yyyy/mm/dd")
'            .体检人员.出生日期 = DateAdd("yy-mm-dd", Val(paraCol("出生日期")), Date)
'            If paraCol("出生日期") = "" Then .体检人员.出生日期 = DateAdd("yyyy", -Val(paraCol("年龄")), Date)
            .体检人员.公民身份号码 = paraCol("公民身份号码")
            .体检人员.卫生种类 = paraCol("卫生种类")
            .体检人员.行业类别 = paraCol("行业类别")
            .体检人员.片区 = paraCol("片区")
            .体检人员.文化程度 = paraCol("文化程度")
            .体检人员.民族 = paraCol("民族")
            .体检人员.职业分类 = paraCol("职业分类")
            .体检人员.体检表类型 = paraCol("体检表类型")
            .体检人员.体检表类别 = paraCol("体检表类别")
'           If paraCol("单位申请编号") = "" Then
'                .体检人员.单位申请编号 = ""
'            Else
                dasubSetQueryTimeout 600
                Set lobjRec = dafuncGetData("select * from 单位档案_单位基本信息表 where 单位名称='" & .体检人员.单位名称 & "'")
                If lobjRec.RecordCount > 0 Then mstr单位申请编号 = lobjRec("申请编号")
                If .体检人员.单位申请编号 <> mstr单位申请编号 Then
                    '给单位编号重新赋值，可以重新获取其卫生种类、行业类别、片区。
                    .体检人员.单位申请编号 = mstr单位申请编号
                End If
'            End If
            
            '保存附加信息
            'For i = 1 To ciptBase.ItemCount
                'If ciptBase.Box1(i - 1).TrueText <> ciptBase.Box1(i - 1).Text And ciptBase.Box1(i - 1).Text <> "" Then
             '   If ciptBase.InfoCollection(i).字典名称 <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
             '       .体检表.Sub填附加信息值 ciptBase.InfoCollection(i).名称, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
             '   Else
             '       .体检表.Sub填附加信息值 ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
             '   End If
           ' Next i
            
            '设置为体检人类别。
            'If ccmb体检人类型.Text = "初检" Then
            '    .体检人类别 = P_EXAM_FIRST
            'Else
            '    .体检人类别 = P_EXAM_ANNUAL
            'End If
            .体检日期 = Now 'Format(cdtpDate.Value, "yyyy-mm-dd hh:mm:ss")
            
            '修改：2004-1-9（增加体检单号）
            '.体检单号 = ctxt体检单号.Text
            '直接取导入的值 2015-7-2 by lanchao
            '.体检人类型 = ccmb体检人类型.Text
            '.体检人类别 = Ccmb体检人类别.Text
            .体检人类型 = paraCol("体检表类型")
            .体检人类别 = paraCol("体检表类别")
            Dim tjbh As String
            tjbh = paraCol("体检表类型") + "-" + paraCol("体检表类别") + "-" + paraCol("危害因素")
            On Error GoTo errHandler
            Set lobj体检表编号 = CreateObject("职业病对象.clsmedicalexamsheet")
            lobj体检表编号.体检表编号 = Trim(tjbh)
            If mcol体检项目.Count = 0 Then
                Set mcol体检项目 = lobj体检表编号.体检表项目集("")
            End If
            Set .col体检项目 = mcol体检项目
            Set lobj体检表编号 = Nothing
        End With
        
        '功能：保存体检项目
        '时间：2012-06-04
        '作者：翁乔
        save优化的体检项目 mcol体检项目, paraCol("系统编号")
        '时间：2012-06-04
        
        If mcol收费项目.Count > 0 Then
            pobj业务对象.Sub体检登记 mobj体检, , , mcol收费项目, Val(1)
        Else
            pobj业务对象.Sub体检登记 mobj体检, , , , Val(1)
        End If
        
        Set lobjRec = CreateObject("职业病界面.clsMoney")
        lobjRec.mstr系统编号 = paraCol("系统编号")
        lobjRec.mstr体检人员姓名 = paraCol("姓名")
        Set lobjRec.col体检项目 = mcol体检项目
        Dim lstr收费批号 As String
    '    lstrError = lobjRec.func收费(lstr收费批号)
        mobj体检.收费批号 = lstr收费批号
        Set lobjRec = Nothing
       ' If lstrError <> "" And lstrError <> "Cancel" Then
      '      MsgBox lstrError, vbOKOnly + vbExclamation, "系统提示"
      '  End If

        
        'mobj体检.Sub保存体检登记信息
        
        '2012-06-25 于登淼 ↓
        '初始化体检基本信息表中“各科体检状态”字段
        subInit各科体检状态 mcol体检项目, Trim(paraCol("系统编号"))
        '2012-06-25 于登淼 ↑
        MousePointer = 0
    Exit Sub
errHandler:
   MousePointer = 0
   sfsub错误处理 "职业病史录入", "frmImportExcel", "sub导入文件内容", Err.Number, Err.Description, False
End Sub

'功能：保存职业病登记里选择的体检项目
'作者：翁乔
'时间：2012-06-04
'说明：首先要查看数据库里面是否有相同的体检项目，然后再进行增加或者修改

Public Sub save优化的体检项目(ByRef para体检项目 As Collection, ByVal para系统编号 As String)
    Dim lstrSql As String
    Dim MedicProjt As String
    Dim rs As Object
    Dim i As Integer
    Dim col体检项目 As Collection
    On Error GoTo errHandler
    
    Set rs = dafuncGetData("select 名称 from 系统管理_字典_字典内容表 where ID = (select ID from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 名称 like '%科'")
    
    For i = 1 To rs.RecordCount
        
        lstrSql = "delete 职业病体检_结果信息_" & rs("名称") & " where 系统编号='" & para系统编号 & "'"
        dafuncGetData lstrSql
        rs.MoveNext
    Next i
    
    Set col体检项目 = para体检项目
    
    For i = 1 To col体检项目.Count
        MedicProjt = Left(Trim(col体检项目(i)("编码")), 2)
        
        lstrSql = "select 名称 from 系统管理_字典_字典内容表 where ID = (select ID from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 编号= '" & MedicProjt & "'"
        Set rs = dafuncGetData(lstrSql)
        
        lstrSql = "insert into 职业病体检_结果信息_" & rs("名称") & "(系统编号,体检项目) values(" _
            & "'" & para系统编号 & "','" & col体检项目(i)("编码") & "')"
        dafuncGetData lstrSql
    Next i
    
    
    Exit Sub
errHandler:
   sfsub错误处理 "职业病对象", "frmImportExcel", "save优化的体检项目", Err.Number, Err.Description, False
End Sub

'2012-07-09 于登淼
'添加管理界面查询函数
Sub sub更新管理界面查询()
    frmRegisterManage.mstr截止日期 = Now
    frmRegisterManage.sub查询并显示
End Sub

'2012-06-25 于登淼
'添加初始化各科体检状态函数。
'用于判断每个体检人员各个体检结果(与结论)科室的体检状态。
'0代表不需要检验的科室；1代表需要检验的科室；2代表该科室已经检验完；
'3代表该科室体检结果与结论不可以再修改。(其中，2、3状态都可以下最终结论)
'状态是一个长度为13的字符串(6-25时有13个填写结果的科室，字符串长度为18)
Sub subInit各科体检状态(paraCol As Collection, paraSysNo As String)
    Dim i As Integer
    Dim paraDeptNo As Integer
    Dim paraState, strSQL As String
    
    
    For i = 1 To 19: paraState = paraState & "0": Next
    paraState = paraState & "1"
    
    For i = 1 To paraCol.Count
        paraDeptNo = CInt(Left(paraCol.Item(i).Item(1), 2))
        paraState = Left(paraState, paraDeptNo - 1) & "1" & Right(paraState, Len(paraState) - (paraDeptNo))
    Next
    
    strSQL = "update 职业病体检_体检基本信息表 set 各科体检状态='" & paraState & "' where 系统编号='" & paraSysNo & "'"
    dafuncGetData strSQL
End Sub

