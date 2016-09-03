VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportManage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "报告管理"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10725
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton coptType 
      Caption         =   "所有体检完成人员"
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "8023部队已打印"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "8023部队未打印"
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox cchkAll 
      Caption         =   "全选 "
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.OptionButton coptType 
      Caption         =   "未打印"
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton coptType 
      Caption         =   "已打印"
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   953
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   10095
      _cx             =   17806
      _cy             =   10610
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认显示近一个月的数据；执行任何操作，请勾选列表的对应行。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   5220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总记录数："
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   300
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   330
   End
End
Attribute VB_Name = "frmReportManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'名称：职业病体检界面(报告管理)
'函数：
'功能：职业病体检界面(报告管理)管理打印未打印报告即查询预览
'作者：罗李奎
'时间：203.01
'***************************************

Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1


Private mstr开始日期 As String
Private mstr截止日期 As String
Private mstr体检表名称 As String
Private mstr体检类别 As String
Private mstr系统编号 As String
Private mstr报告编号 As String
Private mstr姓名 As String
Private mstr单位名称 As String

Private mstrState As String '报告打印状态
'查询结果
Private mobjQueryResult As Object

Private mcolIndex As New Collection

'功能：返回当前窗体是否已经加载标志。这是系统平台所要求的。
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property


'Private Sub cgrdMain_Click()
'     With cgrdMain
'        If .Row < 1 Or .Row > .rows - 1 Then Exit Sub
'        mstr系统编号 = .TextMatrix(.Row, mcolIndex("系统编号"))
'        mstr体检表名称 = .TextMatrix(.Row, mcolIndex("体检表名称"))
'        mstr体检类别 = .TextMatrix(.Row, mcolIndex("体检类别"))
'        mstr体检类别 = .TextMatrix(.Row, mcolIndex("体检类别"))
'        mstrState = .TextMatrix(.Row, mcolIndex("报告状态"))
'        mstr报告编号 = .TextMatrix(.Row, mcolIndex("报告编号"))
'    End With
'End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '查询
        mstrQuery = 1
        With frmQuery
            '显示旧的查询条件。
            .pstr开始日期 = mstr开始日期
            .pstr截止日期 = mstr截止日期
            .pstr体检表名称 = mstr体检表名称
            .pstr姓名 = mstr姓名
            .pstr单位名称 = mstr单位名称
            .pstr系统编号 = mstr系统编号
            '获取新的查询条件。
            .Show 1, Me
            If .pblnOk Then
                mstr开始日期 = .pstr开始日期
                mstr截止日期 = .pstr截止日期
                mstr体检表名称 = .pstr体检表名称
                mstr系统编号 = .pstr系统编号
                mstr单位名称 = .pstr单位名称
                mstr姓名 = .pstr姓名
               
                '重新查询。
                sub查询并显示
            End If
        End With
    
    Case 2 '刷新
        sub查询并显示
    Case 4
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "报告管理", "frmReportmanage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub

Private Sub cchkAll_Click()
    Dim i As Integer
    If cchkAll.Value = 1 Then
        For i = 1 To cgrdMain.rows - 1
           cgrdMain.Cell(flexcpChecked, i, 0) = flexChecked
        Next i
    Else
        For i = 1 To cgrdMain.rows - 1
           cgrdMain.Cell(flexcpChecked, i, 0) = flexUnchecked
        Next i
    End If
End Sub



'Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Col <> 0 Then
'        Cancel = True
'    Else
''        cgrdMain.CellChecked = flexChecked
'    End If
'End Sub


Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    subClear
    sub查询并显示
    Exit Sub
errHandler:
    sfsub错误处理 "报告管理", "frmReportmanage", "coptType_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    '显示进度。
    frmProcess.proPercent.Max = 4
    frmProcess.Label1.Caption = "正在初始化界面，请等待..."
    frmProcess.proPercent.Value = 1
    frmProcess.Show
    DoEvents
    Me.Enabled = False
    MousePointer = 11
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    With lcol工具栏按钮
        .Add "查询(&Q)108"
        .Add "|"
        .Add "打印(&P)107"
        .Add "|"
        .Add "预览(&V)102"
        .Add "|"
        .Add "导出(&O)110"
        .Add "|"
        .Add "单位体检报告(&M)107"
        .Add "|"
        .Add "退出"
    End With
    frmProcess.proPercent.Value = 2
    DoEvents
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlb工具栏
    End With
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""

'            With cgrdMain
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检条码号"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "姓名"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "性别"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "年龄"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "单位名称"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检类型"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检日期"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检日期"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检日期"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检日期"
'        End With
    '默认为列出近一月记录
    mstr截止日期 = Format(Now(), "yyyy-mm-dd")
    mstr开始日期 = Format(DateAdd("m", -1, Now()), "yyyy-mm-dd")
    frmProcess.proPercent.Value = 3
    DoEvents
    sub查询并显示
    frmProcess.proPercent.Value = 4
    Unload frmProcess
    
'    Exit Sub
errHandler:
    Me.Enabled = True
    MousePointer = 0
    If Err.Number <> 0 Then
        sfsub错误处理 "报告管理", "frmReportManage", "form_load", Err.Number, Err.Description, False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

Public Sub sub查询并显示()
    Dim lobjRec As Object
    Dim strSQL As String
    On Error GoTo errHandler
    If mstr截止日期 <> "" And Len(mstr截止日期) < 13 Then
        mstr截止日期 = mstr截止日期 + " 23:59:59"
    End If
    strSQL = "exec 职业病体检_查询体检报告信息 '" & mstr开始日期 & "','" & mstr截止日期 & "','" & mstr体检表名称 & "','" & mstr系统编号 & "','" & mstr单位名称 & "','" & mstr姓名 & "' "
    dasubSetQueryTimeout 6000
    Set mobjQueryResult = dafuncGetData(strSQL)
    
'    '增加8023判断打印或不打印的人存储过程  2016-1-8 by 牟俊
'     Set lobjRec = dafuncGetData("exec 职业病体检_查询体检报告信息8023部队 '" & mstr开始日期 & "','" & mstr截止日期 & "','" & mstr体检表名称 & "','" & mstr系统编号 & "','" & mstr单位名称 & "','" & mstr姓名 & "'")
    
    
    If coptType(0).Value Then
       mobjQueryResult.Filter = "报告状态='未打印'"
    Else
       mobjQueryResult.Filter = "报告状态='已打印'"
    End If
'    '屏蔽8023部队和所有人员信息的判断 2016-1-8 by 牟俊 ↓
'    If coptType(0).Value Then
'       mobjQueryResult.Filter = "报告状态='未打印'"
'    ElseIf coptType(1).Value Then
'       mobjQueryResult.Filter = "报告状态='已打印'"
''    Set lobjRec = dafuncGetData("exec 职业病体检_查询体检报告信息8023部队 '" & mstr开始日期 & "','" & mstr截止日期 & "','" & mstr体检表名称 & "','" & mstr系统编号 & "','" & mstr单位名称 & "','" & mstr姓名 & "'")
'    ElseIf coptType(2).Value Then
'     lobjRec.Filter = "报告状态='未打印'"
'    ElseIf coptType(3).Value Then
'     lobjRec.Filter = "报告状态='已打印'"
'     Else
'        Dim lobsy As Object
'        Set lobsy = dafuncGetData("exec 职业病体检_查询体检报告信息所有人员 '" & mstr开始日期 & "','" & mstr截止日期 & "','" & mstr体检表名称 & "','" & mstr系统编号 & "','" & mstr单位名称 & "','" & mstr姓名 & "'")
''        Set lobjRec = dafuncGetData("exec 职业病体检_查询体检报告信息8023部队 '" & mstr开始日期 & "','" & mstr截止日期 & "','" & mstr体检表名称 & "','" & mstr系统编号 & "','" & mstr单位名称 & "','" & mstr姓名 & "'")
'        lobsy.Filter = "报告状态='所有体检完成人员'"
'    End If
'     '屏蔽8023部队和所有人员信息的判断 2016-1-8 by 牟俊 ↑
    With cgrdMain
        .rows = 1
If coptType(0).Value = True Or coptType(1).Value = True Then
        
        If Not (mobjQueryResult.EOF Or mobjQueryResult.BOF) Then
            Set .DataSource = mobjQueryResult
            
'            .Sort = flexSortGenericDescending
            'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
            .AutoSize 0, .cols - 1, 0, 0
'            .ExplorerBar = flexExSort
'            .DataMode = flexDMFree
            Dim i As Long
            Set mcolIndex = New Collection
            For i = 0 To .cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            .ColHidden(mcolIndex("报告编号")) = True
            For i = 1 To .rows - 1
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            Next i
'            .Editable = True
        End If
        clblInfo.Caption = .rows - 1
        
        
'        '8023和其它人员信息查询显示屏蔽  2016-1-8 by 牟俊 ↓
'ElseIf coptType(2).Value = True Or coptType(3).Value = True Then
'        If Not (lobjRec.EOF Or lobjRec.BOF) Then
'            Set .DataSource = lobjRec
'
''            .Sort = flexSortGenericDescending
'            'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
'            .AutoSize 0, .cols - 1, 0, 0
''            .ExplorerBar = flexExSort
''            .DataMode = flexDMFree
''            Dim i As Long
'            Set mcolIndex = New Collection
'            For i = 0 To .cols - 1
'                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'            Next
'            .ColHidden(mcolIndex("报告编号")) = True
'            For i = 1 To .rows - 1
'                .Cell(flexcpChecked, i, 0) = flexUnchecked
'            Next i
''            .Editable = True
'        End If
'        clblInfo.Caption = .rows - 1
'Else
'        If Not (lobsy.EOF Or lobsy.BOF) Then
'            Set .DataSource = lobsy
'
''            .Sort = flexSortGenericDescending
'            'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
'            .AutoSize 0, .cols - 1, 0, 0
''            .ExplorerBar = flexExSort
''            .DataMode = flexDMFree
''            Dim i As Long
'            Set mcolIndex = New Collection
'            For i = 0 To .cols - 1
'                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'            Next
'            .ColHidden(mcolIndex("报告编号")) = True
'            For i = 1 To .rows - 1
'                .Cell(flexcpChecked, i, 0) = flexUnchecked
'            Next i
''            .Editable = True
'        End If
'        clblInfo.Caption = .rows - 1
'    '8023和其它人员信息查询显示屏蔽  2016-1-8 by 牟俊 ↑
    
End If
    End With
    Exit Sub
errHandler:
    sfsub错误处理 " 报告管理", "frmReportmanage", "sub查询并显示", Err.Number, Err.Description, True
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 80
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 80
    ctlb工具栏.Width = Me.ScaleWidth
    
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Select Case Operate

    Case "查询"
        cmnuItemView_Click 1

    Case "打印"
        subPrint False  '是否预览
        Cancel = True
        
    Case "预览"
     
        subPrint True  '是否预览
        Cancel = True
'         Dim msg As String
'         Cancel = True
'         '         mstr报告管理 = "报告管理" '用来加载信息时判断是哪个窗体发出的请求
'             If mstr报告编号 = "" Then
'                    '判断word信息是否保存了
'                    If coptType(0).Value = True Then
'                        '加载word信息进度条
'                            frmProcess.proPercent.Max = 4
'                            frmProcess.Label1.Caption = "正在加载，请等待..."
'                            frmProcess.proPercent.Value = 0
'                            frmProcess.Show 0, Me
'                            DoEvents
'                             sub编辑word文档 Me, mstr系统编号, mstr体检表名称, False
'                    Else
'                        msg = MsgBox("该信息没有保存word,需要加载信息，现在是否加载信息？", vbYesNo + vbDefaultButton1 + vbQuestion, "系统提示！")
'                            If msg = vbYes Then
'                            '加载word信息进度条
'                            frmProcess.proPercent.Max = 4
'                            frmProcess.Label1.Caption = "正在加载，请等待..."
'                            frmProcess.proPercent.Value = 0
'                            frmProcess.Show 0, Me
'                            DoEvents
'
'                            sub编辑word文档 Me, mstr系统编号, mstr体检表名称, False
'
'                            Else
'                                 Exit Sub
'                            End If
'                    End If
'
'                     Unload frmProcess
'                      '控制体检状态（word的处理并不好，需要把数据库功能添加之后，这步骤才完善）
'                    If pstrFilename = "" Then Exit Sub
'                    If coptType(0).Value = True Then pobj业务对象.func写入单人当前体检状态 mstr系统编号, 7   '"已发报告"
'
'             Else
'                '读取保存的word
'               sub读取word文档 Me, mstr系统编号, mstr报告编号, False
'            End If
'    Case "退回"
'        If cgrdMain.SelectedRows = 0 Or cgrdMain.Row > cgrdMain.rows - 1 Then
'            MsgBox "请选择数据！"
'        Else
'            If MsgBox("确定要退回总检科？", vbYesNo, "系统提示") = vbYes Then
'               dafuncGetData "update 职业病体检_体检基本信息表 set 体检状态='5' where 系统编号='" & Trim(cgrdMain.TextMatrix(cgrdMain.Row, 0)) & "'"
'            cgrdMain.RowHidden(cgrdMain.Row) = True
'        End If
'        End If
    Case "导出"
        If cgrdMain.rows <= 1 Then
            MsgBox "没有导出的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        Dim lstrFile As String
        ccmdFile.Filter = "Excel文件 (*.xls)|*.xls|文本文件 (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            '2012-04-14 于登淼 ↓
            '认为第0列，为系统编号。设置其列保存时为string
            cgrdMain.ColDataType(0) = flexDTString
            cgrdMain.SaveGrid lstrFile, flexFileExcel, True   '导出excel系统编号为数字
            'cgrdMain.SaveGrid lstrFile, flexFileTabText, True
            '2012-04-14 于登淼↑
        End If
    Case "退出"
         Cancel = True
         subClear
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "报告管理", "frmReportManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub
Private Sub subClear()
        mstr体检表名称 = ""
        mstr系统编号 = ""
        mstr姓名 = ""
        mstr单位名称 = ""
End Sub
'供WORD宏调用，用于保存WORD至数据库
Public Sub subSave(ByVal paraFile As String, ByVal paraNo As Integer, ByVal para监测编号 As String)
    subSaveDoc paraFile, paraNo, para监测编号
End Sub

'打印体检表
Private Sub subPrint(ByVal para预览 As Boolean)


    Dim i As Integer
    Dim lobj文书 As Object
    Dim lcolSysNo As Collection
    On Error GoTo errHandler
  
    Set lobj文书 = CreateObject("职业病文书.cls文书")
'    sum = 0
    With cgrdMain
        For i = 1 To .rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                Set lcolSysNo = New Collection
                lcolSysNo.Add .TextMatrix(i, 0)

                 
                 lobj文书.Sub打印文书 "职业健康体检_" & .TextMatrix(i, mcolIndex("体检类型")), lcolSysNo, para预览

                If para预览 = False Then
                    dafuncGetData "update 职业病体检_体检基本信息表 set 体检状态='7' where 系统编号='" & Trim(.TextMatrix(i, 0)) & "'"
                    .RowHidden(i) = True
                End If
            End If
        Next i
        If lcolSysNo.Count < 1 And .rows > 1 Then
            MsgBox "请勾选要打印或预览的体检表！", vbInformation, "系统提示"
            Exit Sub
        End If
    End With
errHandler:
   
End Sub


