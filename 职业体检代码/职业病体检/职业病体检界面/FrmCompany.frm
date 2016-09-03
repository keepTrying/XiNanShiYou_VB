VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCompany 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "单位档案"
   ClientHeight    =   7350
   ClientLeft      =   240
   ClientTop       =   375
   ClientWidth     =   10050
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7724.65
   ScaleMode       =   0  'User
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "选择"
      Height          =   375
      Left            =   7920
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox ctxt编号 
      Height          =   270
      Left            =   8160
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox copt保存后清空 
      Caption         =   "保存后清空"
      Height          =   255
      Left            =   6360
      TabIndex        =   35
      Top             =   600
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   32
      Top             =   2760
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088771250
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
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
      ExplorerBar     =   1
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
   Begin VB.Frame Frame3 
      Caption         =   "  现单位信息录入   "
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   9735
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   1335
         Left            =   5880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox ctxt单位名称 
         Height          =   300
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox ctxt负责人 
         Height          =   300
         Left            =   1440
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox ctxt联系电话 
         Height          =   300
         Left            =   4200
         TabIndex        =   24
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox ctxt单位地址 
         Height          =   300
         Left            =   1440
         TabIndex        =   23
         Top             =   1320
         Width           =   4335
      End
      Begin VB.ComboBox ccmb经济性质 
         Height          =   300
         Left            =   4200
         TabIndex        =   22
         Text            =   "ccmb经济性质"
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Ccmb行业类别 
         Height          =   300
         Left            =   1440
         TabIndex        =   21
         Text            =   "Ccmb行业类别"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   5
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "负责人："
         Height          =   180
         Left            =   600
         TabIndex        =   30
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "联系电话："
         Height          =   180
         Left            =   3240
         TabIndex        =   29
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "单位或地址："
         Height          =   180
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "经济性质："
         Height          =   180
         Left            =   3240
         TabIndex        =   27
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "行业类别："
         Height          =   180
         Left            =   480
         TabIndex        =   26
         Top             =   1080
         Width           =   900
      End
   End
   Begin VB.Frame cfram基本信息 
      Caption         =   "登记基本信息（非快速录入时黄色为必录项，快速录入时只需照相):"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   9360
      Width           =   6300
      Begin VB.TextBox ctxt民族 
         Height          =   300
         Left            =   4800
         TabIndex        =   18
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox ccmb体检时期 
         Height          =   300
         Left            =   8160
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox ctxt份数 
         Height          =   270
         Left            =   4440
         TabIndex        =   14
         Text            =   "1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox ctxt体检单号 
         Height          =   315
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox ctxtTubeNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6000
         TabIndex        =   0
         Top             =   1320
         Width           =   1575
      End
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6360
         TabIndex        =   3
         Top             =   1320
         Width           =   345
      End
      Begin 录入控件.ctlInputDictGrid c字典表 
         Height          =   3255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
         Cols            =   10
         Count           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 录入控件.ctlInputFrame ciptBase 
         Height          =   975
         Left            =   6120
         TabIndex        =   2
         Top             =   2280
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1720
         BackColor       =   15791081
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   ""
         Rows            =   1
         Cols            =   27
         DistanceofRow   =   0
         AutoSize        =   0   'False
         FormatString    =   "身份证号,1,0,12"
         Count           =   1
         titleInputBox0001=   "身份证号"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   12
         orderInputBox0001=   1
         valueInputBox0001=   ""
         datatypeInputBox0001=   3
         colInputBox0001 =   0
         rowInputBox0001 =   1
         PassWordCharInputBox0001=   0   'False
         主键InputBox0001=   0   'False
         允许等于最大值InputBox0001=   0   'False
         允许等于最小值InputBox0001=   0   'False
         字典名称InputBox0001=   ""
         显示字典字段InputBox0001=   ""
         保存字典字段InputBox0001=   ""
         名称InputBox0001=   "输入框 1"
         缺省值InputBox0001=   ""
         保存缺省值InputBox0001=   ""
         长度InputBox0001=   0
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         允许多选InputBox0001=   0   'False
         ErrColor        =   12648447
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "请将二代身份证放在读卡器上！"
         Height          =   180
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2520
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "民族："
         Height          =   180
         Left            =   4800
         TabIndex        =   17
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "体检时期："
         Height          =   180
         Left            =   8160
         TabIndex        =   15
         Top             =   480
         Width           =   900
      End
      Begin VB.Label clbl份数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "份数："
         Height          =   180
         Left            =   3600
         TabIndex        =   13
         Top             =   2880
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检单号："
         Height          =   180
         Index           =   7
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label clbl旧体检日期 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8760
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上次体检日期："
         Height          =   180
         Index           =   4
         Left            =   8640
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "保存后请看状态栏"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6600
         TabIndex        =   8
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6000
         TabIndex        =   7
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管编号："
         Height          =   180
         Index           =   1
         Left            =   6000
         TabIndex        =   6
         Top             =   1080
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
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
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：建立单位档案（赶时间，功能小，所以没有分开。全部写在一起了。）
'作者：张令
'日期：2013.02.28

Option Explicit
Private mblnInUse As Boolean
Private mcolIndex As Collection
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

'单击表格，激活修改按钮。
Private Sub cgrdMain_Click()
    ctbMain.Buttons(5).Enabled = True
    
End Sub

'双击表格，填充导入界面的“单位名称”
Private Sub cgrdMain_DblClick()
    Dim lstrSysNo As String
    Dim lobjRec As Object
    lstrSysNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("申请编号"))
    Set lobjRec = dafuncGetData("select 单位名称 from 单位档案_单位基本信息表 where 申请编号='" & lstrSysNo & "'")
    frmImportExcel.ctxt单位名称 = lobjRec(0)
    Unload Me
End Sub

Private Sub Command1_Click()
    frmAddRegister.ccmbUnit.Text = cgrdMain.TextMatrix(cgrdMain.Row, 1)
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
End Sub
'初始化界面
Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lobjDetl As Object
    Dim i As Integer
    On Error GoTo errHandler
   
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    MousePointer = 0
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    mobjGUI.pbln自动设置字典高度 = False
    
    '设置工具栏上所需要的各种按钮。
    Dim lcol工具栏按钮 As New Collection           '工具栏上的按钮初始化集合。
    With lcol工具栏按钮
        .Add "查询(&C)108"
        .Add "|"
        .Add "保存(&T)101"
        .Add "|"
        .Add "修改"
        .Add "|"
        .Add "删除"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
        Set .c录入板 = ciptBase
        Set .c字典表 = c字典表
        Set .c状态栏 = ctbMain
        
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcol工具栏按钮, ""
    End With
    Text1.Text = "操作说明：1、添加：在“现有单位信息录入" _
            & "”中录入信息以后，点击保存即可。2、修改" _
            & "：点击表格中一条需要修改的数据，点击修改" _
            & "，然后修改其数据，最后保存即可。3、删除" _
            & ": 选中需要删除的数据点击删除即可?4?查" _
            & "询：在现单位信息录入中录入数据，点击查询" _
            & "即可（为空表示不做条件查询）。"
    
    SubSelect
    subClear
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 " 单位档案管理", "FrmCompany", "Form_Load", 6666, lstrError, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    
    Set mobjGUI = Nothing
End Sub

'功能：处理工具栏上按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lstrError As String
    Dim lstrSysNo As String
    
    SubSelect
'    lstrSysNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("申请编号"))
'    lstrSysNo = ""
    Select Case Operate
    Case "查询"
        subQuery
    Case "保存"
        If ctxt单位名称 = "" Then
            MsgBox "单位名称不能为空！"
            Exit Sub
        End If
        
        '生成申请编号。
        subSave (ctxt编号)
        SubSelect

        lstrSysNo = ""
        subClear
        If ctbMain.Buttons(5).Enabled = False Then
            ctbMain.Buttons(5).Enabled = True
        End If
    Case "修改"
        ctxt编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("申请编号"))
        lstrSysNo = ctxt编号
        subUpdata lstrSysNo
        SubSelect
'        ctxt编号 = ""
        ctbMain.Buttons(5).Enabled = False
    Case "删除"
        ctxt编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("申请编号"))
        lstrSysNo = ctxt编号
        If cgrdMain.Row = 0 Or cgrdMain.Row > cgrdMain.rows - 1 Then
            MsgBox "请选择需要删除的数据！"
        Else
            If MsgBox("确定要删除该单位信息？", vbYesNo, "系统提示") = vbYes Then
                subDelete lstrSysNo
            End If
        End If
        SubSelect
        ctxt编号 = ""
        lstrSysNo = ""
    Case "退出"
        Unload Me
    End Select
    Exit Sub
errHandler:
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "单位档案管理", "FrmCompany", "mobjGUI_BeforeOperate", 6666, lstrError, False
End Sub

 
'保存
Private Sub subSave(ByVal paraNo As String)
    Dim lobjRec As Object
    Dim lstrSql As String
    Dim mstrSQNo As String
    Set lobjRec = dafuncGetData("select * from 单位档案_单位基本信息表 where 申请编号='" & paraNo & "'")
    If (lobjRec.BOF Or lobjRec.EOF) Then
    
        mstrSQNo = func取新的申请编号(pstr工作站代号)
        lstrSql = " insert into 单位档案_单位基本信息表(申请编号,单位名称,负责人,电话,经济性质,行业类别,地址) values('" & mstrSQNo & "','" & Trim(ctxt单位名称) & "" _
                            & "','" & Trim(ctxt负责人) & "','" & Trim(ctxt联系电话) & "','" & Trim(ccmb经济性质) & "','" & Trim(Ccmb行业类别) & "','" & Trim(ctxt单位地址) & "')"
        dafuncGetData lstrSql
    Else
        lstrSql = "update 单位档案_单位基本信息表 set 单位名称='" & Trim(ctxt单位名称) & "',负责人='" & Trim(ctxt负责人) & "',电话='" & Trim(ctxt联系电话) & "" _
                                    & "',经济性质='" & Trim(ccmb经济性质) & "',行业类别='" & Trim(Ccmb行业类别) & "',地址='" & Trim(ctxt单位地址) & "" _
                                    & "' where 申请编号='" & paraNo & "'"
        dafuncGetData lstrSql
    End If
    If copt保存后清空.Value = True Then
        subClear
    End If
End Sub

'查询并显示
Private Sub SubSelect()
    Dim lobjRec As Object
    Dim i As Integer
    Set lobjRec = dafuncGetData("select 申请编号,单位名称,负责人,电话,经济性质,行业类别,地址 from 单位档案_单位基本信息表")
    Set cgrdMain.DataSource = lobjRec
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    cgrdMain.AutoSize 0, cgrdMain.cols - 1, 0, 0
    cgrdMain.ColHidden(mcolIndex("申请编号")) = True
End Sub

'获取申请编号
Public Function func取新的申请编号(ByVal para用户编号 As String) As String
    On Error GoTo errHandler
    Dim mtempSql As String
    Dim lrstTemp As Object
    '注意：因数据访问对象未提供对存储过程使用参数返回结果的方法，在此必须修改源存储过程，改用记录集返回申请编号和档案编号
    mtempSql = "exec 单位档案_生成申请编号 "
    Set lrstTemp = dafuncGetData(mtempSql)
    'lrstTemp.Open
    If lrstTemp.RecordCount <> 0 Then
        func取新的申请编号 = lrstTemp.Fields("申请编号")
    Else
        func取新的申请编号 = ""
    End If
    Exit Function
errHandler:
    sfsub错误处理 "单位档案业务对象", "ClsManageUnitFile", "func取新的申请编号", Err.Number, Err.Description, True
End Function

'修改
Private Sub subUpdata(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lstrSql As String
    Set lobjRec = dafuncGetData("select * from 单位档案_单位基本信息表 where 申请编号='" & paraSysNo & "'")
    If lobjRec.RecordCount = 0 Then Exit Sub
    ctxt单位名称.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
    ctxt负责人.Text = IIf(IsNull(lobjRec("负责人")), "", lobjRec("负责人"))
    ctxt联系电话.Text = IIf(IsNull(lobjRec("电话")), "", lobjRec("电话"))
    ccmb经济性质.Text = IIf(IsNull(lobjRec("经济性质")), "", lobjRec("经济性质"))
    Ccmb行业类别.Text = IIf(IsNull(lobjRec("行业类别")), "", lobjRec("行业类别"))
    ctxt单位地址.Text = IIf(IsNull(lobjRec("地址")), "", lobjRec("地址"))
End Sub

'删除
Private Sub subDelete(ByVal paraSysNo As String)
    dafuncGetData ("delete 单位档案_单位基本信息表 where 申请编号='" & paraSysNo & "'")
End Sub

'保存后清空界面
Private Sub subClear()
    Dim lobjRec As Object
    Dim lobjDetl As Object
    Dim i As Integer
    ctxt单位名称 = ""
    ctxt负责人 = ""
    ctxt联系电话 = ""
    ctxt单位地址 = ""
    ctxt编号 = ""
    '获取经济性质
    Set lobjRec = pobjDict.FetchEx("经济性质字典")
    ccmb经济性质.Clear
    ccmb经济性质.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb经济性质.AddItem lobjRec("名称")
        ccmb经济性质.ItemData(ccmb经济性质.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ccmb经济性质.ListIndex = 0
    Set lobjRec = CreateObject("职业病对象.clsmedicalexamtemplateset")
    Set lobjDetl = lobjRec.行业类别
    Ccmb行业类别.Clear
    Ccmb行业类别.AddItem ""
    For i = 1 To lobjDetl.RecordCount
        Ccmb行业类别.AddItem lobjDetl("名称")
        Ccmb行业类别.ItemData(Ccmb行业类别.NewIndex) = lobjDetl("编号")
        lobjDetl.MoveNext
    Next
    Ccmb行业类别.ListIndex = 0
End Sub

'按条件查询
Private Sub subQuery()
    Dim lstrSql As String
    Dim lobjRec As Object
    Dim i As Integer
    lstrSql = "select 申请编号,单位名称,负责人,电话,经济性质,行业类别,地址 from 单位档案_单位基本信息表 where 1=1"
    If ctxt单位名称 <> "" Then
        lstrSql = lstrSql & " and 单位名称='" & Trim(ctxt单位名称) & "'"
    End If
    If ctxt负责人 <> "" Then
        lstrSql = lstrSql & " and 负责人='" & Trim(ctxt负责人) & "'"
    End If
    If ctxt联系电话 <> "" Then
        lstrSql = lstrSql & " and 电话='" & Trim(ctxt联系电话) & "'"
    End If
    If ccmb经济性质 <> "" Then
        lstrSql = lstrSql & " and 经济性质='" & Trim(ccmb经济性质) & "'"
    End If
    If Ccmb行业类别 <> "" Then
        lstrSql = lstrSql & " and 行业类别='" & Trim(Ccmb行业类别) & "'"
    End If
    If ctxt单位地址 <> "" Then
        lstrSql = lstrSql & " and 地址='" & Trim(ctxt单位地址) & "'"
    End If
    
    Set lobjRec = dafuncGetData(lstrSql)
    
    Set cgrdMain.DataSource = lobjRec
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    cgrdMain.AutoSize 0, cgrdMain.cols - 1, 0, 0
    cgrdMain.ColHidden(mcolIndex("申请编号")) = True
End Sub
