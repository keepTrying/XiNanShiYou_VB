VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCareerHstMage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "受检者个人信息管理"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton coptType 
      Caption         =   "未录入受检者个人信息"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton coptType 
      Caption         =   "体检中"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton coptType 
      Caption         =   "未下结论"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "待复核"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "待复查"
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   6255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   10095
      _cx             =   2088781198
      _cy             =   2088774425
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
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12705
      _ExtentX        =   22410
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
      Left            =   720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总记录数："
      Height          =   180
      Left            =   8520
      TabIndex        =   8
      Top             =   840
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   9360
      TabIndex        =   7
      Top             =   840
      Width           =   90
   End
End
Attribute VB_Name = "frmCareerHstMage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'名称：职业病史(受检者个人信息)管理界面
'函数：
'功能：职业病史(受检者个人信息)管理界面上工具栏控制
'作者：Yunle Liu
'时间：2012.03
'***************************************

Option Explicit
Public mblninuse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

'查询条件
Private mstr开始日期 As String
Private mstr截止日期 As String
Private mstr体检表名称 As String
Private mstr单位名称 As String
Private mstr姓名 As String
Private mstr体检单号 As String
Private mstr试管编号 As String
Private mstr系统编号 As String
Private mstr身份证号 As String
'查询结果
Private mobjQueryResult As Object

Private mcolIndex As New Collection

'功能：返回当前窗体是否已经加载标志。这是系统平台所要求的。
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblninuse
End Property

Private Sub cmnuItemPrint_Click(Index As Integer)
    Dim lcol编号 As Collection
    On Error GoTo errHandler
    Set lcol编号 = New Collection
    Select Case Index
    Case 1
        '打印体检表
        lcol编号.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
        pobj业务对象.Sub打印文书 "体检表", lcol编号, True
        
    Case 2
        '打印体检结果单
        lcol编号.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
        pobj业务对象.Sub打印文书 "体检结果单", lcol编号, True
    End Select
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstManage", "cmnuItemPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemRegister_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '初检登记
        '访问标志 = 0 '用于控制焦点
        frmCareerHstRegt.ctxtsysno = ""
        
        'frm刷条码.Show 1, Me
        frmCareerHstRegt.Show 1, Me
        
        '重新查询。
        sub查询并显示
    Case 2 '年检登记
        If cgrdMain.Row >= 1 Then
            'FrmRegister.pstr系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        Else
            'FrmRegister.pstr系统编号 = ""
        End If
        'FrmRegister.Show 1, Me
        
        '重新查询。
        sub查询并显示
    
    Case 3 '复查登记
        If cgrdMain.Row < 1 Then
            MsgBox "没有需要复查的人！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        'FrmRegisterAgain.pstr旧系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        'FrmRegisterAgain.Show 1, Me
        
    Case 5 '修改
        If cgrdMain.Row < 1 Then
            MsgBox "没有需要修改的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        访问记号 = 1
        pub系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        'frmCareerHstRegt.ctxtsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        frmCareerHstRegt.Show 1, Me
        '重新查询。
        sub查询并显示
    
    Case 6 '删除
        If cgrdMain.Row < 1 Then
            MsgBox "没有可以删除的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        If Not coptType(3) Then
            If MsgBox("你确认要删除该体检记录吗？一旦删除后将不能恢复！", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
                pobj业务对象.sub删除体检登记 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
                    
                cgrdMain.RemoveItem cgrdMain.Row
                clblInfo = cgrdMain.Rows - 1
            End If
        Else
            MsgBox "已下体检结论的记录不允许删除！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
    '修改人：罗李奎 2012-12-3  ↓
    '说明：退出时清空条件查询
    'bug 号 0000078
     Case 7 '退出
        mstr体检表名称 = ""
        mstr单位名称 = ""
        mstr姓名 = ""
        mstr体检单号 = ""
        mstr试管编号 = ""
        mstr系统编号 = ""
        mstr身份证号 = ""
        
        Unload Me
      ' 罗李奎  2012-12-3   ↑
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstmanage", "cmnuItemRegister_Click", Err.Number, Err.Description, False
End Sub


Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '查询
        With frmQuery
            '显示旧的查询条件。
            .pstr开始日期 = mstr开始日期
            .pstr截止日期 = mstr截止日期
            .pstr体检表名称 = mstr体检表名称
            .pstr姓名 = mstr姓名
            .pstr单位名称 = mstr单位名称
            .pstr体检单号 = mstr体检单号
            .pstr试管编号 = mstr试管编号
            .pstr身份证号 = mstr身份证号
            
            '获取新的查询条件。
            .Show 1, Me
            If .pblnOk Then
                mstr开始日期 = .pstr开始日期
                mstr截止日期 = .pstr截止日期
                mstr体检表名称 = .pstr体检表名称
                mstr单位名称 = .pstr单位名称
                mstr姓名 = .pstr姓名
                mstr体检单号 = .pstr体检单号
                mstr试管编号 = .pstr试管编号
                mstr系统编号 = .pstr系统编号
                mstr身份证号 = .pstr身份证号
                '重新查询。
                sub查询并显示
            End If
        End With
    
    Case 2 '刷新
        sub显示查询结果
    Case 4
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstmage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub







Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    sub显示查询结果
    
    ctlb工具栏.Buttons(4).Enabled = coptType(1).Value
    'cmnuItemRegister(3).Enabled = coptType(1).Value
    ctlb工具栏.Buttons(5).Enabled = coptType(1).Value
    ctlb工具栏.Buttons(6).Enabled = coptType(0).Value Xor coptType(1).Value
    '修改人：罗李奎 2012-12-3 ↓
    '说明：判断在体检中不显示删除按钮
    'bug号：0000079
    If coptType(1) = True Then
        ctlb工具栏.Buttons(6).Enabled = False
     End If
    '修改人：罗李奎 2012-12-3  ↑
    ctlb工具栏.Buttons(7).Enabled = coptType(0).Value Xor coptType(1).Value
    
    'cmnuItemRegister(5).Enabled = coptType(0).Value
    'cmnuItemRegister(6).Enabled = coptType(0).Value
    '修改人：张令 2012.12.05
    'bug号：0000073
    '说明：当已下结论和待复查选中时，受检者个人信息录入按钮不可用！    ↓↓
    If coptType(3).Value = True Or coptType(4).Value = True Then
        ctlb工具栏.Buttons(3).Enabled = False
    Else
        ctlb工具栏.Buttons(3).Enabled = True
    End If
    '2012.12.05     ↑↑
    '2012-04-14 于登淼 ↓
    '菜单里删除“打印”大项，包括(1)“体检表”，(2)“体检结果单”
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 于登淼 ↑
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstmanage", "coptType_Click", Err.Number, Err.Description, False
End Sub





Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblninuse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblninuse = True
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    '修改：2002-7-1（杨春）简化取消结论的操作。该为操作单选框。
    With lcol工具栏按钮
        .Add "查询(&Q)108"
        .Add "|"
        .Add "受检者信息录入(&R)102"
        '.Add "复查登记(&R)103"
        .Add "|"
        .Add "修改"
        .Add "删除"
        .Add "|"
        .Add "导出(&O)113"
        '.Add "|"
        '.Add "打印(&P)107"
        '2012-02-13 于登淼 ↓
        '修改内容：添加“单位导入”按钮，快捷键Alt+I，工具栏图片编号110
        '          添加“打印条码”按钮，快捷键Alt+P，工具栏图片编号107
        '           添加“打印条码”按钮，快捷键Alt+Q，工具栏图片编号107
        '.Add "单位导入(&I)110"
        '.Add "打印新条码(&P)107"
        '.Add "重新打条码(&Q)107"
        '2012-02-13 于登淼 ↑
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlb工具栏
    End With
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""

    '2012-05-23 翁乔 ↓↓↓
    '界面权限设置
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_受检者个人信息录入科_职业病史登记") = False Then
        ctlb工具栏.Buttons(2).Visible = False
        ctlb工具栏.Buttons(3).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_受检者个人信息录入科_修改") = False Then
        ctlb工具栏.Buttons(4).Visible = False
        ctlb工具栏.Buttons(5).Visible = False
    End If
'    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_受检者个人信息录入科_删除") = False Then
        ctlb工具栏.Buttons(6).Visible = False
        ctlb工具栏.Buttons(7).Visible = False
'    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_受检者个人信息录科入_导出") = False Then
        ctlb工具栏.Buttons(8).Visible = False
        ctlb工具栏.Buttons(9).Visible = False
    End If
    '2012-06-15 于登淼 ↓
    '打印功能取消
'''    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_受检者个人信息录入科_打印") = False Then
'''        ctlb工具栏.Buttons(10).Visible = False
'''        ctlb工具栏.Buttons(11).Visible = False
'''    End If
    '2012-06-15 于登淼 ↑
    Set lobjTmp = Nothing
    '2012-05-23 ↑↑↑

    '缺省显示最近一周的体检人员。
    '2012-06-15 于登淼 ↓
    '默认查询所有时间段内体检人员
    mstr开始日期 = Format(DateAdd("d", -14, Date), "yyyy-mm-dd")
    '之前显示的时间为固定，修改为缺省二周  2015-6-19
    'mstr开始日期 = Format(CDate("2012-06-01"), "yyyy-mm-dd")
    '2012-06-15 于登淼 ↑mstr截止日期 = Format(Date, "yyyy-mm-dd")
    mstr体检表名称 = ""
    mstr单位名称 = ""
    mstr姓名 = ""
    mstr体检单号 = ""
    mstr试管编号 = ""
    
    sub查询并显示
    
    ctlb工具栏.Buttons(4).Enabled = coptType(1).Value
    'cmnuItemRegister(3).Enabled = coptType(1).Value
    ctlb工具栏.Buttons(5).Enabled = coptType(1).Value
    ctlb工具栏.Buttons(6).Enabled = coptType(0).Value
    ctlb工具栏.Buttons(7).Enabled = coptType(0).Value
    'cmnuItemRegister(5).Enabled = coptType(0).Value
    'cmnuItemRegister(6).Enabled = coptType(0).Value

    '2012-04-14 于登淼 ↓
    '菜单里删除“打印”大项，包括(1)“体检表”，(2)“体检结果单”
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 于登淼 ↑

    '2012-02-23 于登淼 ↓
    cgrdMain.HighLight = flexHighlightWithFocus
    cgrdMain.SelectionMode = flexSelectionListBox
    '2012-02-23 于登淼 ↑
    访问记号 = 0

    Exit Sub
errHandler:
   sfsub错误处理 "职业病史录入", "frmcareerhstmanage", "form_load", Err.Number, Err.Description, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblninuse = False
    
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

Public Sub sub查询并显示()
    On Error GoTo errHandler
    '2012-06-15 于登淼 ↓
    '更改体检状态
    
    'Set mobjQueryResult = pobj业务对象.func职业病史管理界面查询(mstr开始日期, mstr截止日期, mstr体检表名称, mstr单位名称, mstr姓名, mstr体检单号, mstr试管编号, mstr系统编号)
    Set mobjQueryResult = pobj业务对象.func体检管理界面查询(mstr开始日期, mstr截止日期, mstr体检表名称, mstr单位名称, mstr姓名, mstr体检单号, mstr试管编号, mstr系统编号, mstr身份证号)
    '2012-06-15 于登淼 ↑
    
    sub显示查询结果

    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstmage", "sub查询并显示", Err.Number, Err.Description, True
End Sub

Private Sub sub显示查询结果()
    On Error GoTo errHandler
    Dim lstrWhere As String
    Dim lstrsql As String
    '2012-06-15 于登淼 ↓
    '更改体检状态
'''    If coptType(0).Value Then
'''        mobjQueryResult.Filter = "体检状态='待病史录入'"
'''    ElseIf coptType(1).Value Then
'''        mobjQueryResult.Filter = "体检状态='待体检'"
'''        'mobjQueryResult.Filter = "体检状态='已下结论' and 复查体检表名<>'' and 复查系统编号=''"
'''    Else
'''        mobjQueryResult.Filter = "体检状态='待下结论'"
'''        'mobjQueryResult.Filter = "(体检状态='已下结论' and  复查体检表名='') or (体检状态='已下结论' and 复查系统编号<>'')"
'''    End If
    If coptType(0).Value Then
        mobjQueryResult.Filter = "体检状态='未录入受检者个人信息'"
    ElseIf coptType(1).Value Then
'        mobjQueryResult.Filter = "体检状态='体检中'"
'说明： 体检中也可以补录未录入受检个人信息
'修改人: 罗李奎 2012 - 12 - 12 ↓ 2015-6-26  by lanchao 去掉'未录入去掉受检者个人信息'
        'mobjQueryResult.Filter = "体检状态='体检中'or 体检状态='未录入去掉受检者个人信息'"
        mobjQueryResult.Filter = "体检状态='体检中'"
  '修改人: 罗李奎 2012 - 12 - 12 ↑
     
    ElseIf coptType(2).Value Then
        mobjQueryResult.Filter = "体检状态='未下结论'"
    ElseIf coptType(3).Value Then
'        mobjQueryResult.Filter = "体检状态='已下结论'"    '将以下结论改为待复核  2015-11-27 by 牟俊
        mobjQueryResult.Filter = "体检状态='待复核'"
    ElseIf coptType(4).Value Then
        'mobjQueryResult.Filter = "体检状态='待复查'"
    '复检状态
    '修改者：罗李奎 2012-12-7 ↓
       '过滤出正在体检中的人
       'Bug号：0000072
    
        Dim lobjRec As Object
        
        lstrsql = "select a.系统编号,姓名,性别,年龄,单位名称,试管编号,体检表编号 as 体检表名, " _
                    & "convert(varchar(10),体检日期,120) as 体检日期, " _
                    & "体检结论,isnull(复查体检表编号,'') as 复查体检表编号, " _
                    & "isnull(复查系统编号,'') as 复查系统编号, " _
                    & "体检状态= '待复查' from 职业病体检_体检基本信息表 a inner join  职业病体检_体检人员基本信息表 b on a.系统编号=b.系统编号 and 复查状态 = '0'"
        lstrWhere = ""
        If mstr开始日期 <> "" Then
            lstrWhere = " and 体检日期>='" & mstr开始日期 & "'"
        End If
        If mstr截止日期 <> "" Then
            lstrWhere = " and 体检日期<='" & mstr截止日期 & "'"
        End If
        If mstr体检表名称 <> "" Then
            lstrWhere = " and a.体检表编号='" & mstr体检表名称 & "'"
        End If
        If mstr单位名称 <> "" Then
            lstrWhere = " and 单位名称 like '%" & mstr单位名称 & "%'" & ""
        End If
        If mstr姓名 <> "" Then
            lstrWhere = " and 姓名 like '%" & mstr姓名 & "%'" & ""
        End If
        If mstr系统编号 <> "" Then
            lstrWhere = " and a.系统编号='" & mstr系统编号 & "'"
        End If
        If lstrWhere <> "" Then
            lstrWhere = " where" & Right(lstrWhere, Len(lstrWhere) - 4)
        End If
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrsql)
        If Not lobjRec.EOF Then
            With cgrdMain
                Set .DataSource = lobjRec
                clblInfo = .Rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .Cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrdMain.Rows > 1 Then
                Dim i As Long
                Set mcolIndex = New Collection
                For i = 0 To cgrdMain.Cols - 1
                    mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
                Next
                cgrdMain.ColHidden(mcolIndex("试管编号")) = True
                cgrdMain.ColHidden(mcolIndex("体检结论")) = True
             End If
            
            Exit Sub
        Else
            cgrdMain.Rows = 1
            Exit Sub
        End If
         '修改者：罗李奎 2012-12-7 ↑
    End If
    '2012-06-15 于登淼 ↑

    With cgrdMain
        Set .DataSource = mobjQueryResult
        clblInfo = .Rows - 1
        
        .Col = 0
        .Sort = flexSortGenericDescending
        
        '2012-06-15 于登淼 ↓
        'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
        .AutoSize 0, .Cols - 1, 0, 0
        .ExplorerBar = flexExSort
        .DataMode = flexDMFree
        '2012-05-15 于登淼 ↑
    End With
         If cgrdMain.Rows > 1 Then
'            Dim i As Long
            Set mcolIndex = New Collection
            For i = 0 To cgrdMain.Cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            cgrdMain.ColHidden(mcolIndex("试管编号")) = True
            cgrdMain.ColHidden(mcolIndex("体检结论")) = True
        End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstmage", "sub显示查询结果", Err.Number, Err.Description, True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    ctlb工具栏.Width = Me.ScaleWidth - ctlb工具栏.Left * 2
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Select Case Operate
    '2012-06-15 于登淼 ↓
    '若当前选中某个体检人员，则将该体检人员基本信息传递到录入界面中
    Case "受检者信息录入"
        cmnuItemRegister_Click 5
    '2012-06-15 于登淼 ↑
    Case "查询"
        cmnuItemView_Click 1
        
    Case "职业病史登记"
        cmnuItemRegister_Click 1
        Cancel = True
        
    Case "复查登记"
        cmnuItemRegister_Click 3
    
    Case "修改"
        Cancel = True
        cmnuItemRegister_Click 5
    
    Case "删除"
        Cancel = True
        cmnuItemRegister_Click 6
    '修改人：罗李奎 2012-12-3  ↓
    '说明：退出时清空条件查询
    'bug 号 0000078
     Case "退出"
         Cancel = True
         cmnuItemRegister_Click 7
     ' 罗李奎  2012-12-3   ↑
    Case "导出"
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
        
    '2012-02-13 于登淼 ↓
    '修改内容：弹出打印条码窗口
    'Case "打印新条码"
       ' FrmPrintBarCode.Show
    'Case "重新打条码"
       ' FrmPrintBarCodeAgain.Show
    '2012-02-13 于登淼 ↑
    '2012-02-16 ↓
    'Case "单位导入"
        'frmImportExcel.Show
    '2012-02-16 于登淼 ↑
    Case "打印"
        Cancel = True
        If Not cgrdMain.Row > 0 Then Exit Sub
        frmCareerHstRegt.subPrint cgrdMain.TextMatrix(cgrdMain.Row, 0)
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstmage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub
