VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegisterManage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "职业健康体检登记"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13755
   Icon            =   "frmRegisterManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ctxtLabelNum 
      Height          =   270
      Left            =   8760
      TabIndex        =   11
      Text            =   "2"
      Top             =   900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton coptType 
      Caption         =   "体检中"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton coptType 
      Caption         =   "待复查"
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "已下结论"
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "未下结论"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   6015
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   12375
      _cx             =   2088785220
      _cy             =   2088774002
      Appearance      =   1
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VB.OptionButton coptType 
      Caption         =   "未打清单"
      Height          =   300
      Index           =   1
      Left            =   9480
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "未拍照"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   1005
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
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Label clblLabelNum 
      Caption         =   "标签数"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   7440
      TabIndex        =   4
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总记录数："
      Height          =   180
      Left            =   6360
      TabIndex        =   3
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmRegisterManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

'查询条件
Public selectneednum As String '根据身份证查询时用到的系统编号
Private mstr开始日期 As String
Public mstr截止日期 As String
Private mstr体检表名称 As String
Private mstr单位名称 As String
Private mstr姓名 As String
Private mstr体检单号 As String
Private mstr试管编号 As String
Private mstr系统编号 As String
Private mstr身份证号 As String
'查询结果
Private mobjQueryResult As Object
Private mintState As Integer
Private mcolIndex As New Collection
Private hasQueryWindows As Boolean
Private printLabelNumLegal As Boolean

'功能：返回当前窗体是否已经加载标志。这是系统平台所要求的。
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
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
    sfsub错误处理 "职业病界面", "frmRegisterManage", "cmnuItemPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemRegister_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '初检登记
        访问标志 = 0 '用于控制焦点
        FrmRegister.pstr系统编号 = ""
        
        
        'frm刷条码.Show 1, Me
'        FrmRegister.Move 1000, 500
        FrmRegister.Show 0, Me
'        FrmRegister.Move 700, 350
        '重新查询。
'        sub查询并显示
        '2012-07-13 于登淼 ↓
        '没加入这些功能，故去掉
'''    Case 2 '年检登记
'''        If cgrdMain.Row >= 1 Then
'''            FrmRegister.pstr系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
'''        Else
'''            FrmRegister.pstr系统编号 = ""
'''        End If
'''        FrmRegister.Show 1, Me
'''
'''        '重新查询。
'''        sub查询并显示
'''
'''    Case 3 '复查登记
'''        If cgrdMain.Row < 1 Then
'''            MsgBox "没有需要复查的人！", vbOKOnly + vbExclamation, "系统提示"
'''            Exit Sub
'''        End If
'''        FrmRegisterAgain.pstr旧系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
'''        FrmRegisterAgain.Show 1, Me
        '2012-07-13 于登淼 ↑
    Case 5 '修改
        If cgrdMain.Row < 1 Then
            MsgBox "没有需要修改的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        FrmRegister.clblsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        '2012-07-11 于登淼 ↓
        '修改时，默认刷身份证修改
        FrmRegister.Check身份证.Value = 1
        '2012-07-11 于登淼 ↑
        
        访问标志 = 1
        FrmRegister.Show 0, Me
        
        '重新查询。
        'sub显示查询结果
    
    Case 6 '删除
        If cgrdMain.Row < 1 Then
            MsgBox "没有可以删除的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        If coptType(0) Or coptType(1) Or coptType(2) Or coptType(3) Then
            If MsgBox("你确认要删除该体检记录吗？一旦删除后将不能恢复！", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
                
                pobj业务对象.sub删除体检登记 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
                
                cgrdMain.RemoveItem cgrdMain.Row
                clblInfo = cgrdMain.rows - 1
                
                '2012-06-13 于登淼 ↓
                '删除体检记录时，删除该体检人员照相照片和身份证照片
                Dim lobjDelPhoto As Object
                Set lobjDelPhoto = CreateObject("职业病对象.clsPersonExamed")
                lobjDelPhoto.func删除身份证照片 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号")) & "IDcard", "职业病体检"
                lobjDelPhoto.func删除照相照片 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号")), "职业病体检"
                sub查询并显示
                '2012-06-13 于登淼 ↑
            End If
        ElseIf coptType(4).Value Then
            MsgBox "已下体检结论的记录不允许删除！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        ElseIf coptType(5).Value Then
            If MsgBox("你确认要删除该体检记录吗？一旦删除后将不能恢复！", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
                dafuncGetData "update 职业病体检_体检基本信息表 set 复查状态='' where 系统编号='" & cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号")) & "'"
            End If
            sub查询并显示
            Exit Sub
        End If
    '2012-08-17 于登淼 ↓
    '增加复查登记功能
    Case 7  '复查
        If cgrdMain.rows < 1 Then
            MsgBox "没有可以复查的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        If cgrdMain.Row < 1 Or cgrdMain.Row > cgrdMain.rows - 1 Then
            MsgBox "请选择要复查的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        访问标志 = 2
        FrmRegisterAgain.clblsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        FrmRegisterAgain.pstr复查系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("复查系统编号"))
        
        '复查时，默认刷身份证修改
'        FrmRegisterAgain.Check身份证.Value = 1
        FrmRegisterAgain.Show 0, Me
        
        '重新查询。
        sub查询并显示
    '2012-08-17 于登淼 ↑
    Case 8
        If cgrdMain.Row < 1 Then
            MsgBox "没有需要照相的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        pstrPhoto = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        FrmPhoto.clblsysno.Text = pstrPhoto
        '2012-07-11 于登淼 ↓
        '修改时，默认刷身份证修改
'        frmPhoto.Check身份证.Value = 1
        '2012-07-11 于登淼 ↑
        
'        访问标志 = 1
        FrmPhoto.Show 0, Me
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "cmnuItemRegister_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '查询
        hasQueryWindows = False
        With frmQuery
            '显示旧的查询条件。
            .pstr开始日期 = mstr开始日期
            .pstr截止日期 = mstr截止日期
            .pstr体检表名称 = mstr体检表名称
            .pstr姓名 = mstr姓名
            .pstr单位名称 = mstr单位名称
            .pstr体检单号 = mstr体检单号
            .pstr试管编号 = mstr试管编号
            .pstr系统编号 = mstr系统编号
           .pstr身份证号 = mstr身份证号
            '获取新的查询条件。
            .Show 1, Me
            If .pblnOk Then
                hasQueryWindows = True
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
            'add 20150504 从查询界面点击查询后刷新cgrdmian，然后如果有记录则直接选中第一行，然后传值给frmregister
            If cgrdMain.rows > 1 And mstr身份证号 <> "" Then
            '只要查询到有记录就能显示  2015-11-17 by 牟俊
            If cgrdMain.rows > 1 Then
                cgrdMain.Row = 1
                cmnuItemRegister_Click 5
                selectneednum = mstr系统编号
            Else
'             sub查询并显示
              MsgBox "没有查询到信息，请确认身份信息！", vbOKOnly + vbExclamation, "系统提示"
              sub全部显示
              Exit Sub
            End If
            End If
            End If
            
        End With
        
    Case 2 '刷新
        sub显示查询结果
    Case 4
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub



Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    sub显示查询结果
    If cgrdMain.rows > 1 Then
        cgrdMain.ColHidden(mcolIndex("复查体检表编号")) = False
        cgrdMain.ColHidden(mcolIndex("复查系统编号")) = False
        cgrdMain.ColHidden(mcolIndex("体检状态")) = False
        If coptType(0).Value = True Then
            cgrdMain.ColHidden(mcolIndex("复查体检表编号")) = True
            cgrdMain.ColHidden(mcolIndex("复查系统编号")) = True
            cgrdMain.ColHidden(mcolIndex("体检状态")) = True
        End If
    End If
    '2012-06-15 于登淼 ↓
    'coptType控件内容全部更改，相应操作也进行更改
    ctlb工具栏.Buttons(4).Enabled = coptType(5).Value
'    cmnuItemRegister(3).Enabled = coptType(1).Value

    
    ctlb工具栏.Buttons(6).Enabled = coptType(0).Value Or coptType(1).Value 'Or coptType(5).Value
    ctlb工具栏.Buttons(7).Enabled = coptType(0).Value Or coptType(1).Value Or coptType(5).Value
    '2012-12-18 刘云乐
    'Bug No:0000033
    '工具栏界面与菜单编辑器不一致
'    cmnuItemRegister(5).Enabled = coptType(0).Value Or coptType(1).Value
'    cmnuItemRegister(6).Enabled = coptType(0).Value Or coptType(1).Value Or coptType(5).Value
    '2012-12-18 刘云乐
'    ctlb工具栏.Buttons(3).Enabled = Not coptType(5).Value
   ctlb工具栏.Buttons(3).Enabled = coptType(0).Value
'    cmnuItemRegister(1).Enabled = Not coptType(5).Value
    ctlb工具栏.Buttons(10).Enabled = coptType(0).Value
    ctlb工具栏.Buttons(12).Enabled = coptType(3).Value Or coptType(2).Value
'    ctlb工具栏.Buttons(13).Enabled = coptType(1).Value Or coptType(2).Value
    ctlb工具栏.Buttons(15).Enabled = coptType(0).Value
    '2012-06-15 于登淼 ↑

    '2012-07-06 于登淼 ↓
    '增加打印标签张数的使用判断
    'clblLabelNum.Visible = ctlb工具栏.Buttons(13).Enabled
    'ctxtLabelNum.Visible = ctlb工具栏.Buttons(13).Enabled
    '2012-07-06 于登淼 ↑
    
    '2012-04-14 于登淼 ↓
    '菜单里删除“打印”大项，包括(1)“体检表”，(2)“体检结果单”
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 于登淼 ↑
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "coptType_Click", Err.Number, Err.Description, False
End Sub

Private Sub ctxtLabelNum_LostFocus()
    Dim num As Integer
    printLabelNumLegal = True
    If IsNumeric(ctxtLabelNum.Text) = True Then
        num = CInt(ctxtLabelNum.Text)
        If num <> ctxtLabelNum.Text Then
            ctxtLabelNum.Text = "2"
            printLabelNumLegal = False
        Else
            If num <= 0 Then ctxtLabelNum.Text = "2": printLabelNumLegal = False
        End If
    Else
        ctxtLabelNum.Text = "2"
        printLabelNumLegal = False
    End If
End Sub

Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    '显示进度。
    frmProcess.proPercent.Max = 8
    frmProcess.Label1.Caption = "正在初始化界面，请等待..."
    frmProcess.proPercent.Value = 1
    frmProcess.Show
    DoEvents
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    With lcol工具栏按钮
        .Add "查询(&Q)108"
        .Add "新增(&A)103"
'        .Add "|"
        .Add "照相(&R)101"  '3
        .Add "复查登记(&F)103"
        .Add "|"
        .Add "修改" '6
        .Add "删除" '7
        .Add "|"
        .Add "导出(&O)113"  '9
        .Add "单位导入(&I)102"   '10
        .Add "|"
        .Add "打印清单(&P)107"    '12
        .Add "打印标签(&U)107"    '13
        .Add "|"
        .Add "校核通过(&J)106"    '15
        .Add "|"
        .Add "退出"
    End With
    '2012-12-18 刘云乐
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlb工具栏
    End With
    frmProcess.proPercent.Value = 2
    DoEvents
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""
    ctlb工具栏.Buttons(13).Visible = True
    ctlb工具栏.Buttons(15).Visible = False
    ctlb工具栏.Buttons(16).Visible = False
'    ctlb工具栏.Buttons(13).Enabled = False    '屏蔽"打印标签" 2015-11-19 by 牟俊
    '2012-05-22 翁乔 ↓↓↓ 2012-06-15 于登淼 微调权限设置
    '界面权限设置
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_初检登记") = False Then
        ctlb工具栏.Buttons(2).Visible = False
        ctlb工具栏.Buttons(3).Visible = False
'        cmnuItemRegister(1).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_复查登记") = False Then
        ctlb工具栏.Buttons(4).Visible = False
        ctlb工具栏.Buttons(5).Visible = False
        'cmnuItemRegister(3).Checked = False
    End If
    frmProcess.proPercent.Value = 3
    DoEvents
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_修改") = False Then
        ctlb工具栏.Buttons(6).Visible = False
'        cmnuItemRegister(5).Visible = False
'        cmnuItemRegister(5).Checked = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_删除") = False Then
        ctlb工具栏.Buttons(7).Visible = False
        ctlb工具栏.Buttons(8).Visible = False
'        cmnuItemRegister(6).Visible = False
'        cmnuItemRegister(6).Checked = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_导出") = False Then
        ctlb工具栏.Buttons(9).Visible = False
    End If
    frmProcess.proPercent.Value = 4
    DoEvents
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_单位导入") = False Then
        ctlb工具栏.Buttons(10).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_打印清单") = False Then
        ctlb工具栏.Buttons(12).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_打印试管标签") = False Then
'        ctlb工具栏.Buttons(13).Visible = False
        ctlb工具栏.Buttons(14).Visible = False
    End If
    frmProcess.proPercent.Value = 5
    DoEvents
'''    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_体检登记_年检登记") = False Then
'''        cmnuItemRegister(2).Checked = False
'''    End If
    Set lobjTmp = Nothing
    '2012-05-22 ↑↑↑

    '缺省显示最近一周的体检人员。
    '2012-06-15 于登淼 ↓
    '默认查询所有时间段内体检人员
    mstr开始日期 = Format(DateAdd("d", -30, Now), "yyyy-mm-dd")
'    mstr开始日期 = mstr开始日期 & " 00:00:00"
'    mstr开始日期 = Format(CDate("2000-01-01 00:00:00"), "yyyy-mm-dd hh:mm:ss")
    '2012-06-15 于登淼 ↑
    mstr截止日期 = Format(Now, "yyyy-mm-dd")
'    mstr截止日期 = Format(Now, "yyyy-mm-dd hh:mm:ss")
    mstr体检表名称 = ""
    mstr单位名称 = ""
    mstr姓名 = ""
    mstr体检单号 = ""
    mstr试管编号 = ""
    mstr身份证号 = ""
    frmProcess.proPercent.Value = 6
    DoEvents
    sub查询并显示
    frmProcess.proPercent.Value = 7
    DoEvents
    ctlb工具栏.Buttons(4).Enabled = coptType(1).Value
    'cmnuItemRegister(3).Enabled = coptType(1).Value
    
    ctlb工具栏.Buttons(6).Enabled = coptType(0).Value
    ctlb工具栏.Buttons(7).Enabled = coptType(0).Value
'    cmnuItemRegister(5).Enabled = coptType(0).Value
'    cmnuItemRegister(6).Enabled = coptType(0).Value

    '2012-04-14 于登淼 ↓
    '菜单里删除“打印”大项，包括(1)“体检表”，(2)“体检结果单”
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 于登淼 ↑
 
    '2012-02-23 于登淼 ↓
    cgrdMain.HighLight = flexHighlightWithFocus
    cgrdMain.SelectionMode = flexSelectionListBox
    cgrdMain.AllowBigSelection = False
    '2012-02-23 于登淼 ↑
    
    '2012-06-20 于登淼 ↓
    '初始化一系列单选项
    coptType(0).Value = 1
    coptType_Click (0)
    '2012-06-20 于登淼 ↑
    frmProcess.proPercent.Value = 8
    Unload frmProcess
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

Public Sub sub查询并显示()
    On Error GoTo errHandler
    If hasQueryWindows = False Then mstr截止日期 = Format(Now, "yyyy-mm-dd")       'mstr截止日期 = Now
    Set mobjQueryResult = pobj业务对象.func体检管理界面查询(mstr开始日期, mstr截止日期, mstr体检表名称, mstr单位名称, mstr姓名, mstr体检单号, mstr试管编号, mstr系统编号, mstr身份证号)
    
    sub显示查询结果
    
'    Dim i As Long
'    If cgrdMain.rows > 1 Then
'        Set mcolIndex = New Collection
'        For i = 0 To cgrdMain.cols - 1
'            mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'        Next
'        cgrdMain.ColHidden(mcolIndex("试管编号")) = True
'        cgrdMain.ColHidden(mcolIndex("体检结论")) = True
'    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "sub查询并显示", Err.Number, Err.Description, True
End Sub
Public Sub sub全部显示()
    On Error GoTo errHandler
    If hasQueryWindows = False Then mstr截止日期 = Now
    Set mobjQueryResult = pobj业务对象.func体检管理界面查询(mstr开始日期, mstr截止日期, "", "", "", "", "", "", "")
    
    sub显示查询结果
    
'    Dim i As Long
'    If cgrdMain.rows > 1 Then
'        Set mcolIndex = New Collection
'        For i = 0 To cgrdMain.cols - 1
'            mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'        Next
'        cgrdMain.ColHidden(mcolIndex("试管编号")) = True
'        cgrdMain.ColHidden(mcolIndex("体检结论")) = True
'    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "sub查询并显示", Err.Number, Err.Description, True
End Sub

Private Sub sub显示查询结果()
    Dim i As Integer
    On Error GoTo errHandler
    
    '2012-06-15 于登淼 ↓
    '更改体检状态
'''    If coptType(0).Value Then
'''        mobjQueryResult.Filter = "体检状态='未校核'"
'''    ElseIf coptType(1).Value Then
'''        mobjQueryResult.Filter = "体检状态='已下结论' and 复查体检表名<>'' and 复查系统编号=''"
'''    Else
'''        mobjQueryResult.Filter = "(体检状态='已下结论' and  复查体检表名='') or (体检状态='已下结论' and 复查系统编号<>'')"
'''    End If
    If coptType(0).Value Then
        mobjQueryResult.Filter = "体检状态='未校核'or 体检状态='未打清单'"
'    ElseIf coptType(1).Value Then
'        mobjQueryResult.Filter = "体检状态='未打清单'"
    ElseIf coptType(2).Value Then
        mobjQueryResult.Filter = "体检状态='体检中' or 体检状态='未录入受检者个人信息'"
    ElseIf coptType(3).Value Then
        mobjQueryResult.Filter = "体检状态='未下结论'"
    ElseIf coptType(4).Value Then
       ' mobjQueryResult.Filter = "体检状态='未复核' or 体检状态='已复核'"
       '修改：罗李奎 2012-12-7 ↓
       '过滤体检状态已经下结论的信息
       'bug号：000072
       mobjQueryResult.Filter = "体检状态='待复核' or 体检状态='已复核' or 体检状态='已发报告'"
         '修改：罗李奎 2012-12-7 ↑
    ElseIf coptType(5).Value Then
'        mobjQueryResult.Filter = "体检状态='待复查'"
        '复检状态，翁乔；2012-10-30
        Dim lobjRec As Object
        dasubSetQueryTimeout 600
        Dim lstrSql As String
        lstrSql = "select a.系统编号,姓名,性别,年龄,单位名称,危害因素,现工种, 体检表编号 as 体检表名,b.公民身份号码 ," _
                    & "convert(varchar(10),体检日期,120) as 体检日期, " _
                    & "isnull(复查体检表编号,'') as 复查体检表编号, " _
                    & "isnull(复查系统编号,'') as 复查系统编号,a.复查原因,a.复查项目," _
                    & "体检状态= '待复查' from 职业病体检_体检基本信息表 a inner join  职业病体检_体检人员基本信息表 b on a.系统编号=b.系统编号 and 复查状态 = '0'"
        Set lobjRec = dafuncGetData(lstrSql)
        cgrdMain.rows = 1
        
        If Not lobjRec.EOF Then
            With cgrdMain
                Set .DataSource = lobjRec
                If cgrdMain.rows > 1 Then
                    Set mcolIndex = New Collection
                    For i = 0 To cgrdMain.cols - 1
                        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
                    Next
                End If
                clblInfo = .rows - 1
                .Col = 0
'                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
'                .DataMode = flexDMFree
                clblInfo = .rows - 1
            End With
            
            Exit Sub
        Else
            cgrdMain.rows = 1
            clblInfo = cgrdMain.rows - 1
            Exit Sub
        End If
    End If
    '2012-06-15 于登淼 ↑
    
    With cgrdMain
        Set .DataSource = mobjQueryResult
        If cgrdMain.rows > 1 Then
            Set mcolIndex = New Collection
            For i = 0 To cgrdMain.cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            cgrdMain.ColHidden(mcolIndex("试管编号")) = True
'        '如果没查到人，提示查询条件有误  2016-1-6 by 牟俊
'        Else
'            MsgBox "查询条件有误", vbOKOnly
        End If
        clblInfo = .rows - 1
        .Col = 0
'        .Sort = flexSortGenericDescending
        '2012-06-15 于登淼 ↓
        'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
        .AutoSize 0, .cols - 1, 0, 0
        .ExplorerBar = flexExSort
'        .DataMode = flexDMFree
        '2012-05-15 于登淼 ↑
        
    End With

        If cgrdMain.rows > 1 Then
'        Dim i As Long
            Set mcolIndex = New Collection
            For i = 0 To cgrdMain.cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            cgrdMain.ColHidden(mcolIndex("试管编号")) = True
            cgrdMain.ColHidden(mcolIndex("体检结论")) = True
        End If
    
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "sub显示查询结果", Err.Number, Err.Description, True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Dim lobjFile As Object
    Dim i As Integer
    
    Select Case Operate
    Case "查询"
        cmnuItemView_Click 1
    
    Case "新增"
        frmAddRegister.Show 1
   
    Case "照相"
'    cmnuItemRegister_Click 1   '初见登记
        cmnuItemRegister_Click 8
        Cancel = True
        
    Case "修改"
        Cancel = True
        cmnuItemRegister_Click 5
    
    Case "删除"
        Cancel = True
        cmnuItemRegister_Click 6
    Case "导出"
        If cgrdMain.rows <= 1 Then
            MsgBox "没有需要导出的记录！", vbOKOnly + vbExclamation, "系统提示"
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
    
    '2012-06-15 于登淼 ↓ 打印条码内容全部取消
'''    '2012-02-13 于登淼 ↓
'''    '修改内容：弹出打印条码窗口
'''    Case "打印新条码"
'''        FrmPrintBarCode.Show
'''    Case "重新打条码"
'''        FrmPrintBarCodeAgain.Show
'''    '2012-02-13 于登淼 ↑
    '2012-06-15 于登淼 ↑
    
    '2012-02-16 ↓
    Case "单位导入"
        frmImportExcel.Show
    '2012-02-16 于登淼 ↑
'        sub查询并显示
    '2012-06-15 于登淼 ↓ 更改：2012-10-17 罗李奎
    Case "打印清单"
        '调用打印函数；罗李奎；2012-10-25
        Dim j As Integer
        Dim para系统编号 As Collection
        Set para系统编号 = New Collection

        Set lobjFile = CreateObject("职业病文书.cls文书")
        For j = 0 To cgrdMain.SelectedRows - 1
             para系统编号.Add (cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("系统编号")))
        Next j


'        Dim j As Integer
'        Dim para系统编号 As Collection
'        Set para系统编号 = New Collection
'
'        Set lobjFile = CreateObject("职业病文书.cls文书")
'             For j = 1 To cgrdMain.SelectedRows
'             para系统编号.Add (cgrdMain.TextMatrix(j, mcolIndex("系统编号")))
'            Next j
        lobjFile.func打印体检清单 para系统编号

'             lobjFile.func打印体检清单(cgrdMain.TextMatrix(cgrdMain.SelectedRows, mcolIndex("系统编号")))
        Set lobjFile = Nothing

        '更改当前体检状态。打印清单之后，就进入体检状态。
        For j = 0 To cgrdMain.SelectedRows - 1
             pobj业务对象.func写入单人当前体检状态 cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("系统编号")), 2
        Next j
        sub查询并显示
        
    '2012-06-15 于登淼 ↑ 更改：2012-10-17  罗李奎
    Case "打印标签"
'         Dim j As Integer
        '控制打印标签内容
        frmAssayDeptSelect.Show 1
        Dim paraSelectedDeptName As Collection
        If cgrdMain.SelectedRows < 1 Then Exit Sub
        If frmAssayDeptSelect.pblnOk Then
            If frmAssayDeptSelect.selectedDeptName.Count > 1 Then
                Set paraSelectedDeptName = frmAssayDeptSelect.selectedDeptName
                For j = 0 To cgrdMain.SelectedRows - 1
                    For i = 2 To paraSelectedDeptName.Count
                        Set lobjFile = CreateObject("职业病文书.cls文书")
'                       lobjFile.func打印试管标签 cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("系统编号")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("姓名")), paraSelectedDeptName.Item(i)
                        lobjFile.func打印试管标签 cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("系统编号")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("姓名")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("性别")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("年龄")), paraSelectedDeptName.Item(i)
                    Next i
                Next j
            End If
        Else
            Exit Sub
        End If
        
'''        '控制打印标签张数
'''        ctxtLabelNum_LostFocus  '将焦点转移，执行ctxtLabelNum_lostfocus判断值是否合法
'''        If printLabelNumLegal = False Then Exit Sub
'''        Dim printTimes As Integer
'''        printTimes = CInt(ctxtLabelNum.Text)
'''        Set lobjFile = CreateObject("职业病文书.cls文书")
'''        While printTimes > 0
'''            lobjFile.func打印试管标签 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号")), cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("姓名"))
'''            printTimes = printTimes - 1
'''        Wend
        
    Case "校核通过"
        If cgrdMain.Row < 1 Then
            MsgBox "没有需要校核的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        mintState = 1
        For i = 0 To cgrdMain.SelectedRows - 1
            pobj业务对象.func写入单人当前体检状态 cgrdMain.TextMatrix(cgrdMain.SelectedRow(i), mcolIndex("系统编号")), mintState
            pobj业务对象.func写入校核人信息 cgrdMain.TextMatrix(cgrdMain.SelectedRow(i), mcolIndex("系统编号")), um用户编号
        Next
        sub查询并显示
    '2012-08-17 于登淼 ↓
    '增加复查登记功能
    Case "复查登记"
        cmnuItemRegister_Click 7
    '2012-08-17 于登淼 ↑
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

