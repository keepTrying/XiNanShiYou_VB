VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmRegisterManage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "体检管理"
   ClientHeight    =   7635
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11070
   Icon            =   "frmRegisterManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton coptType 
      Caption         =   "已下结论"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "待复查"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "未下结论"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   10815
      _cx             =   88492676
      _cy             =   88484845
      _ConvInfo       =   1
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
      BackColorBkg    =   16777215
      BackColorAlternate=   15791081
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
      SelectionMode   =   1
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
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   1440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1005
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   6240
      TabIndex        =   6
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总记录数："
      Height          =   180
      Left            =   5280
      TabIndex        =   5
      Top             =   720
      Width           =   900
   End
   Begin VB.Menu cmnuView 
      Caption         =   "查看   "
      Begin VB.Menu cmnuItemView 
         Caption         =   "查询(&Q)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "刷新"
         Index           =   2
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "退出(&X)"
         Index           =   4
      End
   End
   Begin VB.Menu cmnuRegister 
      Caption         =   "体检登记   "
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "初检登记(&N)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "年检登记(&Y)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "复查登记(&R)"
         Index           =   3
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "修改(&U)"
         Index           =   5
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "删除(&D)"
         Index           =   6
      End
   End
   Begin VB.Menu cmnuPrint 
      Caption         =   "打印"
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "体检表"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "体检结果单"
         Index           =   2
      End
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
Private mstr开始日期 As String
Private mstr截止日期 As String
Private mstr体检表名称 As String
Private mstr单位名称 As String
Private mstr姓名 As String
Private mstr体检单号 As String
Private mstr试管编号 As String
Private mstr系统编号 As String
'查询结果
Private mobjQueryResult As Object

Private mcolIndex As New Collection

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
    sfsub错误处理 "体检界面", "frmRegisterManage", "cmnuItemPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemRegister_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '初检登记
        FrmRegister.pstr系统编号 = ""
        FrmRegister.Show 1, Me
        
        '重新查询。
        sub查询并显示
        
    Case 2 '年检登记
        If cgrdMain.Row >= 1 Then
            FrmRegister.pstr系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        Else
            FrmRegister.pstr系统编号 = ""
        End If
        FrmRegister.Show 1, Me
        
        '重新查询。
        sub查询并显示
    
    Case 3 '复查登记
        If cgrdMain.Row < 1 Then
            MsgBox "没有需要复查的人！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        FrmRegisterAgain.pstr旧系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        FrmRegisterAgain.Show 1, Me
        
    Case 5 '修改
        If cgrdMain.Row < 1 Then
            MsgBox "没有需要修改的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        FrmEditRegister.系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        FrmEditRegister.Show 1, Me
        
        '重新查询。
        sub查询并显示
    
    Case 6 '删除
        If cgrdMain.Row < 1 Then
            MsgBox "没有可以删除的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        If coptType(0) Then
            If MsgBox("你确认要删除该体检记录吗？一旦删除后将不能恢复！", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
                pobj业务对象.sub删除体检登记 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
                oesubSave "删除体检人员信息，其系统编号为：" & cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号")) & "，姓名为：" & cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("姓名")), "删除体检人员"
                cgrdMain.RemoveItem cgrdMain.Row
            End If
        Else
            MsgBox "已下体检结论的记录不允许删除！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
    
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面", "frmRegisterManage", "cmnuItemRegister_Click", Err.Number, Err.Description, False
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
            .pstr系统编号 = mstr系统编号
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
    sfsub错误处理 "体检界面", "frmRegisterManage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    sub显示查询结果
    
    ctlb工具栏.Buttons(4).Enabled = coptType(1).Value
    cmnuItemRegister(3).Enabled = coptType(1).Value
    
    ctlb工具栏.Buttons(6).Enabled = coptType(0).Value
    ctlb工具栏.Buttons(7).Enabled = coptType(0).Value
    cmnuItemRegister(5).Enabled = coptType(0).Value
    cmnuItemRegister(6).Enabled = coptType(0).Value
    
    cmnuItemPrint(1).Enabled = coptType(0).Value
    cmnuItemPrint(2).Enabled = coptType(2).Value
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面", "frmRegisterManage", "coptType_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    '修改：2002-7-1（杨春）简化取消结论的操作。该为操作单选框。
    With lcol工具栏按钮
        .Add "查询(&Q)108"
        .Add "|"
        .Add "初检登记(&R)101"
        .Add "复查登记(&R)103"
        .Add "|"
        .Add "修改"
        .Add "删除"
        .Add "|"
        .Add "导出(&O)111"
        .Add "传输(&O)112"
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

    '缺省显示最近一周的体检人员。
    mstr开始日期 = Format(DateAdd("d", -7, Date), "yyyy-mm-dd")
    mstr截止日期 = Format(Date, "yyyy-mm-dd")
    mstr体检表名称 = ""
    mstr单位名称 = ""
    mstr姓名 = ""
    mstr体检单号 = ""
    mstr试管编号 = ""
    
    sub查询并显示
    
    ctlb工具栏.Buttons(4).Enabled = coptType(1).Value
    cmnuItemRegister(3).Enabled = coptType(1).Value
    
    ctlb工具栏.Buttons(6).Enabled = coptType(0).Value
    ctlb工具栏.Buttons(7).Enabled = coptType(0).Value
    cmnuItemRegister(5).Enabled = coptType(0).Value
    cmnuItemRegister(6).Enabled = coptType(0).Value

    cmnuItemPrint(1).Enabled = coptType(0).Value
    cmnuItemPrint(2).Enabled = coptType(2).Value

    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmRegistermanage", "Form_Load", 6666, lstrError, False
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
    Set mobjQueryResult = pobj业务对象.func体检管理界面查询(mstr开始日期, mstr截止日期, mstr体检表名称, mstr单位名称, mstr姓名, mstr体检单号, mstr试管编号, mstr系统编号)
    
    sub显示查询结果

    Dim i As Long
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.Cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面", "frmRegisterManage", "sub查询并显示", Err.Number, Err.Description, True
End Sub

Private Sub sub显示查询结果()
    On Error GoTo errHandler
    If coptType(0).Value Then
        mobjQueryResult.Filter = "体检状态='未下结论'"
    ElseIf coptType(1).Value Then
        mobjQueryResult.Filter = "体检状态='已下结论' and 复查体检表名<>'' and 复查系统编号=''"
    Else
        mobjQueryResult.Filter = "(体检状态='已下结论' and  复查体检表名='') or (体检状态='已下结论' and 复查系统编号<>'')"
    End If
    Set cgrdMain.DataSource = mobjQueryResult
    
    clblInfo = cgrdMain.Rows - 1

    Exit Sub
errHandler:
    sfsub错误处理 "体检界面", "frmRegisterManage", "sub显示查询结果", Err.Number, Err.Description, True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Select Case Operate
    Case "查询"
        cmnuItemView_Click 1
        
    Case "初检登记"
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
    Case "导出"
        Dim lstrFile As String
        ccmdFile.Filter = "Excel文件 (*.xls)|*.xls|文本文件 (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            cgrdMain.SaveGrid lstrFile, flexFileTabText, True
        End If
    Case "传输"
        frmTransfer.Show 1
    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面", "frmRegisterManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub
