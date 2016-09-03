VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintPaper 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "打印体检文书"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frmPrintPaper.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox cchkPreview 
      Caption         =   "打印前预览"
      Height          =   285
      Left            =   8760
      TabIndex        =   20
      Top             =   840
      Width           =   1395
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   5400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CheckBox cchkPrintAll 
      Caption         =   "全部打印"
      Height          =   300
      Left            =   7200
      TabIndex        =   19
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox ccmbPaper 
      Height          =   300
      ItemData        =   "frmPrintPaper.frx":0442
      Left            =   1230
      List            =   "frmPrintPaper.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   885
      Width           =   3615
   End
   Begin VB.Frame cframSearch 
      Appearance      =   0  'Flat
      Caption         =   "查询体检："
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   10545
      Begin VB.CheckBox cchk体检结果 
         Caption         =   "阳性"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox cchk体检结果 
         Caption         =   "正常"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox ctxt姓名 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   6
         Top             =   720
         Width           =   1260
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   6600
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "..."
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   690
      End
      Begin MSComCtl2.DTPicker cdtpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   375
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   129236992
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.TextBox ctxtEndNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7440
         TabIndex        =   8
         Top             =   720
         Width           =   1740
      End
      Begin VB.TextBox ctxtStartNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   7
         Top             =   690
         Width           =   1860
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "按系统编号和姓名查询"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   705
         Width           =   2175
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "按单位和日期查询"
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   435
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.CommandButton ccmdSearch 
         Caption         =   "查询(F2)"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   680
         Width           =   1050
      End
      Begin VB.Label clblInfo 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   6240
         TabIndex        =   23
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至："
         Height          =   180
         Left            =   7080
         TabIndex        =   17
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Left            =   4155
         TabIndex        =   16
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
         Height          =   180
         Index           =   0
         Left            =   2355
         TabIndex        =   15
         Top             =   435
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检单位："
         Height          =   180
         Left            =   5670
         TabIndex        =   14
         Top             =   435
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar ctlbTool 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   5055
      Left            =   0
      TabIndex        =   21
      Top             =   2760
      Width           =   10815
      _cx             =   69094020
      _cy             =   69083860
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
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择格式："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmPrintPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：邓恒

Private WithEvents mobjGUI As cls界面通用对象 '界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj体检集 As Object
Private mstr系统编号固定部分 As String

Private mblnInUse As Boolean
Private mblnSys As Boolean
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchk体检结果_Click(Index As Integer)
    If mblnSys Then Exit Sub
    ccmdSearch_Click
End Sub

Private Sub ctxtEndNo_GotFocus()
    On Error Resume Next
    If ctxtEndNo = "" Then
'        ctxtEndNo = mstr系统编号固定部分
'        ctxtEndNo.SelStart = Len(ctxtEndNo)
'        ctxtEndNo.SelLength = 0
    End If

End Sub

Private Sub ctxtStartNo_GotFocus()
    On Error Resume Next
    If ctxtStartNo = "" Then
'        ctxtStartNo = mstr系统编号固定部分
'        ctxtStartNo.SelStart = Len(ctxtStartNo)
'        ctxtStartNo.SelLength = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyF2
        ccmdSearch_Click
 
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    Dim lcolInfo As New Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    MousePointer = 11
    
    '设置窗体正在使用的标志。
    mblnInUse = True
'    csbMain.Panels(1) = "窗体正在初始化，请稍侯..."
    
    Set mobj体检集 = CreateObject("体检对象.clsMedicalExamSet")
   
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    Set lcolInfo = New Collection
    With lcolInfo
        .Add "预览"
        .Add "打印"
        .Add "|"
        .Add "导出(&O)111"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlbTool
        'Set .c状态栏 = csbMain
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcolInfo, ""
    End With
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    
    '加入单位名称
    Set lcolInfo = pobj业务对象.当日工作记忆簿.单位名称集
    For i = 1 To lcolInfo.Count
        ccmbUnit.AddItem lcolInfo(i)
    Next i
    
    Dim lobj体检 As Object
    '创建体检对象，获取系统编号前面固定部分。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    mstr系统编号固定部分 = lobj体检.系统编号固定部分
    Set lobj体检 = Nothing
    
    '清空
    cgrdMain.Rows = 1
    cdtpDate.Value = Format(Now, "yyyy-mm-dd")
    
    If ccmbPaper.ListCount > 0 Then
        ccmbPaper.ListIndex = 0
    End If
    cgrdMain.Editable = False
    
    MousePointer = 0
'    csbMain.Panels(1) = "请先选择体检表格式。"
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检管理界面", "frmPrintPaper", "Form_Load", 6666, lstrError, False
    MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    cframSearch.Width = Me.ScaleWidth - cframSearch.Left - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj体检集 = Nothing
    mblnInUse = False
End Sub

Private Sub ccmbPaper_Click()
    On Error Resume Next
    cframSearch.Enabled = True
    cgrdMain.Editable = True
    cgrdMain.Rows = 1
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    clblInfo.Caption = ""
    mblnSys = True
    If ccmbPaper.Text = "体检登记表" Or ccmbPaper.Text = "体检单" Then
        cchk体检结果(0).Value = 1
        cchk体检结果(1).Value = 1
    End If
    mblnSys = False
'    csbMain.Panels(1) = "请输入查询条件，然后按“查询”按钮。"
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error Resume Next
    gfsubShowComboList ccmbUnit
    
End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ccmbUnit_LostFocus()
    On Error GoTo errHandler
    Dim i As Integer
    If Trim(ccmbUnit.Text) = "" Then Exit Sub
    
    '判断录入的单位是否在列表中存在，不存在则加入列表
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '加入ccmbUnit
        ccmbUnit.AddItem ccmbUnit.Text
    End If

    Exit Sub
errHandler:
    
End Sub

Private Sub ccmdLocateUnit_Click()
    Dim lobjRec As Object  '单位定位返回的结果记录。
    
    On Error GoTo errHandler
    
    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
            
            '把定位的单位加入工作记忆簿。
            ccmbUnit_LostFocus
        End If
    End If
    
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmdSearch_Click()
    Dim lobj体检集  As Object
    Dim lstr系统编号 As String
    Dim lstrError As String
    
    Dim i As Integer
    On Error GoTo errHandler
    
'    csbMain.Panels(1) = "正在查询数据库，请稍侯..."
    MousePointer = 11
    cgrdMain.Rows = 1
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    
    '清空体检集的属性。
    mobj体检集.subClear
    
    If coptChoise(0).Value Then
        '按体检日期和单位
        mobj体检集.从体检日期 = Format(cdtpDate.Value, "yyyy-mm-dd")
        mobj体检集.到体检日期 = Format(cdtpDate.Value, "yyyy-mm-dd")
        mobj体检集.单位名称 = ccmbUnit.Text
    Else
        '按起始系统编号和结束系统编号
        mobj体检集.从系统编号 = ctxtStartNo.Text
        mobj体检集.到系统编号 = ctxtEndNo.Text
        mobj体检集.姓名 = ctxt姓名.Text
    End If
    
    If ccmbPaper.Text = "体检登记表" Or ccmbPaper.Text = "体检单" Then
        '若是查询体检登记表，只能是未下完体检结论的。
        mobj体检集.体检状态 = P_LOGIN_STATUS & "," & P_EXAMING_STATUS & "," & P_CONCLUED_STATUS
    ElseIf ccmbPaper.Text = "体检结果单" Then
        '若是查询体检结果单，只能是一下完体检结论的。
        mobj体检集.体检状态 = P_ENDED_STATUS
    End If
    
        
    Set lobj体检集 = mobj体检集.元素集old("系统编号,试管编号,姓名,性别,年龄,单位名称,体检单号,体检表名称,体检日期=convert(varchar(10),体检日期,20),体检结论=isnull(体检结论,'')")

    
    If cchk体检结果(0).Value = 1 And cchk体检结果(1).Value = 0 Then
        lobj体检集.Filter = "体检结论='正常' or 体检结论=''"
    ElseIf cchk体检结果(0).Value = 0 And cchk体检结果(1).Value = 1 Then
        lobj体检集.Filter = "体检结论<>'正常' and 体检结论<>''"
    
    End If
    
    If lobj体检集.RecordCount = 0 Then
        '没找到相应体检人员
        lstrError = "未查找到可以打印该类文书的体检人员，请重新输入查找条件！"
        If ccmbPaper.Text = "体检登记表" Or ccmbPaper.Text = "体检单" Then
            lstrError = "未查找到可以打印该类文书的体检人员（还未下体检结论的），请重新输入查找条件！"
        Else
            lstrError = "未查找到可以打印该类文书的体检人员（已下体检结论的），请重新输入查找条件！"
        End If
        Err.Raise 6666, , lstrError
    Else
        '加入到表格中
        cgrdMain.FormatString = ""
        Set cgrdMain.DataSource = lobj体检集
'        gfsubLoadGridFromRec cgrdMain, lobj体检集, False, "系统编号,姓名,性别,年龄,单位名称,体检表名称,体检日期,体检结论,需要复查,已经复查"
'        If cgrdMain.Rows > 1 Then
'            cgrdMain.Rows = cgrdMain.Rows - 1
'        End If
        ctlbTool.Buttons(1).Enabled = True
        ctlbTool.Buttons(2).Enabled = True
    End If
    clblInfo.Caption = "查询结果：" & cgrdMain.Rows - 1 & "人次。"
errHandler:
    If Err <> 0 Then
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检管理界面", "frmPrintPaper", "ccmdSearch_Click", 6666, lstrError, False
    End If
    Set lobj体检集 = Nothing
'    csbMain.Panels(1) = ""
    MousePointer = 0
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub coptChoise_Click(Index As Integer)
    On Error Resume Next
    cgrdMain.Rows = 1
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    If coptChoise(0).Value Then
        cdtpDate.Enabled = True
        ccmdLocateUnit.Enabled = True
        ccmbUnit.Enabled = True
        ctxtStartNo.Enabled = False
        ctxtEndNo.Enabled = False
        ctxt姓名.Enabled = False
        cdtpDate.SetFocus
    Else
        cdtpDate.Enabled = False
        ccmdLocateUnit.Enabled = False
        ccmbUnit.Enabled = False
        ctxtStartNo.Enabled = True
        ctxtEndNo.Enabled = True
        ctxt姓名.Enabled = True
        mblnSys = True
        cchk体检结果(0).Value = 1
        cchk体检结果(1).Value = 1
        mblnSys = False
        
        ctxt姓名.SetFocus
    End If
End Sub

Private Sub ctxtEndNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ctxtStartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtEndNo.SetFocus
    End If
End Sub


'功能：处理工具栏上按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lcol编号  As Collection
    Dim i As Integer
    On Error GoTo errHandler
    
    Select Case Operate
        Case "预览"
            Set lcol编号 = New Collection
            If cgrdMain.Row < 1 Then
                Err.Raise 6666, , "打印预览，必须选择要打印的体检记录。"
            Else
                lcol编号.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
            End If
            '打印预览。
            pobj业务对象.Sub打印文书 ccmbPaper.Text, lcol编号, False, True
            Cancel = True
            
        Case "打印"
'            csbMain.Panels(1) = "打印" & Trim(ccmbPaper.Text) & "中，请稍侯..."
            '全部打印时加入所有的编号
            Set lcol编号 = New Collection
            If cchkPrintAll.Value = 1 Then
                For i = 1 To cgrdMain.Rows - 1
                    lcol编号.Add cgrdMain.TextMatrix(i, 0)
                Next i
            Else
                If cgrdMain.Row < 1 Then
                    Err.Raise 6666, , "你没有选择“全部打印”，必须选择要打印的体检记录。"
                Else
                    lcol编号.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
                End If
            End If
            '打印
            If cchkPreview.Value = 1 Then
                If lcol编号.Count > 1 Then
                    If Not sffuncMsg("打印多个文书，不能进行预览。你是否继续？", sf询问) Then
                        Err.Raise 6666, , "你取消了打印操作。"
                    End If
                End If
                pobj业务对象.Sub打印文书 ccmbPaper.Text, lcol编号, True
            Else
                pobj业务对象.Sub打印文书 ccmbPaper.Text, lcol编号, False
            End If
            '打印完毕清除该条数据
            If cchkPrintAll.Value = 1 Then
                cgrdMain.Rows = 1
            Else
                cgrdMain.RemoveItem cgrdMain.Row
            End If
'            csbMain.Panels(1) = Trim(ccmbPaper.Text) & "打印完毕，请继续操作"
            Cancel = True
            
        Case "导出"
            Dim lstrFile As String
            ccmdFile.Filter = "Excel文件 (*.xls)|*.xls|文本文件 (*.txt)|*.txt"
            ccmdFile.ShowSave
            lstrFile = ccmdFile.FileName
            If lstrFile <> "" Then
                cgrdMain.SaveGrid lstrFile, flexFileTabText, True
            End If
            
    End Select
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检管理界面", "frmPrintPaper", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
