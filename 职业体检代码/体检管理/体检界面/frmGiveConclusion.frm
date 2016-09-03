VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGiveConclusion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "下体检结论"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11445
   ClipControls    =   0   'False
   Icon            =   "frmGiveConclusion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox cchk全选 
      Caption         =   "全选"
      Height          =   255
      Left            =   2520
      TabIndex        =   32
      Top             =   720
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.OptionButton coptType 
      Caption         =   "未下结论"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   31
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "已下结论"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   30
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   60
      TabIndex        =   21
      Top             =   7200
      Width           =   10335
      Begin VB.TextBox ctxtDoctor 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   345
         Left            =   3960
         TabIndex        =   27
         Top             =   240
         Width           =   1140
      End
      Begin VB.Frame cframPrintPaper 
         Appearance      =   0  'Flat
         Caption         =   "输出文书："
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   5160
         TabIndex        =   22
         Top             =   120
         Width           =   5055
         Begin VB.OptionButton coptPaper 
            BackColor       =   &H00F0F3E9&
            Caption         =   "打印体检结果通知单"
            Height          =   240
            Index           =   0
            Left            =   1920
            TabIndex        =   25
            Top             =   240
            Width           =   1980
         End
         Begin VB.OptionButton coptPaper 
            BackColor       =   &H00F0F3E9&
            Caption         =   "打印体检结果单"
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1650
         End
         Begin VB.OptionButton coptPaper 
            BackColor       =   &H00F0F3E9&
            Caption         =   "不打印"
            Height          =   240
            Index           =   2
            Left            =   4080
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   930
         End
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
         Height          =   345
         Left            =   1080
         TabIndex        =   26
         Top             =   240
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
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
         Format          =   69074944
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下结论时间："
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下结论医师："
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "列出体检人员"
      ForeColor       =   &H80000008&
      Height          =   3435
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   2355
      Begin VB.OptionButton coptBatch 
         Caption         =   "系统编号(条码号)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1485
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton coptBatch 
         Caption         =   "所有未下结论"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton coptBatch 
         Caption         =   "体检日期"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox ctxtNo 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   1800
         Width           =   2145
      End
      Begin VB.ComboBox ccmbSheet 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2220
      End
      Begin VB.CommandButton ccmdAddPerson 
         Caption         =   "添加(&Q)"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker cdtpStart 
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
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   69074944
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "选择体检表"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "体检项目结果："
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2925
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   7305
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdResult 
         Height          =   2505
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4419
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
         BackColor       =   16777215
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "^常规项目       |^结果           |^化验项目          |^结果           "
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
         ExplorerBar     =   0
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
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   1111
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchk刷条码 
         Caption         =   "刷条码"
         Height          =   375
         Left            =   9600
         TabIndex        =   33
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame cfram结论 
      Appearance      =   0  'Flat
      Caption         =   "结论(必须按“修改”按钮才可以修改）"
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   7320
      TabIndex        =   10
      Top             =   4320
      Width           =   3975
      Begin VB.CommandButton ccmdUpdateConclusion 
         Appearance      =   0  'Flat
         Caption         =   "修改(&M)"
         Enabled         =   0   'False
         Height          =   465
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin RichTextLib.RichTextBox ctxtDiagnosis 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1085
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmGiveConclusion.frx":0442
      End
      Begin RichTextLib.RichTextBox ctxtConclusion 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   873
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmGiveConclusion.frx":04DF
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
      Begin RichTextLib.RichTextBox ctxtTemplate 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmGiveConclusion.frx":057C
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复查体检表："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊断和处理意见："
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检结论："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   4200
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdPerson 
      Height          =   3225
      Left            =   2520
      TabIndex        =   19
      Top             =   960
      Width           =   8745
      _cx             =   87571521
      _cy             =   87561785
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
      BackColor       =   8454016
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   8454016
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
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSearch      =   2
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
End
Attribute VB_Name = "frmGiveConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mcolFieldIndex As Collection 'cgrdPerson网格中各列名所对应的列号。

Private mstr系统编号固定部分 As String
Private mbln可以取消结论 As Boolean  '表示当前用户是否具有取消结论的权限。
Private mblnInUse As Boolean

Private mobj记忆 As cls用户操作记忆 '修改：2001-12-29（增加该对象）。

Private mblnSys As Boolean

'功能：表明当前窗体是否已加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
Public Property Get pblnInUse() As Boolean
    On Error GoTo errHandler
    pblnInUse = mblnInUse
    Exit Property
errHandler:
    sfsub错误处理 "体检界面部件", "frmGiveConclusion", "Property Get pblnInUse", Err.Number, Err.Description, True
End Property



Private Sub cchk全选_Click()
    Dim i As Long
    For i = 1 To cgrdPerson.Rows - 1
        cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("允许办证"), i, mcolFieldIndex("允许办证")) = IIf(cchk全选.Value = 1, flexChecked, flexUnchecked)
    Next
End Sub

Private Sub cchk刷条码_Click()
    On Error Resume Next
    If coptBatch(2).Value Then
        ctxtNo.SetFocus
    End If
End Sub

'修改：2002-1-16（重新显示体检信息）。
Private Sub cgrdPerson_AfterSort(ByVal Col As Long, Order As Integer)
    On Error Resume Next
    If cgrdPerson.Row > 0 Then
        cgrdPerson_Click
    End If
End Sub

Private Sub cgrdPerson_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If coptType(0) Then
        If Col <> mcolFieldIndex("允许办证") Then Cancel = True
    Else
        Cancel = True
    End If
End Sub

Private Sub coptBatch_Click(Index As Integer)
    On Error Resume Next
    If coptBatch(Index).Value Then
        If Index = 0 Then
            cdtpStart.SetFocus
        ElseIf Index = 1 Then
'            ctxt单位名称.SetFocus
        Else
            ctxtNo.SetFocus
        End If
    End If
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error Resume Next
    subReset
    If coptType(0).Value Then
        '下结论
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
        ctbMain.Buttons(4).Enabled = True
        ctbMain.Buttons(5).Enabled = False
        Frame1(0).ForeColor = &HFF0000    '蓝色。
        Frame1(0).Enabled = True
        cdtpDate.Enabled = True
        coptBatch(1).Visible = True
    Else
        '取消结论
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
        ctbMain.Buttons(4).Enabled = False
        ctbMain.Buttons(5).Enabled = True
        Frame1(0).ForeColor = &HFF00FF    '红色。
        Frame1(0).Enabled = True
        cdtpDate.Enabled = False
        coptBatch(1).Visible = False
        If coptBatch(1).Value Then
            coptBatch(0).Value = True
        End If
    End If
    If coptBatch(0) Then
        cdtpStart.SetFocus
    ElseIf coptBatch(2) Then
        ctxtNo.SetFocus
    End If
End Sub

Private Sub ctxtNo_GotFocus()
    On Error Resume Next
    If cchk刷条码.Value = 1 Then
        ctxtNo = ""
    ElseIf ctxtNo.Text = "" Then
        ctxtNo.Text = mstr系统编号固定部分
        ctxtNo.SelLength = 0
        ctxtNo.SelStart = Len(mstr系统编号固定部分)
    End If
End Sub

Private Sub ctxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And ctxtNo <> "" Then
        ccmdAddPerson_Click
    End If
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
        .Add "去掉人员(&D)129"
        .Add "清空人员(&R)106"
        .Add "|"
        .Add "保存结论(&O)109"
        .Add "取消结论(&K)104"
        .Add "|"
        .Add "导出(&O)111"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
'        Set .c状态栏 = csbMain
    End With
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    '显示"下结论医师"为当前用户名。
    ctxtDoctor = um用户名
    cdtpDate.Value = Date
    cdtpStart.Value = Date
    
    '获取系统编号固定部分。
    Dim lobj体检 As Object '体检对象，获取系统编号的固定部分。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    mstr系统编号固定部分 = lobj体检.系统编号固定部分
    Set lobj体检 = Nothing
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    '判断是否有取消结论的权限。
    mbln可以取消结论 = umfunc校验用户权限("取消体检结论")
    If Not mbln可以取消结论 Then
        mbln可以取消结论 = umfunc校验用户权限("体检管理_取消体检结论")
    End If
    
    Dim lobj体检表模板集 As Object
    Dim lcolInfo As Collection
    Dim i As Long
    
    Set lobj体检表模板集 = CreateObject("体检对象.ClsMedicalExamTemplateSet")
    Set lcolInfo = lobj体检表模板集.元素集
    
    ccmbSheet.Clear
    '修改：2002-8-14（杨春）增加“<所有>”选择项。
    If lcolInfo.Count > 0 Then
        ccmbSheet.AddItem "<所有>"
    End If
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next i
    If ccmbSheet.ListCount > 1 Then
        ccmbSheet.ListIndex = 1
    End If

    '修改：2001-12-29（获取操作记忆值）。
    On Error Resume Next
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = um用户编号
    mobj记忆.业务名 = "体检管理"
    
    If mobj记忆.记忆项值("下结论时刷条码") = "是" Then
        cchk刷条码.Value = 1
    Else
        cchk刷条码.Value = 0
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmGiveConclusion", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If coptBatch(0) Then
        cdtpStart.SetFocus
    ElseIf coptBatch(1) Then
'        ctxt单位名称.SetFocus
    Else
        ctxtNo.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub

Private Sub cdtpStart_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdAddPerson.SetFocus
    End If
End Sub



Private Sub ccmdAddPerson_Click()
    Dim lobjRec As Object    '获取的可以下结论的体检记录。
    Dim llngMaxRow As Long
    Dim i As Long, j As Long
    
    On Error GoTo errHandler
    
'    If coptBatch(1).Value And ctxt单位名称 = "" Then
'        MsgBox "请输入单位名称！", vbOKOnly + vbExclamation, "系统提示"
'        ctxt单位名称.SetFocus
'        Exit Sub
'    End If
    If coptBatch(2).Value And ctxtNo = "" Then
        MsgBox "请输入系统编号，或刷入体检表上的条码！", vbOKOnly + vbExclamation, "系统提示"
        ctxtNo.SetFocus
        Exit Sub
    End If
    
    '获取输入的系统编号（范围）、日期的可以下体检结论的体检记录。
    If ctbMain.Buttons(5).Enabled Then
        '要取消结论，获取已下结论的体检记录。
        Set lobjRec = pobj业务对象.Func获取体检结论已确定的体检记录(IIf(coptBatch(0).Value, Format(cdtpStart.Value, "yyyy-mm-dd"), ""), "", IIf(ccmbSheet.ListIndex > 0, ccmbSheet.Text, ""), IIf(coptBatch(2).Value, ctxtNo.Text, ""))
    Else
        '要保存结论，获取还未下结论的体检记录。
        If Not coptBatch(1).Value Then
            Set lobjRec = pobj业务对象.Func获取已下结论但未确定的体检记录(IIf(coptBatch(0).Value, Format(cdtpStart.Value, "yyyy-mm-dd"), ""), "", IIf(ccmbSheet.ListIndex > 0, ccmbSheet.Text, ""), IIf(coptBatch(2).Value, ctxtNo.Text, ""))
        Else
            Set lobjRec = pobj业务对象.Func获取所有已下结论但未确定的体检记录()
        End If
    End If
    
    '修改：2001-11-15（杨春）为了提高显示效率，使用DataSource。
    cgrdPerson.Redraw = False
    
    If lobjRec.recordcount = 0 Then
        If ctbMain.Buttons(5).Enabled Then
            '取消结论。
            sffuncMsg "没有你指定范围内体检完毕，并且已下体检结论的体检记录。" & Chr(13) & Chr(10) & "请注意：你只能取消你自己所下的体检结论。", sf警告
        Else
            '下结论。
            sffuncMsg "没有你指定范围内体检完毕，并且未下体检结论的体检记录。" & Chr(13) & Chr(10) & "请进入常规、化验结果录入界面，调出你需要下结论的体检人员的体检记录，看是否已登记了所有体检项目的体检结果。", sf警告
        End If
        If cgrdPerson.Rows = 1 Then
            Set cgrdPerson.DataSource = lobjRec
            cgrdPerson.Rows = lobjRec.recordcount + 1
        End If
    Else
        '显示获取记录到cgrdPerson中。
        If cgrdPerson.Rows = 1 Then
            Set cgrdPerson.DataSource = lobjRec
            cgrdPerson.Rows = lobjRec.recordcount + 1
        Else
            '排除重复的纪录。
            gfsubAppendGridFromRecWithUnique cgrdPerson, lobjRec, mcolFieldIndex("系统编号")
        End If
    End If
    lobjRec.Close
    
    '获取cgrdPerson中各列的列号。
    Set mcolFieldIndex = New Collection
    For i = 0 To cgrdPerson.Cols - 1
        mcolFieldIndex.Add i, cgrdPerson.TextMatrix(0, i)
    Next
    
    '设置不正常得为红色。
    For i = 1 To cgrdPerson.Rows - 1
        Select Case cgrdPerson.TextMatrix(i, mcolFieldIndex("体检结论"))
        Case "正常", "健康"
        Case Else
            cgrdPerson.Cell(flexcpBackColor, i, 0, i, cgrdPerson.Cols - 1) = &H8A5AFA
        End Select
        If coptType(0).Value Then
            '默认正常的允许办证。
            Select Case cgrdPerson.TextMatrix(i, mcolFieldIndex("体检结论"))
            Case "正常", "健康"
                cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("允许办证"), i, mcolFieldIndex("允许办证")) = flexChecked
            Case Else
                cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("允许办证"), i, mcolFieldIndex("允许办证")) = flexUnchecked
            End Select
        Else
            If cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("允许办证"), i, mcolFieldIndex("允许办证")) <> flexChecked And cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("允许办证"), i, mcolFieldIndex("允许办证")) <> flexUnchecked Then
                If cgrdPerson.TextMatrix(i, mcolFieldIndex("允许办证")) = "1" Then
                    cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("允许办证"), i, mcolFieldIndex("允许办证")) = flexChecked
                Else
                    cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("允许办证"), i, mcolFieldIndex("允许办证")) = flexUnchecked
                End If
            End If
        End If
        cgrdPerson.TextMatrix(i, mcolFieldIndex("允许办证")) = ""
    
    Next
    cdtpStart.SetFocus
    cgrdPerson.AutoSize 0, cgrdPerson.Cols - 1
    If cgrdPerson.Rows = 1 Then
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
    Else
        ctbMain.Buttons(2).Enabled = True
    End If
'    cgrdPerson.ColHidden(mcolFieldIndex("系统编号")) = True
    cgrdPerson.ColHidden(0) = True
    cgrdPerson.Redraw = True
    cgrdPerson.Editable = True
    If coptBatch(0).Value Then
        cdtpStart.SetFocus
    ElseIf coptBatch(1).Value Then
'        ctxt单位名称.SetFocus
    Else
        ctxtNo.SetFocus
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmGiveConclusion", "ccmdAddPerson_Click", 6666, lstrError, False
    cgrdPerson.Redraw = True
    Exit Sub
    Resume
End Sub


Private Sub ccmdUpdateConclusion_Click()
    Dim llngRow As Long '体检人员网格的当前行号。
    
    On Error GoTo errHandler
    llngRow = cgrdPerson.Row
    
    '设置frmUpdateConclusion的属性。
    With frmUpdateConclusion
        .系统编号 = cgrdPerson.TextMatrix(llngRow, mcolFieldIndex("系统编号"))
        .体检结论 = ctxtConclusion.Text
        .诊断处理意见 = ctxtDiagnosis.Text
        .复查体检表名 = ctxtTemplate.Text
    End With
    
    '启动修改体检结论窗体。
    frmUpdateConclusion.Show 1, Me
    
    '获取frmUpdateConclusion相应属性，并修改界面上的ctxtConclusion和ctxtDiagnosis。
    With frmUpdateConclusion
        ctxtConclusion.Text = .体检结论
        ctxtDiagnosis.Text = .诊断处理意见
        ctxtTemplate.Text = .复查体检表名
    End With
    
    '修改cgrdPerson的当前行。
    With cgrdPerson
        .TextMatrix(llngRow, mcolFieldIndex("体检结论")) = ctxtConclusion.Text
        .TextMatrix(llngRow, mcolFieldIndex("诊断和处理意见")) = ctxtDiagnosis.Text
        .TextMatrix(llngRow, mcolFieldIndex("复查体检表名")) = ctxtTemplate.Text
        
        Select Case .TextMatrix(llngRow, mcolFieldIndex("体检结论"))
        Case "正常", "健康"
            '绿色。
            .Cell(flexcpBackColor, llngRow, 0, llngRow, .Cols - 1) = &HC0FFC0
            ctxtConclusion.BackColor = &HC0FFC0
            .Cell(flexcpChecked, llngRow, mcolFieldIndex("允许办证"), llngRow, mcolFieldIndex("允许办证")) = flexChecked
            
        Case Else
            '红色。
            .Cell(flexcpBackColor, llngRow, 0, llngRow, .Cols - 1) = &H8A5AFA
            ctxtConclusion.BackColor = &H8A5AFA
            .Cell(flexcpChecked, llngRow, mcolFieldIndex("允许办证"), llngRow, mcolFieldIndex("允许办证")) = flexUnchecked
        End Select
        
    End With
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmGiveConclusion", "ccmdUpdateConclusion_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub cgrdPerson_Click()
    Dim lobjRec As Object
    Dim lobj体检表 As Object   'clsMedicalExamSheet体检表对象，以获取当前体检人员的体检结果。
    Dim lcolInfo As Collection '当前体检人员的体检项目及其结果,item:clsFactTestItem。
    Dim lobjItem As Variant    'clsFactTestItem，lcolInfo中的元素。
    Dim i As Long
    
    On Error GoTo errHandler
    If cgrdPerson.Row <= 0 Then Exit Sub
    
    MousePointer = 11
'    csbMain.Panels(1) = "正在获取当前体检人员的体检结果，请稍候..."
    
    '获取并显示当前体检人员的体检结论到修改区。
    ctxtConclusion.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("体检结论"))
    ctxtDiagnosis.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("诊断和处理意见"))
    ctxtTemplate.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("复查体检表名"))
    
    '创建体检表对象。
    Set lobj体检表 = CreateObject("体检对象.clsMedicalExamSheet")
    lobj体检表.系统编号 = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("系统编号"))
    
    cgrdResult.Redraw = False
    
    
    '优化算法：2001-4-22。
    Set lobjRec = lobj体检表.优化体检项目集("常规")
    If lobjRec.recordcount > 0 Then
        cgrdResult.Rows = lobjRec.recordcount + 1
    Else
        cgrdResult.Rows = 1
    End If
    i = 1
    Do While Not lobjRec.EOF
        cgrdResult.TextMatrix(i, 0) = lobjRec!体检项目名称
        cgrdResult.TextMatrix(i, 1) = IIf(IsNull(lobjRec!体检结果), "", lobjRec!体检结果)
        If cgrdResult.TextMatrix(i, 1) <> "" Then
            cgrdResult.TextMatrix(i, 1) = cgrdResult.TextMatrix(i, 1) & IIf(IsNull(lobjRec!单位), "", lobjRec!单位)
        End If
        If IIf(IsNull(lobjRec!单项结论), "", lobjRec!单项结论) = "不合格" Then
            cgrdResult.Cell(flexcpBackColor, i, 1, i, 1) = &H8A5AFA
        Else
            cgrdResult.Cell(flexcpBackColor, i, 1, i, 1) = vbWhite
        End If
        i = i + 1
        lobjRec.movenext
    Loop
    
    '优化算法：2001-4-22。
    Set lobjRec = lobj体检表.优化体检项目集("化验")
    i = 1
    Do While Not lobjRec.EOF
        If i = cgrdResult.Rows Then
            cgrdResult.Rows = cgrdResult.Rows + 1
        End If
        cgrdResult.TextMatrix(i, 2) = lobjRec!体检项目名称
        cgrdResult.TextMatrix(i, 3) = IIf(IsNull(lobjRec!体检结果), "", lobjRec!体检结果)
        If cgrdResult.TextMatrix(i, 3) <> "" Then
            cgrdResult.TextMatrix(i, 3) = cgrdResult.TextMatrix(i, 3) & IIf(IsNull(lobjRec!单位), "", lobjRec!单位)
        End If
        If IIf(IsNull(lobjRec!单项结论), "", lobjRec!单项结论) = "不合格" Then
            cgrdResult.Cell(flexcpBackColor, i, 3, i, 3) = &H8A5AFA
        Else
            cgrdResult.Cell(flexcpBackColor, i, 3, i, 3) = vbWhite
        End If
        i = i + 1
        lobjRec.movenext
    Loop
    Do While i < cgrdResult.Rows
        cgrdResult.TextMatrix(i, 2) = ""
        cgrdResult.TextMatrix(i, 3) = ""
        i = i + 1
    Loop
    
    '刷新体检结果网格。
    cgrdResult.Redraw = True
    
    '若是取消结论，显示下结论日期，体检医师。
    If ctbMain.Buttons(5).Enabled Then
        If IsDate(cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("下结论日期"))) Then
            cdtpDate.Value = Format(cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("下结论日期")), "yyyy-mm-dd")
        Else
            cdtpDate.Value = Format(Date, "yyyy-mm-dd")
        End If
        ctxtDoctor.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("下结论医师姓名"))
    End If
    
    '设置"去掉人员"、“清空人员”按钮可用。
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(2).Enabled = True
    If ctbMain.Buttons(5).Enabled Then
        '取消结论，“修改”按钮不可用。
        ccmdUpdateConclusion.Enabled = False
    Else
        '下结论，设置“修改”按钮可用。
        ccmdUpdateConclusion.Enabled = True
    End If
    
    Select Case cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("体检结论"))
    Case "正常", "健康"
        '绿色。
        ctxtConclusion.BackColor = &HC0FFC0
    Case Else
        '红色。
        ctxtConclusion.BackColor = &H8A5AFA
    End Select
    
    
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmGiveConclusion", "cgrdPerson_Click", 6666, lstrError, False
    
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub





Private Sub Form_Resize()
    On Error Resume Next
    cgrdPerson.Width = Me.ScaleWidth - cgrdPerson.Left - 60
    Frame2.Width = Me.ScaleWidth - Frame2.Left - 60
    Frame2.Top = Me.ScaleHeight - Frame2.Height - 60
    Frame1(1).Top = Frame2.Top - Frame1(1).Height - 60
    cfram结论.Top = Frame1(1).Top
    cfram结论.Left = Me.ScaleWidth - cfram结论.Width - 60
    Frame1(1).Width = cfram结论.Left - Frame1(1).Left - 60
    
    cgrdResult.Width = Frame1(1).Width - cgrdResult.Left - 60
    
    cgrdPerson.Height = Frame1(1).Top - cgrdPerson.Top - 30
    Frame1(0).Height = Frame1(1).Top - Frame1(0).Top - 30
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '修改：2002-9-26（杨春）保存操作记忆值。
    mobj记忆.sub覆盖记忆值 "下结论时刷条码", IIf(cchk刷条码.Value = 1, "是", "否")
    
    '释放模块级对象。
    Set mobjGUI = Nothing
    
    Unload frmUpdateConclusion
    
    '设置标志pblnInUse。
    mblnInUse = False
End Sub

'功能：恢复界面初始态。
Private Sub subReset()
    On Error GoTo errHandler
    On Error Resume Next
    
    '清空cgrdPers，录入区，只有"下结论"、"取消结论"按钮可用(若"mbln可以取消结论"=false，则"取消结论"按钮不可用)。
    cgrdPerson.Rows = 1
    cgrdResult.Rows = 1
    ctxtConclusion.Text = ""
    ctxtDiagnosis.Text = ""
    ctxtTemplate.Text = ""
    cdtpDate.Value = Format(Date, "yyyy-mm-dd")
    ctxtDoctor.Text = um用户名
    
    ccmdUpdateConclusion.Enabled = False
    Frame1(0).Enabled = False
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmGiveConclusion", "subReset", 6666, lstrError, True
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    
    On Error GoTo errHandler
    Select Case Operate
    Case "保存结论"
        Dim lstr文书名称 As String
        Dim lbln是否预览文书 As Boolean
        
        If cgrdPerson.Rows = 1 Then
            sffuncMsg "请先添加需要下结论的人员到网格中！", sf警告
            Exit Sub
        End If
        If cgrdPerson.Rows > 2 Then
            If Not sffuncMsg("你确认要保存网格中所有体检人员的体检结论吗？" & Chr(13) & Chr(10) & "你若选择“是”，则这些人的体检结论会保存下来，并且不能再修改其体检结果。若要修改体检结果，只能先取消结论。", sf询问) Then
                Exit Sub
            End If
        End If
        MousePointer = 11
        
        '获取需要打印的文书。
        If coptPaper(0).Value Then
            lstr文书名称 = "体检结果通知单"
        ElseIf coptPaper(1).Value Then
            lstr文书名称 = "体检结果单"
        Else
            lstr文书名称 = ""
        End If
        
        '保存体检人员网格中所有人的体检结论。
        Dim llngRow As Long
        Dim lbln允许办证 As Boolean
        i = 1
        llngRow = cgrdPerson.Rows - 1
        For i = 1 To llngRow
'            csbMain.Panels(1) = "正在确定第" & i & "个人（共" & (llngRow) & "个人）的体检结论。 "
            '保存体检结论。
            With cgrdPerson
                If .Cell(flexcpChecked, 1, mcolFieldIndex("允许办证"), 1, mcolFieldIndex("允许办证")) = flexChecked Then
                    lbln允许办证 = True
                Else
                    lbln允许办证 = False
                End If
                pobj业务对象.Sub确定体检结论 .TextMatrix(1, mcolFieldIndex("系统编号")), .TextMatrix(1, mcolFieldIndex("体检结论")), Format(cdtpDate.Value, "yyyy-mm-dd"), .TextMatrix(1, mcolFieldIndex("诊断和处理意见")), .TextMatrix(1, mcolFieldIndex("复查体检表名")), lstr文书名称, lbln是否预览文书, lbln允许办证
            End With
            cgrdPerson.RemoveItem 1
        Next
        
        '清空网格。
        cgrdPerson.Rows = 1
        cgrdResult.Rows = 1
        ccmdUpdateConclusion.Enabled = False
        ctxtConclusion.Text = ""
        ctxtDiagnosis.Text = ""
        ctxtTemplate.Text = ""
'        csbMain.Panels(1) = "保存结论完毕。"
    
        MousePointer = 0
        Cancel = True
    Case "取消结论"
        If cgrdPerson.Row > 0 Then
            If Not mbln可以取消结论 Then
                MsgBox "对不起，你没有取消结论的权限！", vbOKOnly + vbInformation, "系统提示"
                Exit Sub
            End If
            '询问。
            If sffuncMsg("你确认要取消当前体检人员“" & cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("姓名")) & " ”的体检结论吗？", sf询问) Then
                '通过业务对象取消当前体检人员的体检结论。
                pobj业务对象.Sub取消体检结论 cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("系统编号"))
            
                '把当前体检人员从体检人员网格删除。
                cgrdPerson.RemoveItem cgrdPerson.Row
                cgrdResult.Rows = 1
            End If
        Else
            If cgrdPerson.Rows = 1 Then
                sffuncMsg "请先添加需要取消结论的人员到网格中！", sf警告
                Exit Sub
            End If
        End If
        Cancel = True
    Case "去掉人员"
        '从网格中删除当前行。
        If cgrdPerson.Row > 0 Then
            cgrdPerson.RemoveItem cgrdPerson.Row
        End If
        If cgrdPerson.Rows = 1 Then
            ctbMain.Buttons(1).Enabled = False
            ctbMain.Buttons(2).Enabled = False
        End If
        Cancel = True
    Case "清空人员"
        cgrdPerson.Rows = 1
        cgrdResult.Rows = 1
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
        ccmdUpdateConclusion.Enabled = False
        Cancel = True
    Case "导出"
        Dim lstrFile As String
        ccmdFile.Filter = "Excel文件 (*.xls)|*.xls|文本文件 (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            cgrdPerson.SaveGrid lstrFile, flexFileTabText, True
        End If
    
    End Select
    
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmGiveConclusion", "ctbMain_ButtonClick", 6666, lstrError, False
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
