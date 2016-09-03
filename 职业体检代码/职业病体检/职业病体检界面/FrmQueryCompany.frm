VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmQueryCompany 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "职业健康体检-单位统计"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11895
      Begin VSFlex8Ctl.VSFlexGrid cgrdList 
         Height          =   4965
         Left            =   0
         TabIndex        =   12
         Top             =   840
         Width           =   11895
         _cx             =   20981
         _cy             =   8758
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
         SelectionMode   =   0
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
      Begin VB.Timer Timer1 
         Left            =   6960
         Top             =   0
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         TabIndex        =   1
         Top             =   -120
         Width           =   11895
         Begin VB.Frame Frame4 
            Caption         =   "查询条件"
            Height          =   735
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   11775
            Begin VB.CommandButton ccmdQuery 
               Caption         =   "查询"
               Height          =   375
               Left            =   10320
               TabIndex        =   14
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox flag名称 
               Caption         =   "Check1"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.CheckBox flag日期 
               Caption         =   "Check1"
               Enabled         =   0   'False
               Height          =   375
               Left            =   4440
               TabIndex        =   5
               Top             =   240
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.TextBox ctxtCompanyName 
               Height          =   375
               Left            =   1080
               TabIndex        =   4
               Top             =   240
               Width           =   2175
            End
            Begin VB.CommandButton CmdFCompany 
               Caption         =   "单位定位"
               Height          =   375
               Left            =   3360
               TabIndex        =   3
               Top             =   240
               Width           =   975
            End
            Begin MSComCtl2.DTPicker DTP截止时间 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   375
               Left            =   8040
               TabIndex        =   7
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               Format          =   60489728
               CurrentDate     =   41027
            End
            Begin MSComCtl2.DTPicker DTP开始时间 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   375
               Left            =   5520
               TabIndex        =   8
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               Format          =   60489728
               CurrentDate     =   41027
            End
            Begin VB.Label Label3 
               Caption         =   "至"
               Height          =   255
               Left            =   7800
               TabIndex        =   11
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label2 
               Caption         =   "体检时间"
               Height          =   255
               Left            =   4680
               TabIndex        =   10
               Top             =   345
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "单位名称"
               Height          =   255
               Left            =   360
               TabIndex        =   9
               Top             =   360
               Width           =   855
            End
         End
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5565
      Top             =   4485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQueryCompany.frx":0000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQueryCompany.frx":005E
            Key             =   "Back"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Height          =   600
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   3120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Lab记录数量 
      Height          =   255
      Left            =   9600
      TabIndex        =   16
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Lab记录数 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "FrmQueryCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'窗体：职业病体检单位人员信息查询打印界面
'功能：对职业病体检单位人员信息的查询和打印
'作者：陶露
'时间：2012-04-28
'备注：暂无

Option Explicit
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Public mblnInUse As Boolean
Dim lojb查询统计函数 As Object    '查询统计函数
Private indX, indY As Integer       '记录鼠标点击vsflexgrid的坐标。
Private pstrPerson As String
Private mobjRec As Object
'该界面共用对象
Private pobj体检 As Object
Private pobjItem As Object
Private pobj体检表模板 As Object
Private pobj体检结果业务 As Object
Private pobj科室 As Object

Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property

'通过定位单位名称进行条件查询
Private Sub ccmdQuery_Click()
'    If cgrdList.rows > 1 Then Exit Sub
    sub开始查询
End Sub

Private Sub cgrdList_DblClick()
    indX = cgrdList.MouseRow
    indY = cgrdList.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < cgrdList.rows And indY >= 0 And indY < cgrdList.cols Then
        pstrPerson = cgrdList.TextMatrix(indX, 0)
'        sub列出单位信息
    End If
End Sub



Private Sub CmdFCompany_Click()
    Dim lobjRec As Object                       '单位定位返回的结果记录。
    
    On Error GoTo errHandler
'    Set lobjRec = pobj业务对象.func单位定位      '启动单位定位界面。  注释于2015-11-2 因为要调用新的界面（单位定位查询界面）

    frmQueryCompanyLocation.Show 1, Me            '调用单位定位查询界面，写于2015-11-2  （此过程仅仅添加了这一句，同时也仅仅注释掉上面那句，其它未动）
    
    
    
    '以下内容未改动过，还是原来启动单位定位界面时的内容   2015-11-2
    '获取定位的单位，显示在“单位名称”录入框中。(暂时只显示“单位名称”)
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxtCompanyName.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
        End If
    End If
    Exit Sub
errHandler:
    'If Err.Number = 0 Then Exit Sub
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "Form1", "CmdFCompany_Click", 6666, lstrError, False
End Sub

Private Sub CmdFCompany_LostFocus()
    flag名称.Value = 1
End Sub

Private Sub DTP截至时间_GotFocus()
    flag日期.Value = 1
End Sub

Private Sub ctxtCompanyName_Change()
    
'    If Trim(ctxtCompanyName.Text) <> "" Then
'        ccmdQuery.Enabled = True
'    Else
'        ccmdQuery.Enabled = False
'    End If

End Sub

Private Sub ctxtCompanyName_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 And ctxtCompanyName.Text <> "" Then
        sub开始查询
    End If
    
End Sub

Private Sub DTP开始时间_GotFocus()
    flag日期.Value = 1
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
    With lcol工具栏按钮
        .Add "预览报告(&N)108"
        .Add "|"
        .Add "打印报告(&M)107"
        .Add "|"
        .Add "导出(@X)113"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlb工具栏
    End With
    
    '按住ctrl可选择多行数据
    cgrdList.HighLight = flexHighlightWithFocus
    cgrdList.SelectionMode = flexSelectionListBox
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""
    '体检医师显示框显示当前用户名
'    ctxtDoctor.Text = um用户名
'    ctxtDoctor.Enabled = False
    '显示当前日期
    DTP开始时间.Value = DateAdd("m", -1, Date)
    DTP截止时间.Value = Date
    
    '查询条件初始化
    '创建查询统计函数对象
    Set lojb查询统计函数 = CreateObject("职业病界面.clsQueryStatis")

    '变量初始化
    Set pobj体检 = CreateObject("职业病对象.clsMedicalExam")
    Set pobj体检表模板 = CreateObject("职业病对象.clsMedicalExamTemplate")
    Set pobj体检结果业务 = CreateObject("职业病体检结果录入.clsCommon")
    Set pobj科室 = pobjDict.Fetch("职业病体检科室字典")
    
    '表格初始化；翁乔；2012-11-01
    '按住ctrl可选择多行数据
    cgrdList.HighLight = flexHighlightWithFocus
    cgrdList.SelectionMode = flexSelectionListBox
    cgrdList.cols = 0
    With cgrdList
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "系统编号"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "姓名"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "性别"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "年龄"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "危害因素"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "职业分类"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "单位名称"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检日期"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    
    
    '2012-05-21 陶露
    '界面权限设置
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_单位统计_打印") = False Then
        ctlb工具栏.Buttons(3).Visible = False
        ctlb工具栏.Buttons(4).Visible = False
    End If
    Set lobjTmp = Nothing
    '2012-05-21
       
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
End Sub

'查询条件
Private Sub sub开始查询()
    Dim lobjRec As Object
    Dim sql As String
    Dim sql清空临时表 As String
    Dim sql查询结果 As String
    
    On Error GoTo errHandler
    
    '根据条件（起止时间，单位名称等）查找的信息到底是要只要检查过的（不管有没有检查完）还是只要检查完成的人员（上面这条是只要检查过的，下面那天是只要检查完成的） 2015-10-29
    sql = "select 系统编号,姓名,性别,年龄,危害因素,现工种,单位名称,体检日期 from 职业病体检_体检基本数据库 where 单位名称 = '" & Trim(ctxtCompanyName.Text) & "' and 体检状态 >=1 and convert(varchar(10),体检日期,120) between '" & Format(DTP开始时间.Value, "yyyy-mm-dd") & "' and '" & Format(DTP截止时间.Value, "yyyy-mm-dd") & "'"
'    sql = "select 系统编号,姓名,性别,年龄,危害因素,职业分类,单位名称,体检日期,复查状态,复查原因,复查项目 from 职业病体检_体检基本数据库 where 单位名称 = '" & Trim(ctxtCompanyName.Text) & "' and 体检状态 in (6,7) and convert(varchar(10),体检日期,120) between '" & Format(DTP开始时间.Value, "yyyy-mm-dd") & "' and '" & Format(DTP截止时间.Value, "yyyy-mm-dd") & "'"
    Set lobjRec = dafuncGetData(sql)
    cgrdList.Clear
    Set cgrdList.DataSource = lobjRec
    Set mobjRec = lobjRec
    
'    sql = "select 系统编号,姓名,性别,年龄,危害因素,职业危害工龄,体检结论,诊断和处理意见,既往病史,体检日期 from 职业病体检_体检报告视图 where"
'
'    sql = sql & " 单位名称 = '" & Trim(ctxtCompanyName.Text) & "' "
'
'    If flag日期.Value = 1 Then
'        sql = sql & " and (体检日期 >= '" & DTP开始时间.Value & "' and 体检日期 <= '" & DTP截止时间.Value & "')"
'    End If
'
'    Set lobjRec = lojb查询统计函数.func返回查询信息(sql)
'    Set mobjRec = lobjRec
'    sql查询结果 = sql
'    '显示之前，先清除表内已有的信息
'    cgrdList.Clear
'
'    Set cgrdList.DataSource = lobjRec
'    cgrdList.Editable = flexEDNone
'    cgrdList.AutoSize 1, cgrdList.Cols - 1, 0, 0
'    sql = ""
'
'    sql清空临时表 = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[职业病体检_单位报表打印表]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)drop table [dbo].[职业病体检_单位报表打印表]"
'    dafuncGetData (sql清空临时表)
'    dafuncGetData (sql查询结果)
    
    '2012-05-23 陶露↓
    'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
    cgrdList.AutoSize 0, cgrdList.cols - 1, 0, 0
    cgrdList.ExplorerBar = flexExSort
    cgrdList.DataMode = flexDMFree
   '2012-05-23↑
   
   
    Lab记录数量.Caption = "以上共有记录：" & lobjRec.RecordCount & "条"   '在标签中显示记录数目  2015-11-3
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetMedicalExamTemplate", "Timer1_Timer", 66666, lstrError, False
    MousePointer = 0
    '恢复界面可以操作。
    Me.Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ctlb工具栏.Width = Me.ScaleWidth - ctlb工具栏.Left * 2
    Frame1.Width = Me.ScaleWidth - ctlb工具栏.Left * 2
'    Frame1.Height = Me.ScaleHeight - Frame1.Top - 20
'    Frame1.Height = Me.ScaleHeight - 50
    cgrdList.Width = Frame1.Width - cgrdList.Left * 2
    cgrdList.Height = Frame1.Height - cgrdList.Top - 20

End Sub

'界面工具栏按钮操作设定
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = True
    
    Dim lcolID As New Collection
    Dim lobj体检类型 As Object
    Dim i As Integer
    Set lobj体检类型 = CreateObject("职业病对象.clsMedicalExam")
    lobj体检类型.系统编号名称 = pstrPerson
    
    Select Case Operate
        Case "预览报告"
            
            If cgrdList.rows <= 1 Then
                MsgBox "不能打印空总检报告！！！"
                Set lobj体检类型 = Nothing
                Exit Sub
            End If
            
            Dim lobjRec As Object, ltempRec As Object
            Dim lcolInfo As Collection, lcolFactor As Collection, lcolInfo2 As Collection, lcolItem As Collection, lcol As Collection, lcol2 As Collection
            Dim lstr As String, ltemp As String
            Dim lint As Integer

            lstr = "select 危害因素,count(*) as 人数 from 职业病体检_体检基本数据库 where 1=1 and 危害因素 <> '' "
            If flag名称.Value = 1 Then
                lstr = lstr & " and 单位名称 = '" & Trim(ctxtCompanyName.Text) & "'"
            End If
            If flag日期.Value = 1 Then
                lstr = lstr & " and (convert(varchar(10),体检日期,120) >= '" & Format(DTP开始时间.Value, "yyyy-mm-dd") & "' and convert(varchar(10),体检日期,120) <= '" & Format(DTP截止时间.Value, "yyyy-mm-dd") & "')"
            End If
            lstr = lstr & " group by 危害因素"
            Set lobjRec = dafuncGetData(lstr)
'            If Not (lobjRec.BOF Or lobjRec.EOF) Then
'                MsgBox ""
'            End If
            Set lcol = New Collection
            Set lcol2 = New Collection
            Set lcolInfo = New Collection
            Set lcolInfo2 = New Collection
            Set lcolItem = New Collection
            Set lcolFactor = New Collection
            lint = 0
            
            Dim temp1 As String, temp2 As String
            '第一部分
            '增加应检查人数（应检人数是登记了的人数）  2015-10-29↓
            Dim SlobjRec As Object
            Dim Slstr As String
            Dim sum人数 As Integer
            Dim mstr开始日期 As String
            Dim mstr截止日期 As String
             Slstr = "select count(*) as 人数 from 职业病体检_体检基本数据库 where 1=1 and 危害因素 <> '' "
            If flag名称.Value = 1 Then
                Slstr = Slstr & " and 单位名称 = '" & Trim(ctxtCompanyName.Text) & "'"
            End If
            If flag日期.Value = 1 Then

                '将体检登记查询的开始时间提前一个月  2015-11-3
                 mstr开始日期 = Format(DateAdd("d", -30, DTP开始时间), "yyyy-mm-dd")

                Slstr = Slstr & " and (convert(varchar(10),建档日期,120) >= '" & mstr开始日期 & "' and convert(varchar(10),建档日期,120) <= '" & Format(DTP截止时间.Value, "yyyy-mm-dd") & "')"
'                Slstr = Slstr & " and (convert(varchar(10),建档日期,120) >= '" & Format(DTP开始时间.Value, "yyyy-mm-dd") & "' and convert(varchar(10),建档日期,120) <= '" & Format(DTP截止时间.Value, "yyyy-mm-dd") & "')"
            End If
            Set SlobjRec = dafuncGetData(Slstr)
            sum人数 = SlobjRec("人数")
            '2015-10-29↑
            
            For i = 1 To lobjRec.RecordCount
                temp1 = lobjRec("危害因素")
                temp2 = lobjRec("人数")
                lcolFactor.Add temp1, "危害因素" & i
                lcolFactor.Add temp2, "人数" & i
'                lcolInfo.Add lobjRec("危害因素")
'                lcolInfo2.Add lobjRec("人数")
                lint = lint + Val(lobjRec("人数"))
                lobjRec.MoveNext
                temp1 = ""
                temp2 = ""
            Next
            lcolFactor.Add Trim(ctxtCompanyName.Text), "单位名称"
            lcolFactor.Add Format(DTP开始时间.Value, "yyyy年mm月dd日"), "体检日期"
            lcolFactor.Add Str(sum人数), "应检人数"   '增加应检人数  2015-10-29
            lcolFactor.Add Str(lint), "实检人数"
            lcolFactor.Add DTP开始时间.Value, "开始日期"
            lcolFactor.Add DTP截止时间.Value, "截止日期"
'            lcolFactor.Add Format(DTP开始时间.Value, yyyy - mm - dd), "开始日期"
'            lcolFactor.Add Format(DTP截止时间.Value, yyyy - mm - dd), "截止日期"

'            Set lobjRec = Nothing
'            lstr = "select distinct b.名称 as 科室名称 from 职业病体检_科室结论表 a,系统管理_字典_字典内容表 b where a.科室 = b.编号" _
'            & " and b.ID = (select ID from 系统管理_字典_字典表列表 where 名称 = '职业病体检科室字典') and b.名称 <> '最终结论录入' " _
'            & " and 系统编号 in (select 系统编号 from 职业病体检_体检基本数据库 where 单位名称 = '" & Trim(ctxtCompanyName.Text) & "'" _
'            & " and (convert(varchar(10),体检日期,120) >= '" & Format(DTP开始时间.Value, "yyyy-mm-dd") & "' and convert(varchar(10),体检日期,120) <= '" & Format(DTP截止时间.Value, "yyy-mm-dd") & "') and 体检状态 = 7)"
            
            Set ltempRec = dafuncGetData("select distinct b.名称 as 科室名称 from 职业病体检_科室结论表 a,系统管理_字典_字典内容表 b where a.科室 = b.编号" _
            & " and b.ID = (select ID from 系统管理_字典_字典表列表 where 名称 = '职业病体检科室字典') and b.名称 <> '最终结论录入' " _
            & " and 系统编号 in (select 系统编号 from 职业病体检_体检基本数据库 where 单位名称 = '" & Trim(ctxtCompanyName.Text) & "'" _
            & " and (convert(varchar(10),体检日期,120) >= '" & Format(DTP开始时间.Value, "yyyy-mm-dd") & "' and convert(varchar(10),体检日期,120) <= '" & Format(DTP截止时间.Value, "yyyy-mm-dd") & "') and 体检状态 = 7)")
            
            While Not (ltempRec.EOF Or ltempRec.BOF)
                ltemp = ltempRec("科室名称")
                lcolItem.Add Left(ltemp, Len(ltemp) - 1)
                ltempRec.MoveNext
            Wend
'            Set lobjRec = Nothing
            lcolFactor.Add lcolItem, "体检项目"
'            lcolFactor.Add lcolInfo, "危害因素"
'            lcolFactor.Add lcolInfo2, "危害结果"
            
'            lstr = "select b.名称,count(*) 人数 from dbo.职业病体检_体检结果视图 a,职业病体检_体检项目设置表 b where a.体检项目=b.编码 and 系统编号 in(" _
'                & " select 系统编号 from 职业病体检_体检基本数据库 where 危害因素 = 'Xray'and 单位名称 = '" & Trim(ctxtCompanyName.Text) & "'" _
'                & " and (体检日期 >= '" & DTP开始时间.Value & "' and 体检日期 <= '" & DTP截止时间.Value & "') and 体检状态 = 7" _
'                & ") and 单项结论 = '不合格' group by b.名称"
'
'            Set lobjRec = dafuncGetData(lstr)
'
'            While Not lobjRec.EOF
'                lcol.Add lobjRec("名称")
'                lcol2.Add lobjRec("人数")
'                lobjRec.nextmove
'            Wend

'            '体检人员表查询  2015-10-30
'        lstr = "select * from 职业病体检_体检表模板体检项目表(体检表名称,体检项目)values('" & mstr体检表名 & "','" & lobjItem.编码 & "') "
'        dafuncGetData (lstr)
 
            
            sub编辑总检报告 lcolFactor
'            Set lobjRec = Nothing
'            Set ltempRec = Nothing
'            Set lcolFactor = Nothing
            
            '第二部分
            
'            If cgrdList.Rows >= 2 Then
'                Dim lobjRec As Object
'                Dim lcolCompany As Collection
'                Set lobjRec = dafuncGetData("select 档案编号,单位名称,地址 from 单位档案_单位定位查询视图 where 单位名称 = '" & Trim(ctxtCompanyName.Text) & "'")
'                If Not lobjRec.EOF Then
'                    Set lcolCompany = New Collection
'                    lcolCompany.Add lobjRec("档案编号"), "档案编号"
'                    lcolCompany.Add lobjRec("单位名称"), "单位名称"
'                    lcolCompany.Add lobjRec("地址"), "单位地址"
'                End If
'                sub编辑单位报表 mobjRec, lcolCompany
'                lcolID.Add cgrdList.TextMatrix(1, 0)    '直接放入第一个系统编号
'                pobj业务对象.Sub打印单位文书 "职业病体检_单位报表", lcolID, False, True
'            End If
        Case "打印报告"
            If cgrdList.rows >= 2 Then
                lcolID.Add cgrdList.TextMatrix(1, 0)    '直接放入第一个系统编号
                pobj业务对象.Sub打印单位文书 "职业病体检_单位报表", lcolID, True, False
            End If
        Case "导出"
            If cgrdList.Row < 1 Then
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
                cgrdList.ColDataType(0) = flexDTString
                cgrdList.SaveGrid lstrFile, flexFileExcel, True   '导出excel系统编号为数字
                'cgrdMain.SaveGrid lstrFile, flexFileTabText, True
                '2012-04-14 于登淼↑
            End If
        Case "退出"
            Unload FrmQueryCompany
            Set FrmQueryCompany = Nothing
            Cancel = True
    End Select
    
    Set lobj体检类型 = Nothing
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "FrmQueryCompany", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

Public Function BigNum(num As Integer) As String
    Select Case num
        Case 0
            BigNum = "零"
        Case 1
            BigNum = "一"
        Case 2
            BigNum = "二"
        Case 3
            BigNum = "三"
        Case 4
            BigNum = "四"
        Case 5
            BigNum = "五"
        Case 6
            BigNum = "六"
        Case 7
            BigNum = "七"
        Case 8
            BigNum = "八"
        Case 9
            BigNum = "九"
        Case Else
            BigNum = "Err"
    End Select
End Function

Private Sub Timer1_Timer()
    Dim lojbRec As Object   '数据库结果对象
    Dim i As Integer
    On Error GoTo errHandler
    
    Timer1.Enabled = False
    '设置时间条件
    DTP开始时间.Value = DateAdd("M", -1, Now)
    DTP截止时间.Value = Now
    
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetMedicalExamTemplate", "Timer1_Timer", 6666, lstrError, False
    MousePointer = 0
    '恢复界面可以操作。
    Me.Enabled = True
End Sub
