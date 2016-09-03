VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmQueryStatis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "职业健康体检-查询统计"
   ClientHeight    =   10680
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "统计结果"
      ForeColor       =   &H000080FF&
      Height          =   2175
      Left            =   7440
      TabIndex        =   49
      Top             =   4080
      Width           =   5775
      Begin VSFlex8Ctl.VSFlexGrid cgrdStatic 
         Height          =   1095
         Left            =   0
         TabIndex        =   50
         Top             =   240
         Width           =   5775
         _cx             =   2088773578
         _cy             =   2088765323
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
         Cols            =   1
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
   End
   Begin VB.CommandButton CmdAction 
      BackColor       =   &H00C0FFFF&
      Caption         =   "查询结果"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton ccmdStatic 
      BackColor       =   &H0080FF80&
      Caption         =   "统计"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "图表 "
      ForeColor       =   &H000080FF&
      Height          =   4215
      Left            =   7440
      TabIndex        =   44
      Top             =   6360
      Width           =   5775
      Begin VB.PictureBox picChart 
         AutoSize        =   -1  'True
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3795
         ScaleWidth      =   5475
         TabIndex        =   45
         Top             =   240
         Width           =   5535
         Begin VB.Image Image1 
            Height          =   3495
            Left            =   240
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5175
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "查询条件 "
      ForeColor       =   &H000080FF&
      Height          =   3135
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Width           =   6255
      Begin VB.TextBox ctxt单位名称 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1200
         TabIndex        =   31
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton ccmd单位定位 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox ccmb查询条件 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox ccmb查询条件 
         Height          =   300
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2280
         Width           =   3855
      End
      Begin VB.ComboBox ccmb查询条件 
         Height          =   300
         Index           =   2
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox ccmb体检分类 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   3855
      End
      Begin VB.OptionButton cop体检结论 
         Caption         =   "不合格"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   25
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton cop体检结论 
         Caption         =   "合格"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox ccmb体检表名 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1080
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTP开始 
         Height          =   300
         Left            =   1200
         TabIndex        =   32
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月"
         Format          =   59637763
         CurrentDate     =   40986
      End
      Begin MSComCtl2.DTPicker DTP截止 
         Height          =   300
         Left            =   3480
         TabIndex        =   33
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月"
         Format          =   59637763
         CurrentDate     =   40986
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "人数："
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "单位名称"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "危害因素"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "行业类别"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "现工种"
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   39
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label7 
         Caption         =   "体检类型"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "体检情况"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "到"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   36
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "登记日期从"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "体检表名"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame fraChartConfig 
      Caption         =   "统计图表设计"
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   2640
      Width           =   6015
      Begin VB.TextBox YAxisTitle 
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "FrmQueryStatis.frx":0000
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox XAxisTitle 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "FrmQueryStatis.frx":0008
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox ChartTitle 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "FrmQueryStatis.frx":0010
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox comb图表样式 
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Text            =   "图表样式"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox comb色彩样式 
         Height          =   300
         Left            =   2400
         TabIndex        =   17
         Text            =   "色彩样式"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox comb图表布局 
         Height          =   300
         Left            =   4200
         TabIndex        =   16
         Text            =   "图表布局"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraChartStatic 
      Caption         =   "统计设计"
      ForeColor       =   &H000080FF&
      Height          =   1695
      Left            =   6480
      TabIndex        =   3
      Top             =   840
      Width           =   6015
      Begin VB.PictureBox tmpPicture 
         Height          =   495
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox cchkRowColSwap 
         Caption         =   "行列交换"
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox combYAxis 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox combXAxis 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton ccmdExportCrt 
         Caption         =   "导出图表"
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton ccmdDrawCrt 
         Caption         =   "绘制图表"
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox CombInterval 
         Height          =   300
         Left            =   2160
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox ctxtInterval 
         Height          =   270
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox cchkInterval 
         BackColor       =   &H00FFC0FF&
         Caption         =   "每隔"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "纵轴"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "横轴"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frame4 
      Caption         =   "查询结果"
      ForeColor       =   &H000080FF&
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   7335
      Begin VSFlex8Ctl.VSFlexGrid cgrdInfo 
         Height          =   3495
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   6855
         _cx             =   2088775483
         _cy             =   2088769557
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
         AllowSelection  =   0   'False
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
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSComDlg.CommonDialog ccdg 
         Left            =   8160
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   2760
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "注：先查询结果，后统计。"
      Height          =   180
      Left            =   9840
      TabIndex        =   48
      Top             =   3600
      Width           =   2160
   End
End
Attribute VB_Name = "FrmQueryStatis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'窗体：职业病查询统计界面
'功能：对职业病体检信息的详细查询和统计
'作者：翁乔
'时间：2012-04-18
'备注：暂无
Option Explicit
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Public mblnInUse As Boolean
Dim lojb科室 As Collection '所有科室
Dim lobj查询统计函数 As Object    '查询统计函数
Dim mobj缓存 As Object      '缓存病历上级目录
Dim isHistory As Boolean    '是否为病历查询
'2012-04-19 于登淼 ↓
'添加统计部分excel相关变量
Private hasStatPerm As Boolean
Private initChart As Boolean
Private xlApp As Object     'Excel.Application
Private xlBook As Object    'Excel.Workbook
Private xlSheet As Object   'Excel.Worksheet
Private xlChart As Object   'Excel.Chart
'2012-04-19 于登淼 ↑

Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property

'Private Sub ccmdBack_Click()
'
'    Set cgrdInfo.DataSource = mobj缓存
'    ccmdBack.Visible = False
'    Set mobj缓存 = Nothing
'
'End Sub

'2012-05-20 于登淼
'添加计算统计函数，按照当前条件统计
Private Sub ccmdStatic_Click()
    Dim XSelected As Integer
    Dim YSelected As Integer
    '修改人：张令 2012.12.04
    'bug号：0000055
    '说明：起始时间不能在结束时间之后，给予提示。↓↓
    If DTP开始.Value > DTP截止.Value Then
        MsgBox "起始日期不能在结束日期之后！"
        DTP开始.Value = DTP截止.Value
        Exit Sub
    End If
    '2012.12.04     ↑↑
    If cgrdInfo.rows = 1 Then Exit Sub
    If isHistory Then Exit Sub
    
    ccmdStatic.Caption = "统计中..."
    
    XSelected = combXAxis.ItemData(combXAxis.ListIndex)
    YSelected = combYAxis.ItemData(combYAxis.ListIndex)
    
    '2012-05-29 于登淼 ↓
    '修改统计子函数，每次传入统计分类和统计内容
    '初始化
'    SSTabGrid.Tab = 1
    Select Case XSelected
    Case 0  '按工种统计
        sub按工种统计 XSelected, YSelected
    Case 1  '按行业统计
        sub按行业统计 XSelected, YSelected
    Case 2  '按单位统计
        sub按单位统计 XSelected, YSelected
    Case 3  '按体检情况统计
        sub按体检情况统计 XSelected, YSelected
    '2012-08-08 于登淼 ↓
    Case 4  '按危害因素统计
        sub按危害因素统计 XSelected, YSelected
    '2012-08-08 于登淼 ↑
    '2013-03-31 刘云乐 ↓
    Case 5
        sub按时间人数统计 XSelected, YSelected
    '2013-03-31 刘云乐 ↓
    End Select
    '2012-05-29 于登淼 ↑
    
    ccmdStatic.Caption = "统计"
    
    '根据cgrdStatic里内容，话excel chart图表
    SelectData (cchkRowColSwap)
    DrawChart

End Sub

'Private Sub ccmd病历查询_Click()
'
'    Dim lobjRec As Object, lobjNo As Object
'    Dim str As String, lstrNo As String
'    Dim i As Integer
'    Dim date1 As String
'    Dim date2 As String
''    str = "select a.系统编号,b.名称,a.文字结论,a.医生编号,a.结论日期 " _
''    & "from 职业病体检_科室结论表 a join 系统管理_字典_字典内容表 b on a.科室 = b.编号 and b.id = 84"
'    If Trim(ctxt病历(0).Text) = "" And Trim(ctxt病历(1).Text) = "" Then
'        MsgBox "请指定体检人员的姓名或系统编号！"
'        Exit Sub
'    End If
'
'    str = "select 系统编号 from 职业病体检_体检基本数据库 where 1 = 1 "
'
'    '修改人：张令 2012.12.18  ↓↓
'    '修改说明：去掉一个单引号。
'    If Trim(ctxt病历(0).Text) <> "" Then
''        str = str & " and 姓名=''" & Trim(ctxt病历(0).Text) & "''"
'        str = str & " and 姓名='" & Trim(ctxt病历(0).Text) & "'"
'    End If
'
'    If Trim(ctxt病历(1).Text) <> "" Then
'        '修改人：张令 2013.01.04   ↓↓
'        '修改说明：由于系统编号不唯一，所以按照系统编号查询会查询不到数据，
'        '          所以就用编号查询当前编号对应的姓名，然后根据姓名查询对应病史数据。
''        str = str & " and 系统编号 =''" & Trim(ctxt病历(1).Text) & "''"
''        str = str & " and 系统编号 ='" & Trim(ctxt病历(1).Text) & "'"
'        Set lobjNo = dafuncGetData("select 姓名 from 职业病体检_体检人员基本信息表 where 系统编号='" & Trim(ctxt病历(1).Text) & "'")
'        If Not (lobjNo.EOF Or lobjNo.BOF) Then
'            str = str & " and 姓名='" & lobjNo(0) & "'"
'        Else
'            MsgBox "姓名与系统编号不符，按姓名查询数据。"
'        End If
'        Set lobjNo = Nothing
'        '修改人：张令 2013.01.04   ↑↑
'    End If
'
'    If Trim(ccmb病历单位.Text) <> "" Then
''        str = str & " and 单位名称=''" & Trim(ccmb病历单位.Text) & "''"
'        str = str & " and 单位名称='" & Trim(ccmb病历单位.Text) & "'"
'    End If
'
'    If Trim(ccmb危害因素.Text) <> "所有" Then
''        str = str & " and 危害因素=''" & Trim(ccmb危害因素.Text) & "''"
'        str = str & " and 危害因素='" & Trim(ccmb危害因素.Text) & "'"
'    End If
'    '修改人：张令 2012.12.18  ↑↑
'    date1 = DTP病历begin.Value
'    date2 = DTP病历end.Value
'
'    '修改人：张令 2012.12.18  ↓↓
'    '修改说明：之前将select语句直接嵌套在exec中执行，语法有问题，提出来，先执行select语句后放值进入exec中执行。
''    str = str & "'"
'    '补充修改人：张令 2013.01.04  ↓↓
'    '修改说明：由于一个人可能有多次体检病史，所以查询出来后要将所有系统编号加如查询条件。
'    Set lobjNo = dafuncGetData(str)
'    lstrNo = ""
'    For i = 0 To lobjNo.RecordCount - 1
'        If Not lstrNo = "" Then
'            lstrNo = lstrNo & ",''" & lobjNo(0) & "''"
'            lobjNo.MoveNext
'        Else
'            lstrNo = "''" & lobjNo(0) & "''"
'            lobjNo.MoveNext
'        End If
'    Next
'    '修改人：张令 2013.01.04   ↑↑
'    If Not lstrNo = "" Then
''    Set lobjRec = dafuncGetData("exec 职业病体检_返回相关病历信息 " & str & ",'" & date1 & "','" & date2 & "'")
'        Set lobjRec = dafuncGetData("exec 职业病体检_返回相关病历信息 '" & lstrNo & "','" & date1 & "','" & date2 & "'")
'        '修改人：张令 2012.12.18  ↑↑
'        cgrdInfo.rows = 1
'        If Not (lobjRec.BOF Or lobjRec.EOF) Then
'            cgrdInfo.SelectionMode = 3
'            Set cgrdInfo.DataSource = lobjRec
'    '        With cgrdInfo
'    '            .Cols = .Cols + 1: .TextMatrix(0, .Cols - 1) = IIf()
'    '        End With
'            cgrdInfo.AutoResize = True
'            cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, , True
'            'cgrdInfo.AutoSizeMode = flexAutoSizeColWidth
'            cgrdInfo.AllowSelection = False
'            isHistory = True
'            Set mobj缓存 = lobjRec
'        End If
'    Else
'        cgrdInfo.rows = 1
'    End If
'End Sub

'翁乔，2012-10-22
Private Sub ccmd单位定位_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '单位定位返回的结果记录。
    Dim lobj单位 As Object
    Dim lobj单位信息 As Object
    Dim mstr单位申请编号 As String
    '启动单位定位界面。
'    Set lobjRec = pobj业务对象.func单位定位        '原调用单位定位注释掉  2016-1-21 by 牟俊
    frmQueryCompanyLocation.Show 1, Me    '调用单位定位查询 2016-1-21 by 牟俊
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxt单位名称.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
        End If
    End If
    
    '把焦点回到单位录入框。保存能保存新单位定位信息。
    ctxt单位名称.SetFocus
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "查询统计界面", "ccmd单位定位_Click", 6666, lstrError, False
End Sub

Private Sub cgrdInfo_DblClick()
    Dim lobjRec As Object
    Dim tempDepart As String
    Dim tempNo As String
    Dim tempDNo As String
    Dim tempDate As String
    On Error GoTo errHandler
    
    If Not isHistory Then Exit Sub
    
    If Not mobj缓存 Is Nothing Then Exit Sub
    
    If cgrdInfo.Row > 0 Then
        tempNo = cgrdInfo.TextMatrix(cgrdInfo.Row, 0)
        tempDepart = cgrdInfo.TextMatrix(cgrdInfo.Row, 1)
        If Trim(tempDepart) = "最终结论录入" Then
            MsgBox "没有可查看体检项！！"
            Exit Sub
        End If
        tempDNo = cgrdInfo.TextMatrix(cgrdInfo.Row, 3)
        tempDate = cgrdInfo.TextMatrix(cgrdInfo.Row, 4)
        Set lobjRec = dafuncGetData("select * from 职业病体检_结果信息_" & tempDepart & " where " _
        & "系统编号='" & tempNo & "' and 体检医师='" & tempDNo & "' and convert(varchar(10),填写时间,120)='" & tempDate & "'")
        
        If Not lobjRec.EOF Then
'            ccmdBack.Visible = True
            Set mobj缓存 = cgrdInfo.DataSource
            Set cgrdInfo.DataSource = lobjRec
        End If
        
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

'查询统计；翁乔；2012-10-22
Private Sub CmdAction_Click()
    Dim i As Integer
    Dim sql As String
    Dim lobjRec As Object
    On Error GoTo errHandler

    sql = "select * from 职业病体检_查询统计视图 where 1=1"
    For i = 0 To ccmb查询条件.Count - 1
        If ccmb查询条件(i).Text <> "所有" Then
            sql = sql & " and " & Label2(i).Caption & " = '" & ccmb查询条件(i).Text & "'"
        End If
    Next
    
    If Trim(ctxt单位名称.Text) <> "" Then
        sql = sql & " and 单位名称 = '" & Trim(ctxt单位名称.Text) & "'"
    End If
    
    If ccmb体检表名.Text <> "所有" Then
        sql = sql & " and 体检表 = '" & ccmb体检表名.Text & "'"
    End If
    '修改人：张令 2012.12.18
    '修改说明：表列名弄错了，应该是“体检人员类型”↓↓
    If ccmb体检分类.Text <> "所有" Then
        sql = sql & " and 体检人员类型 = '" & ccmb体检分类.Text & "'"
    End If
    '修改人：张令 2012.12.18  ↑↑
    Dim dtpTimeTo As Date
    
    dtpTimeTo = Format(DateAdd("m", 1, DTP截止.Value), "yyyy-mm-01")
    dtpTimeTo = Format(DateAdd("d", -1, dtpTimeTo), "yyyy-mm-dd")
    sql = sql & " and (体检日期 between '" & Format(DTP开始.Value, "yyyy-mm" & "-01 00:00:00") & "' and '" & Format(dtpTimeTo, "yyyy-mm-dd" & " 23:59:59") & "')"
    '2013-03-04 刘云乐
    '不需要合格与不合格
'    If cop体检结论(0) Then
'        sql = sql & " and 体检状态 = '已发报告'"
'    Else
'        sql = sql & " and 体检状态 = '待复查'"
'    End If
    sql = sql & " and 体检状态 in('已发报告','待复查','已复核')"
    '2013-03-04 刘云乐
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(sql)
'    If SSTabGrid.TabIndex = 1 Then
'        SSTabGrid.TabIndex = 0
'    End If
    cgrdInfo.rows = 1
    Label8.Caption = "人数：" & cgrdInfo.rows - 1
    If Not (lobjRec.EOF Or lobjRec.BOF) Then
        Set cgrdInfo.DataSource = lobjRec
        cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
        isHistory = False
        
    Else
        MsgBox "没有符合查询条件的结果！", vbInformation, "系统提示"
    End If
    Label8.Caption = "人数：" & cgrdInfo.rows - 1  '统计人数 2016-1-20 by 牟俊
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume

End Sub

'2012-05-30 于登淼
'当x轴为体检状况时，处理的与上面三个稍微不一样。
'此时禁用其它选项“合格人数”、“不合格人数”、“合格率”、“无结果人数”、“金额”
Private Sub combXAxis_Click()
    If combXAxis.ListIndex = 3 Then
        combYAxis.Enabled = False
    Else
        'combYAxis.Enabled = True
    End If
End Sub

'2012-05-20 于登淼 判断时间间隔数字格式
Private Sub ctxtInterval_LostFocus()
    If ctxtInterval.Text = "" And cchkInterval.Value = 1 Then MsgBox ("请输入内容"): Exit Sub
    If IsNumeric(ctxtInterval.Text) = False Then MsgBox ("请输入数字"): Exit Sub
End Sub

Private Sub DTP截止_Change()
    If (DTP截止.Value - DTP开始.Value) / 30 < 1 Then
        MsgBox "日期在一个月以上。"
        DTP开始.Value = DateAdd("m", -1, Format(DTP截止.Value, "yyyy/MM"))
'        Exit Sub
    End If
    DTP开始.Value = Format(DTP开始.Value, "yyyy/MM")
'    DTP截止.Value = Format(DTP截止.Value, "yyyy/MM")
    DTP截止.Value = DateAdd("d", -1, Format(DateAdd("M", 1, Format(DTP截止.Value, "yyyy/MM")), "yyyy/MM"))
End Sub

Private Sub DTP开始_Change()
    If (DTP截止.Value - DTP开始.Value) / 30 < 1 Then
        MsgBox "日期在一个月以上。"
        DTP开始.Value = DateAdd("m", -1, Format(DTP截止.Value, "yyyy/MM"))
'        Exit Sub
    End If
'    DTP开始.Value = Format(DTP开始.Value, "yyyy/MM")
'    DTP截止.Value = Format(DTP截止.Value, "yyyy/MM")
'    DTP截止.Value = DateAdd("d", -1, Format(DateAdd("M", 1, Format(DTP截止.Value, "yyyy/MM")), "yyyy/MM"))
End Sub

Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    Dim i As Integer
    On Error GoTo errHandler
    Dim lojbRec As Object   '数据库结果对象

    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    '设置工具栏上所需要的各种按钮。
    '修改：2002-7-1（杨春）简化取消结论的操作。该为操作单选框。
    With lcol工具栏按钮
        '2012-04-19 于登淼 ↓
        '修改内容：只支持导入导出格式为excel格式。
        .Add "导出Excel(&O)113"
        .Add "|"
        .Add "导出图表(&T)102"
        .Add "预览报告(&L)108"
        .Add "打印报告(&P)107"
        '2012-04-19 于登淼 ↑
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
'    Dim lobjTmp As Object
'    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
'    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_职业病查询统计_导出") = False Then
'        ctlb工具栏.Buttons(1).Visible = False
'    End If
'    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_职业病查询统计_导入") = False Then
'        ctlb工具栏.Buttons(2).Visible = False
'    End If
'    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_职业病查询统计_打印") = False Then
'        ctlb工具栏.Buttons(3).Visible = False
'        ctlb工具栏.Buttons(4).Visible = False
'    End If
'    Set lobjTmp = Nothing
    '2012-05-23 ↑↑↑
    
    '查询条件初始化
    '创建查询统计函数对象
    Set lobj查询统计函数 = CreateObject("职业病界面.clsQueryStatis")
    
    '2012-04-19 于登淼 ↓
    '应该在拥有权限的前提下，进行统计的操作
    'if 有统计的权限=true then
    hasStatPerm = True
'    Else
'        hasStatPerm = False
    'end if
    '打开一个临时的excel文件
    OpenTempExcel
    
    '初始化统计部分下拉列表
    subInitChartList
    subInitStaticList
    
    '2012-04-19 于登淼 ↑
    
    '2012-04-23 于登淼 ↓
    '设置cgrdInfo的排序等
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    '2012-04-23 于登淼 ↑
    
    '2012-05-20 于登淼 ↓
    '添加统计部分控件初始化信息
'    DTP开始.Value = Format(DateAdd("M", -1, Date), "yyyy/MM")
'    DTP截止.Value = DateAdd("d", -1, Format(DateAdd("M", 1, Date), "yyyy/MM"))
'    chkStart.Value = 1
'    chkEnd.Value = 1
'    DTP病历begin.Value = DateAdd("d", -30, Date)
'    DTP病历end.Value = Date
    
    cchkInterval.Value = 1
    ctxtInterval.Text = "1"
    With CombInterval   '没有对月份、季度、年份做详细判断
        .Clear
'        .AddItem "日": .ItemData(.NewIndex) = 1
        .AddItem "月": .ItemData(.NewIndex) = 1
        .AddItem "季": .ItemData(.NewIndex) = 3
        .AddItem "年": .ItemData(.NewIndex) = 12
        .ListIndex = 0
    End With
'    SSTab查询统计结果.TabIndex = 0
'    SSTabGrid.TabIndex = 0
    '2012-05-20 于登淼 ↑
    '没有绘制图表则不能打印报告和导出excel
    If picChart.Picture = 0 Then
'        ctlb工具栏.Buttons(4).Enabled = False
'        ctlb工具栏.Buttons(5).Enabled = False
'        ctlb工具栏.Buttons(4).Enabled = False
    End If
    
    '表格初始化；翁乔；2012；
    cgrdInfo.cols = 0
    With cgrdInfo
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "系统编号"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "收费批号"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "姓名"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "性别"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检表"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检类型"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检日期"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "工种"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "单位名称"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "危害因素"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "卫生种类"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "行业类别"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "复查系统编号"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检结论"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "下结论日期"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "医师姓名"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "收费金额"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检状态"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    If picChart.Picture = 0 Then
        ccmdExportCrt.Enabled = False
    End If
    cop体检结论(0).Value = True
    '修改人：张令 2012.12.04     ↓↓
    'bug号：0000054
    '说明：不用预览和打印报告。
'    ctlb工具栏.Buttons(2).Visible = False
    ctlb工具栏.Buttons(5).Visible = False
    ctlb工具栏.Buttons(4).Visible = False
    '2012.12.04     ↑↑
    '设置时间控件执行余下代码
    Timer1.Enabled = True
    '2012-05-21 陶露
    '界面权限设置
'    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
'    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_查询统计_导出") = False Then
'        ctlb工具栏.Buttons(1).Visible = False
'    End If
'    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_查询统计_导入") = False Then
'        ctlb工具栏.Buttons(2).Visible = False
'    End If
'    Set lobjTmp = Nothing
    '2012-05-21
    
    'MsgBox (DTP开始.Value & " || " & Format(DTP开始.Value, "yyyy-mm-dd") & " " & Format("00:00:00", "hh:mm:ss")) ''''''''''test
    'MsgBox (DTP截止.Value & " || " & Format(DTP截止.Value, "yyyy-mm-dd") & " " & Format("00:00:00", "hh:mm:ss")) ''''''''''test
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    mblnInUse = False
    
    '2012-05-23 于登淼 ↓
    '退出窗体时，强行关闭进程
    If hasStatPerm = True Then CloseTempExcel
    '2012-05-23 于登淼 ↑
    
    Set mobjGUI = Nothing
    Exit Sub

    '2012-05-24 于登淼 ↓
errHandler:
    mblnInUse = False
    Set xlChart = Nothing
    Set xlSheet = Nothing
    xlApp.Workbooks.Close
    xlApp.Quit
    Set xlBook = Nothing
    If Not xlApp Is Nothing Then
        Shell "cmd.exe /c taskkill /f /im excel.exe"
    End If
    Set xlApp = Nothing
    Set mobjGUI = Nothing
    '2012-05-24 于登淼 ↑
End Sub

Private Sub ccmdDrawCrt_Click()
    If cgrdInfo.rows = 1 Then Exit Sub
    If isHistory Then Exit Sub
    SelectData (cchkRowColSwap.Value)
    DrawChart
End Sub

Private Sub cchkRowColSwap_Click()
'    SelectData (cchkRowColSwap.Value)
    selectdataRC (cchkRowColSwap.Value)  '2016-1-27 by 牟俊
    DrawChart
End Sub

Private Sub ccmdExportCrt_Click()
    '
    Dim lstrFile As String
    ccdg.FileName = ""
    ccdg.Filter = "JPEG(*.jpg)|*.jpg|" & _
                "97-03 Excel(*.xls)|*.xls|" & _
                "07 Excel(*.xlsx)|*.xlsx"
    ccdg.ShowSave
    lstrFile = ccdg.FileName
    If lstrFile <> "" Then
        If ccdg.FilterIndex = 1 Then        '将chart存为jpg
            VB.SavePicture picChart.Picture, lstrFile
        ElseIf ccdg.FilterIndex = 2 Then    '将chart与数据存为xls或xlsx(仅支持当前版本吧)
            Select Case xlApp.Application.Version
            Case "12.0"
                xlBook.SaveCopyAs (Replace(lstrFile, ".xls", ".xlsx"))
            Case Else
                xlBook.SaveCopyAs (lstrFile)
            End Select
        ElseIf ccdg.FilterIndex = 3 Then
            Select Case xlApp.Application.Version
            Case "12.0"
                xlBook.SaveCopyAs (lstrFile)
            Case Else
                xlBook.SaveCopyAs (Replace(lstrFile, ".xlsx", ".xls"))
            End Select
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    Picture1.Left = 0
'    Picture1.Width = Me.ScaleWidth - Picture1.Left
'    Picture1.Height = Me.ScaleHeight - Picture1.Top
'    Frame1.Width = Picture1.Width - Frame1.Left
'    Frame1.Height = Picture1.Height - Frame1.Top
'    ctlb工具栏.Width = Frame1.Width - ctlb工具栏.Left
'
''    SSTab查询统计结果.Width = Frame1.Width - SSTab查询统计结果.Left - 60
''    SSTab查询统计结果.Height = Frame1.Height - SSTab查询统计结果.Top - 120
''    Frame4.Width = SSTab查询统计结果.Width - cgrdInfo.Left - 60
'    frame4.Height = Frame1.Height - frame4.Top - 120
'    cgrdInfo.Width = frame4.Width - cgrdInfo.Left - 60
'    cgrdInfo.Height = frame4.Height - cgrdInfo.Top - 120
''    Frame6.Width = SSTab查询统计结果.Width - Frame6.Left - 60
''    Frame6.Height = SSTab查询统计结果.Height - Frame6.Top - 120
'    cgrdStatic.Width = Frame6.Width - cgrdStatic.Left - 60
'    cgrdStatic.Height = Frame6.Height - cgrdStatic.Top - 120
    ctlb工具栏.Width = Me.ScaleWidth
    Frame3.Width = Me.ScaleWidth * 2 / 5
    fraChartStatic.Left = Frame3.Left + Frame3.Width + 20
    fraChartStatic.Width = Me.ScaleWidth - fraChartStatic.Left - Frame3.Left
    fraChartConfig.Left = fraChartStatic.Left
    fraChartConfig.Width = fraChartStatic.Width
    Frame4.Left = Frame3.Left
    Frame1.Left = fraChartStatic.Left
    Frame4.Width = Frame3.Width
    Frame1.Width = fraChartStatic.Width
    Frame4.Height = Me.ScaleHeight - Frame4.Top - 20
    Frame1.Height = Frame4.Height
    cgrdInfo.Width = Frame4.Width - cgrdInfo.Left * 2
    cgrdInfo.Height = Frame4.Height - cgrdInfo.Top - 10
    picChart.Width = Frame1.Width - picChart.Left * 2
    picChart.Height = Frame1.Height - picChart.Top - 10
    Image1.Width = FrmQueryStatis.Width - Frame4.Width - 700 '让image的大小合适
'    Image1.Width = 8000
'    Image1.Height = Frame4.Height - Frame2.Height - 1000
    CmdAction.Left = Frame3.Left + Frame3.Width + 100
    ccmdStatic.Left = CmdAction.Left + CmdAction.Width + 100
    Label9.Left = ccmdStatic.Left + ccmdStatic.Width + 100
    Frame2.Left = Frame1.Left
    Frame2.Width = Frame1.Width
    cgrdStatic.Height = Frame2.Height - cgrdStatic.Top - 10
    cgrdStatic.Width = Frame2.Width - cgrdStatic.Left * 2
    
End Sub



Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Dim para预览报告 As Boolean
    Dim para打印报告 As Boolean
    Cancel = True
    
    Select Case Operate
    '2012-04-19 于登淼 ↓
    '添加excel导入vsflexgrid，vsflexgrid导出excel中
    Case "导入Excel"
        '这个地方可以有多个判断，暂时只做了Excel的。
        '配合后面的读入数据部分，可以判断不同文件类型的输入。
        ccdg.Filter = "Excel file" & "(*.xls)|*.xls" & _
                    "|Batch Files (*.bat)|*.bat|" & _
                    "All Files (*.*)|*.*"
        ccdg.FileName = ""
        ccdg.ShowOpen
        '2012-05-20 于登淼 ↓
'        SSTab查询统计结果.Tab = 1
'        SSTabGrid.Tab = 1
        'subInitChartList
        If ccdg.FileName = "" Then Exit Sub
        '2012-05-20 于登淼 ↑
        sub显示导入信息
        If hasStatPerm = True Then SelectData (cchkRowColSwap): DrawChart
    Case "导出图表"
        ccmdExportCrt_Click
    Case "导出Excel"
        '注意会不会出现ccdg未定义问题。
        Dim lstrFile As String
        ccdg.Filter = "Excel文件 (*.xls)|*.xls" & "|Excel 2007 files (*.xlsx)|*.xlsx"
        ccdg.ShowSave
        lstrFile = ccdg.FileName 'Replace(ccdg.FileName, ".xls", "") & "_" & Date & ".xls"
        If lstrFile <> "" Then
            'cgrdMain.ColDataType(0) = flexDTString '保存时限制列的格式为字符串 flexFileExcel
            xlSheet.SaveAs lstrFile
'            cgrdStatic.SaveGrid lstrFile, flexFileData, True 'true时，导出表头；false不导出表头
        End If
    '2012-04-19 于登淼 ↑
    '2012-06-06 于登淼 ↓
    '预览、打印统计图表报告，水晶报表格式
    Case "预览报告"
        para预览报告 = True
        para打印报告 = False
        sub打印统计报告 para预览报告, para打印报告
    Case "打印报告"
        para预览报告 = False
        para打印报告 = True
        sub打印统计报告 para预览报告, para打印报告
    '2012-06-06 于登淼 ↑
    Case "退出"
        '2012-04-19 于登淼 ↓
'        If hasStatPerm = True Then
'            CloseTempExcel
'        End If
        '2012-04-19 于登淼 ↑
        Set mobj缓存 = Nothing
        Unload Me
    End Select
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmRegisterManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub


'修改：初始化查询条件；翁乔；2012-10-22
Private Sub Timer1_Timer()
    Dim lobjRec As Object   '数据库结果对象
    Dim i As Integer
    On Error GoTo errHandler
    
    Timer1.Enabled = False

    '设置时间条件
    DTP开始.Value = DateAdd("M", -1, Now)
    DTP截止.Value = Now
    
    '读取行业类别
    Set lobjRec = lobj查询统计函数.func读取行业类别
    If lobjRec.RecordCount > 0 Then
        ccmb查询条件(0).AddItem "所有"
        For i = 1 To lobjRec.RecordCount
            ccmb查询条件(0).AddItem lobjRec("名称")
            lobjRec.MoveNext
        Next
    End If
    
    '读取危害因素
    Set lobjRec = lobj查询统计函数.func危害因素
    If lobjRec.RecordCount > 0 Then
        ccmb查询条件(1).AddItem "所有"
'        ccmb危害因素.AddItem "所有"
        For i = 1 To lobjRec.RecordCount
            ccmb查询条件(1).AddItem lobjRec("名称")
'            ccmb危害因素.AddItem lobjRec("名称")
            lobjRec.MoveNext
        Next
    End If
    
    '读取工种
    Set lobjRec = lobj查询统计函数.func读取工种
    If lobjRec.RecordCount > 0 Then
        ccmb查询条件(2).AddItem "所有"
        For i = 1 To lobjRec.RecordCount
            ccmb查询条件(2).AddItem lobjRec("名称")
            lobjRec.MoveNext
        Next
    End If
    
'    ccmb危害因素.ListIndex = 0
    ccmb查询条件(0).ListIndex = 0
    ccmb查询条件(1).ListIndex = 0
    ccmb查询条件(2).ListIndex = 0
    
    '体检分类读取
    Set lobjRec = lobj查询统计函数.func读取体检类型
    If lobjRec.RecordCount > 0 Then
        ccmb体检分类.AddItem "所有"
        For i = 1 To lobjRec.RecordCount
            ccmb体检分类.AddItem lobjRec("名称")
            lobjRec.MoveNext
        Next i
    End If
    ccmb体检分类.Text = "所有"

    '读取体检表名称
    Set lobjRec = lobj查询统计函数.func读取体检表
    If lobjRec.RecordCount > 0 Then
        ccmb体检表名.AddItem "所有"
        For i = 1 To lobjRec.RecordCount
            ccmb体检表名.AddItem lobjRec("体检表名称")
            lobjRec.MoveNext
        Next i
    End If
    ccmb体检表名.Text = "所有"
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetMedicalExamTemplate", "Timer1_Timer", 6666, lstrError, False
    MousePointer = 0
    '恢复界面可以操作。
    Me.Enabled = True
End Sub

'2012-04-19 于登淼
'函数功能：导入excel文档时，显示其内容。
Sub sub显示导入信息()
    On Error GoTo errHandler
    
    With cgrdStatic
        '调整cgrdStatic表格格式
        .FixedRows = 0: .FixedCols = 0
        
        '先清除之前已有的信息
        .Clear
        
        '这个地方可以有多个判断，暂时只做了Excel的。
        '配合后面的读入数据部分，可以判断不同文件类型的输入。
        .LoadGrid ccdg.FileName, flexFileExcel
        .AutoSize 1, .cols - 1, 0, 0
        
        '调整cgrdStatic表格格式
        .FixedRows = 1: .FixedCols = 1
        
        '2012-05-23 陶露
        'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
        cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
        cgrdStatic.ExplorerBar = flexExSort
        cgrdStatic.DataMode = flexDMFree
        '2012-05-23
    End With

    Exit Sub
errHandler:
    '如果没有导入文件，则不提示任何信息。
    If ccdg.FileName = "" Then Exit Sub
    MsgBox ("导入Excel文件错误！")
End Sub


'2012-05-20 于登淼
'判断字符串中是否有“复检”，“复查”，“不合格”。若有，则不合格。
Private Function sub最终合格(ByVal paraCon As String) As Integer
    If paraCon = "" Then sub最终合格 = 1: Exit Function
    If InStr(paraCon, "复检") > 0 Or InStr(paraCon, "复查") > 0 Or InStr(paraCon, "不合格") > 0 Then
        sub最终合格 = 0
        Exit Function
    Else
        sub最终合格 = 1
        Exit Function
    End If
End Function

'2012-04-19 于登淼
Sub SelectData(ByVal ifSwap As Integer)
    ccmdStatic.Caption = "数据整理中..."
    
'    xlSheet.Activate
    xlSheet.rows.Clear
    Dim i, j As Integer
    For i = 0 To cgrdStatic.rows - 1
         For j = 0 To cgrdStatic.cols - 1
            If ifSwap = 0 Then
                xlSheet.Cells(i + 1, j + 1) = cgrdStatic.TextMatrix(i, j)
            Else
                xlSheet.Cells(j + 1, i + 1) = cgrdStatic.TextMatrix(i, j)
            End If
        Next j
    Next i
    xlSheet.Activate  '先选择单元格前要激活所在工作表（不然不退出程序第二次查询统计时会出错）  2016-1-28 by 牟俊
    xlSheet.Cells.Select
'    xlSheet.Shapes.addchart.Select
    xlSheet.Shapes.SelectAll    '2015-12-7 by 牟俊
'    Set xlChart = xlApp.ActiveChart   '第二次及之后使用时，会报错“远程服务器不存在"
    Set xlChart = xlBook.Charts.Add   '2016-1-27 by 牟俊
    initChart = True
    
    ccmdStatic.Caption = "统计"
End Sub

'行列交换后走这个，不走上面的selectdata 2016-1-27 by 牟俊 ↓
Sub selectdataRC(ByVal ifSwap As Integer)
    ccmdStatic.Caption = "数据整理中..."
    
'    xlSheet.Activate
    xlSheet.rows.Clear
    Dim i, j As Integer
    For i = 0 To cgrdStatic.rows - 1
         For j = 0 To cgrdStatic.cols - 1
            If ifSwap = 0 Then
                xlSheet.Cells(i + 1, j + 1) = cgrdStatic.TextMatrix(i, j)
            Else
                xlSheet.Cells(j + 1, i + 1) = cgrdStatic.TextMatrix(i, j)
            End If
        Next j
    Next i
'    xlSheet.Cells.Select
    xlSheet.Shapes.SelectAll
    Set xlChart = xlBook.Charts.Add
    initChart = True
    
    ccmdStatic.Caption = "统计"
End Sub
'2016-1-27 by 牟俊 ↑

'2012-04-19 于登淼
'打开一个临时的excel文件，方便统计时进行操作。
Sub OpenTempExcel()
    Set xlApp = CreateObject("excel.application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlSheet.Activate
End Sub

'2012-04-19 于登淼
'关闭统计过程中用到的excel文件，并释放内存。
'如果不关闭文件，会一直在内存中占用资源
Sub CloseTempExcel()
    On Error GoTo errHandler

    '2012-05-22 于登淼 ↓
    '如果不往临时excel写入内容，这个方法可以在退出窗体时结束进程。
    '如果临时excel里面写入了内容，在当前窗体退出时，进程不能结束。
    '只能当整个系统退出时，VB才关闭进程。
    Set xlChart = Nothing
    Set xlSheet = Nothing
    xlApp.DisplayAlerts = False
    xlBook.Save
    xlBook.Close (True)
    xlApp.Workbooks.Close
    xlApp.Quit
    Set xlBook = Nothing
    '    2012-05-23 于登淼 ↓
    '    强制关闭一个excel进程，至于关闭的哪个，要等会儿测试。
    '    但中间出错后无法执行到这一步，仍然无法关闭excel进程。
'    If Not xlApp Is Nothing Then
'        Shell "cmd.exe /c taskkill /f /im excel.exe"
'    End If
    '    2012-05-23 于登淼 ↑
    Set xlApp = Nothing
    '2012-05-22 于登淼 ↑
    
    Exit Sub
errHandler:
'    xlApp.Quit
'    Set xlApp = Nothing
End Sub

'2012-04-19 于登淼 ↓
'每次选择画图参数后，均要重新画表
Private Sub comb色彩样式_Click()
    If initChart = True Then DrawChart
End Sub

Private Sub comb图表布局_Click()
    If initChart = True Then DrawChart
End Sub

Private Sub comb图表样式_Click()
    If initChart = True Then DrawChart
End Sub

Private Sub XAxesTitle_KeyPress(KeyAscii As Integer)
    If initChart = True And KeyAscii = 13 Then DrawChart
End Sub

Private Sub XAxesTitle_LostFocus()
    If initChart = True Then DrawChart
End Sub

Private Sub YAxesTitle_KeyPress(KeyAscii As Integer)
    If initChart = True And KeyAscii = 13 Then DrawChart
End Sub

Private Sub YAxesTitle_LostFocus()
    If initChart = True Then DrawChart
End Sub

Private Sub ChartTitle_KeyPress(KeyAscii As Integer)
    If initChart = True And KeyAscii = 13 Then DrawChart
End Sub

Private Sub ChartTitle_LostFocus()
    If initChart = True Then DrawChart
End Sub
'2012-04-19 于登淼 ↑

'2012-04-19 于登淼
'设置画图表的各项参数，然后画图
Sub DrawChart()
    ccmdDrawCrt.Caption = "绘制中..."
    'picChart.Picture = LoadPicture()
    Clipboard.Clear
    '2012-05-29 于登淼 ↓
    '自动填充图表标题、X轴、Y轴内容
'    If ChartTitle.Text = "图表标题" Then ChartTitle.Text = combXAxis.Text & "对" & combYAxis.Text & "统计结果图"
'    If XAxisTitle.Text = "X轴标题" Then XAxisTitle.Text = combXAxis.Text & "分类"
'    If YAxisTitle.Text = "Y轴标题" Then YAxisTitle.Text = combYAxis.Text
    '2012-05-29 于登淼 ↑
    ChartTitle.Text = combXAxis.Text & "对" & combYAxis.Text & "统计结果图"
    XAxisTitle.Text = combXAxis.Text & "分类"
    YAxisTitle.Text = combYAxis.Text
    
    '2012-05-30 于登淼 ↓
    '如果按体检情况统计，则大标题和y轴标签固定
    If combXAxis.ListIndex = 3 Then
        ChartTitle.Text = combXAxis.Text & "统计结果图"
        YAxisTitle.Text = "人数"
    End If
    '2012-05-30 于登淼 ↑
    
'    xlChart.ClearToMatchStyle

'    xlChart.ActiveChart.ClearToMatchStyle       '2015-12-8 by 牟俊
    '设置图表样式
'    xlChart.ActiveChart.ChartType = comb图表样式.ItemData(comb图表样式.ListIndex)   '2015-12-8 by 牟俊
     xlChart.ChartType = comb图表样式.ItemData(comb图表样式.ListIndex)

    '设置色彩样式
'    xlChart.ActiveChart.ChartStyle = comb色彩样式.ItemData(comb色彩样式.ListIndex)   '2015-12-8 by 牟俊
'    xlChart.chartstyle = comb色彩样式.ItemData(comb色彩样式.ListIndex)

    'xlChart.ClearToMatchStyle
    
    '2012-05-29 于登淼 ↓
    '修改了当前所有的图表样式与布局的对应关系，细化到每个图表中
    '设置图表、X轴、Y轴标题
    
    Select Case comb图表样式.ListIndex
    Case 0  '柱状图
        '设置图表布局
'        xlChart.applylayout (comb图表布局.ItemData(comb图表布局.ListIndex))
'
'        Select Case comb图表布局.ListIndex + 1 '图表标题
'            Case 1, 2, 3, 5, 6, 8, 9, 10
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
'        Select Case comb图表布局.ListIndex + 1 'Y轴标题
'            Case 5, 6, 7, 8, 9
'                xlChart.Axes(xlValue).AxisTitle.Select
'                xlChart.Axes(xlValue, xlPrimary).AxisTitle.Text = YAxisTitle.Text
'        End Select
'        Select Case comb图表布局.ListIndex + 1 'X轴标题
'            Case 7, 8, 9
'                xlChart.Axes(xlCategory).AxisTitle.Select
'                xlChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = XAxisTitle.Text
'        End Select
    Case 1  '折线图
        '设置图表布局
'        xlChart.applylayout (comb图表布局.ItemData(comb图表布局.ListIndex))
'
'        Select Case comb图表布局.ListIndex + 1 '图表标题
'            Case 1, 2, 3, 5, 6, 8, 9, 10
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
'        Select Case comb图表布局.ListIndex + 1 'Y轴标题
'            Case 1, 5, 6, 7, 10
'                xlChart.Axes(xlValue).AxisTitle.Select
'                xlChart.Axes(xlValue, xlPrimary).AxisTitle.Text = YAxisTitle.Text
'        End Select
'        Select Case comb图表布局.ListIndex + 1 'X轴标题
'            Case 7, 10
'                xlChart.Axes(xlCategory).AxisTitle.Select
'                xlChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = XAxisTitle.Text
'        End Select
    Case 2 '饼图
'        Select Case comb图表布局.ListIndex + 1 '设置图表布局
'            Case 1, 2, 3, 4, 5, 6, 7
'                xlChart.applylayout (comb图表布局.ItemData(comb图表布局.ListIndex))
'        End Select
'        Select Case comb图表布局.ListIndex + 1 '图表标题
'            Case 1, 2, 5, 6, 8, 9, 10, 11
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
    Case 3  '条形图
        '设置图表布局
'        If Not (comb图表布局.ListIndex + 1 = 11) Then xlChart.applylayout (comb图表布局.ItemData(comb图表布局.ListIndex))
'
'        Select Case comb图表布局.ListIndex + 1 '图表标题
'            Case 1, 2, 3, 5, 6, 8, 9, 11
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
'        Select Case comb图表布局.ListIndex + 1 'Y轴标题
'            Case 6, 7, 8
'                xlChart.Axes(xlValue).AxisTitle.Select
'                xlChart.Axes(xlValue, xlPrimary).AxisTitle.Text = YAxisTitle.Text
'        End Select
'        Select Case comb图表布局.ListIndex + 1 'X轴标题
'            Case 7, 8
'                xlChart.Axes(xlCategory).AxisTitle.Select
'                xlChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = XAxisTitle.Text
'        End Select
    End Select
    '2012-05-29 于登淼 ↑
    
    '设置显示相关
    xlApp.ActiveWindow.Visible = True
    xlChart.ChartArea.Select
    xlChart.ChartArea.Copy
    'Set xlChart = ActiveChart
    picChart.AutoSize = True
'    picChart.Picture = Clipboard.GetData

'将图片放到image中（图片较小看不清） 2016-1-27 by 牟俊
    Image1.Picture = Clipboard.GetData
    Image1.Stretch = True

'将图片放到frmhuatu窗体的image中(图片大些能看清) 2016-1-27 by 牟俊
    frmhuatu.Image1.Picture = Clipboard.GetData
    frmhuatu.Image1.Stretch = True
    frmhuatu.Show
    'picChart.AutoSize = False
    'picChart.PaintPicture picChart.Picture, 0, 0, picChart.ScaleWidth, picChart.ScaleHeight
    ccmdDrawCrt.Caption = "绘制图表"
    
    If picChart.Picture <> 0 Or cgrdStatic.rows >= 1 Then
        ctlb工具栏.Buttons(1).Enabled = True
        ctlb工具栏.Buttons(3).Enabled = True
        ctlb工具栏.Buttons(4).Enabled = True

        ccmdExportCrt.Enabled = True
    End If
    
End Sub

'2012-04-19 于登淼
'初始化图表处理相关list
Sub subInitChartList()
    initChart = False
        
    '图表样式list
    comb图表样式.Clear
    comb图表样式.AddItem "柱状图": comb图表样式.ItemData(comb图表样式.NewIndex) = 51    'xlColumnClustered
    comb图表样式.AddItem "折线图": comb图表样式.ItemData(comb图表样式.NewIndex) = 4     'xlLine
    comb图表样式.AddItem "饼状图": comb图表样式.ItemData(comb图表样式.NewIndex) = 69    'xlPieExploded
    comb图表样式.AddItem "条形图": comb图表样式.ItemData(comb图表样式.NewIndex) = 57    'xlBarClustered
    comb图表样式.ListIndex = 0
    
    '色彩样式list
    comb色彩样式.Clear
    comb色彩样式.AddItem "彩色1": comb色彩样式.ItemData(comb色彩样式.NewIndex) = 2
    comb色彩样式.AddItem "彩色2": comb色彩样式.ItemData(comb色彩样式.NewIndex) = 10
    comb色彩样式.AddItem "彩色3": comb色彩样式.ItemData(comb色彩样式.NewIndex) = 18
    comb色彩样式.AddItem "彩色4": comb色彩样式.ItemData(comb色彩样式.NewIndex) = 26
    comb色彩样式.AddItem "彩色5": comb色彩样式.ItemData(comb色彩样式.NewIndex) = 34
    comb色彩样式.AddItem "彩色6": comb色彩样式.ItemData(comb色彩样式.NewIndex) = 42
    comb色彩样式.ListIndex = 0
     
    '图表布局list
    comb图表布局.Clear
    comb图表布局.AddItem "布局1": comb图表布局.ItemData(comb图表布局.NewIndex) = 1
    comb图表布局.AddItem "布局2": comb图表布局.ItemData(comb图表布局.NewIndex) = 2
    comb图表布局.AddItem "布局3": comb图表布局.ItemData(comb图表布局.NewIndex) = 3
    comb图表布局.AddItem "布局4": comb图表布局.ItemData(comb图表布局.NewIndex) = 4
    comb图表布局.AddItem "布局5": comb图表布局.ItemData(comb图表布局.NewIndex) = 5
    comb图表布局.AddItem "布局6": comb图表布局.ItemData(comb图表布局.NewIndex) = 6
    comb图表布局.AddItem "布局7": comb图表布局.ItemData(comb图表布局.NewIndex) = 7
    comb图表布局.AddItem "布局8": comb图表布局.ItemData(comb图表布局.NewIndex) = 8
    comb图表布局.AddItem "布局9": comb图表布局.ItemData(comb图表布局.NewIndex) = 9
    comb图表布局.AddItem "布局10": comb图表布局.ItemData(comb图表布局.NewIndex) = 10
    comb图表布局.AddItem "布局11": comb图表布局.ItemData(comb图表布局.NewIndex) = 11
    comb图表布局.ListIndex = 0
    
End Sub

'2012-05-29 于登淼
'初始化统计参照（横轴）和统计（纵轴）内容下拉列表
Sub subInitStaticList()
    'X轴选项
    combXAxis.Clear
    combXAxis.AddItem "按工种": combXAxis.ItemData(combXAxis.NewIndex) = 0
    combXAxis.AddItem "按行业": combXAxis.ItemData(combXAxis.NewIndex) = 1
    combXAxis.AddItem "按单位": combXAxis.ItemData(combXAxis.NewIndex) = 2
    combXAxis.AddItem "按体检情况": combXAxis.ItemData(combXAxis.NewIndex) = 3
    '2012-08-08 于登淼 ↓
    combXAxis.AddItem "按危害因素": combXAxis.ItemData(combXAxis.NewIndex) = 4
    '2012-08-08 于登淼 ↑
    '2013-03-31 刘云乐 ↓
    combXAxis.AddItem "按初检/复检": combXAxis.ItemData(combXAxis.NewIndex) = 5
    '2013-03-31 刘云乐 ↑
    combXAxis.ListIndex = 5
    
    'Y轴选项
    combYAxis.Clear
    combYAxis.AddItem "合格人数": combYAxis.ItemData(combYAxis.NewIndex) = 0
    combYAxis.AddItem "不合格人数": combYAxis.ItemData(combYAxis.NewIndex) = 1
    combYAxis.AddItem "合格率": combYAxis.ItemData(combYAxis.NewIndex) = 2
    combYAxis.AddItem "无结果人数": combYAxis.ItemData(combYAxis.NewIndex) = 3
    combYAxis.AddItem "金额": combYAxis.ItemData(combYAxis.NewIndex) = 4
    '2013-03-31 刘云乐 ↓
    combYAxis.AddItem "人数": combYAxis.ItemData(combYAxis.NewIndex) = 5
    '2013-03-31 刘云乐 ↑
    combYAxis.ListIndex = 5
    
End Sub

'2012-05-29 于登淼
'按工种统计函数，统计内容包括：合格人数、不合格人数、合格率、无结果人数、金额
Sub sub按工种统计(ByVal XSelected As Integer, ByVal YSelected As Integer)

    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '根据时间划分了多少行，当前所在时间行
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no工种, no结论 As Integer
    Dim cur工种 As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    '去除重复的系统编号
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP截止.Year - DTP开始.Year) * 12 + (DTP截止.Month - DTP开始.Month) + 1
    
    '初始化统计信息
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP截止.Value - DTP开始.Value < 0 Then MsgBox ("结束日期必须大于初始日期"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "人数：" & cgrdInfo.rows - 1
'    Label9.Caption = "系统编号不重复：" & TotalLines
    
    no结论 = 0
    Set cur工种 = New Collection
    cur工种.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        strSQL = "select 现工种 from 职业病体检_体检人员基本信息表 where 系统编号='" & strSysNo & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("现工种"))
        
        ifAdd = True
        For j = 1 To cur工种.Count
            If cur工种.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then cur工种.Add queryinfo(i, 0)
        
        strSQL = "select * from 职业病体检_科室结论表 where 科室='16' and 系统编号='" & strSysNo & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP截止.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub最终合格(IIf(IsNull(lobjRec("文字结论")), "", lobjRec("文字结论")))
            queryinfo(i, 2) = lobjRec("结论日期")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '有结论，覆盖结果
            
            strSQL = "select * from 职业病体检_体检基本信息表 where 系统编号='" & strSysNo & "'"
            dasubSetQueryTimeout 6000
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("收费金额")), "0", lobjRec("收费金额"))
        Else
            no结论 = no结论 + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '无结论
            queryinfo(i, 4) = "0"
        End If
        
    Next
    cur工种.Remove (1)
    cur工种.Add ""  '无工种名称为未分类，放在集合最后面
'    Label10.Caption = "有结果人数：" & (TotalLines - no结论)
    
    '计算合格人数和总人数
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur工种.Count + 1) * 4) As Double
    no工种 = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP开始.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP开始.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        For j = 0 To cur工种.Count - 1
            If queryinfo(i, 0) = cur工种.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then no工种 = no工种 + 1
        Next j
    Next i
        
    '初始化cgrdStatic表格内容和表头格式
    With cgrdStatic
        .Clear
        .cols = cur工种.Count + IIf(no工种 > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To cur工种.Count - 1
            .TextMatrix(0, i) = cur工种.Item(i)
        Next
        If no工种 > 0 Then .TextMatrix(0, i) = "未分类"
        tmpDate = DTP开始.Value
        i = 1
        While tmpDate < DTP截止.Value
            '修改人：张令 2012.12.10
            '说明：将区间改为CombInterval的值，如：日，月，季，年等。↓↓
'            .TextMatrix(i, 0) = "第" & i & "区间"
            .TextMatrix(i, 0) = "第" & i & CombInterval.Text
            '2012.12.11   ↑↑
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '将计算结果填入cgrdStatic中
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur工种.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur工种.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur工种.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur工种.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur工种.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    
    dafuncGetData ("update 职业病体检_业务设置信息表 set 设置项目='统计内容_按工种',设置值='统计类别_" & combYAxis.Text & "',枚举来源='" & Xnum & "',说明='" & Ynum & "' where left(设置项目,5)='统计内容_'")
End Sub

Sub sub按行业统计(ByVal XSelected As Integer, ByVal YSelected As Integer)

    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '根据时间划分了多少行，当前所在时间行
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no行业, no结论 As Integer
    Dim cur行业 As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    '去除重复的系统编号
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP截止.Year - DTP开始.Year) * 12 + (DTP截止.Month - DTP开始.Month) + 1
    
    '初始化统计信息
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP截止.Value - DTP开始.Value < 0 Then MsgBox ("结束日期必须大于初始日期"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "人数：" & cgrdInfo.rows - 1
'    Label9.Caption = "系统编号不重复：" & TotalLines
    
    
    no结论 = 0
    Set cur行业 = New Collection
    cur行业.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        strSQL = "select 单位名称 from 职业病体检_体检人员基本信息表 where 系统编号='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("单位名称"))
        
        If Not queryinfo(i, 0) = "" Then
            strSQL = "select * from 单位档案_单位基本信息表 where 单位名称='" & queryinfo(i, 0) & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("行业类别"))
        End If
        
        ifAdd = True
        For j = 1 To cur行业.Count
            If cur行业.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then cur行业.Add queryinfo(i, 0)
        
        strSQL = "select * from 职业病体检_科室结论表 where 科室='16' and 系统编号='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP截止.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub最终合格(IIf(IsNull(lobjRec("文字结论")), "", lobjRec("文字结论")))
            queryinfo(i, 2) = lobjRec("结论日期")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '有结论，覆盖结果
            
            strSQL = "select * from 职业病体检_体检基本信息表 where 系统编号='" & strSysNo & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("收费金额")), "0", lobjRec("收费金额"))
        Else
            no结论 = no结论 + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '无结论
            queryinfo(i, 4) = "0"
        End If
    Next
    cur行业.Remove (1)
    cur行业.Add ""  '无行业名称为未分类，放在集合最后面
'    Label10.Caption = "有结果人数：" & (TotalLines - no结论)
    
    '计算合格人数和总人数
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur行业.Count + 1) * 4) As Double
    no行业 = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP开始.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP开始.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        
        For j = 0 To cur行业.Count - 1
            If queryinfo(i, 0) = cur行业.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then no行业 = no行业 + 1
        Next j
    Next i
        
        
    '初始化cgrdStatic表格内容和表头格式
    With cgrdStatic
        .Clear
        .cols = cur行业.Count + IIf(no行业 > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To cur行业.Count - 1
            .TextMatrix(0, i) = cur行业.Item(i)
        Next
        If no行业 > 0 Then .TextMatrix(0, i) = "未分类"
        tmpDate = DTP开始.Value
        i = 1
        While tmpDate < DTP截止.Value
            '修改人：张令 2012.12.10
            '说明：将区间改为CombInterval的值，如：日，月，季，年等。↓↓
'            .TextMatrix(i, 0) = "第" & i & "区间"
            .TextMatrix(i, 0) = "第" & i & CombInterval.Text
            '2012.12.11   ↑↑
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '将计算结果填入cgrdStatic中
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur行业.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur行业.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur行业.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur行业.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur行业.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    dafuncGetData ("update 职业病体检_业务设置信息表 set 设置项目='统计内容_按行业',设置值='统计类别_" & combYAxis.Text & "',枚举来源='" & Xnum & "',说明='" & Ynum & "' where left(设置项目,5)='统计内容_'")
End Sub

'2012-05-30 于登淼
'按体检单位统计
Sub sub按单位统计(ByVal XSelected As Integer, ByVal YSelected As Integer)
    
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '根据时间划分了多少行，当前所在时间行
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no单位, no结论 As Integer
    Dim cur单位 As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    '去除重复的系统编号
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    Month = (DTP截止.Year - DTP开始.Year) * 12 + (DTP截止.Month - DTP开始.Month) + 1
    
    '初始化统计信息
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP截止.Value - DTP开始.Value < 0 Then MsgBox ("结束日期必须大于初始日期"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "人数：" & cgrdInfo.rows - 1
'    Label9.Caption = "系统编号不重复：" & TotalLines
    
    
    no结论 = 0
    Set cur单位 = New Collection
    cur单位.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows -1
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 0)
            
        strSQL = "select 单位名称 from 职业病体检_体检人员基本信息表 where 系统编号='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("单位名称"))
        
        ifAdd = True
        For j = 1 To cur单位.Count
            If cur单位.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then cur单位.Add queryinfo(i, 0)
        
        strSQL = "select * from 职业病体检_科室结论表 where 科室='16' and 系统编号='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP截止.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub最终合格(IIf(IsNull(lobjRec("文字结论")), "", lobjRec("文字结论")))
            queryinfo(i, 2) = lobjRec("结论日期")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '有结论，覆盖结果
            
            strSQL = "select * from 职业病体检_体检基本信息表 where 系统编号='" & strSysNo & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("收费金额")), "0", lobjRec("收费金额"))
        Else
            no结论 = no结论 + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '无结论
            queryinfo(i, 4) = "0"
        End If
    Next
    cur单位.Remove (1)
    cur单位.Add ""  '无单位名称为未分类，放在集合最后面
'    Label10.Caption = "有结果人数：" & (TotalLines - no结论)
    
    '计算合格人数和总人数
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur单位.Count + 1) * 4) As Double
    no单位 = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP开始.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP开始.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        
        For j = 0 To cur单位.Count - 1
            If queryinfo(i, 0) = cur单位.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then no单位 = no单位 + 1
        Next j
    Next i
        
        
    '初始化cgrdStatic表格内容和表头格式
    With cgrdStatic
        .Clear
        .cols = cur单位.Count + IIf(no单位 > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To cur单位.Count - 1
            .TextMatrix(0, i) = cur单位.Item(i)
        Next
        If no单位 > 0 Then .TextMatrix(0, i) = "未分类"
        tmpDate = DTP开始.Value
        i = 1
        While tmpDate < DTP截止.Value
            '修改人：张令 2012.12.10
            '说明：将区间改为CombInterval的值，如：日，月，季，年等。↓↓
'            .TextMatrix(i, 0) = "第" & i & "区间"
            .TextMatrix(i, 0) = "第" & i & CombInterval.Text
            '2012.12.11   ↑↑
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '将计算结果填入cgrdStatic中
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur单位.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur单位.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur单位.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur单位.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur单位.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    dafuncGetData ("update 职业病体检_业务设置信息表 set 设置项目='统计内容_按单位',设置值='统计类别_" & combYAxis.Text & "',枚举来源='" & Xnum & "',说明='" & Ynum & "' where left(设置项目,5)='统计内容_'")
End Sub

'2012-05-30 于登淼
'按体检情况统计
Sub sub按体检情况统计(ByVal XSelected As Integer, ByVal YSelected As Integer)
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '根据时间划分了多少行，当前所在时间行
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no结论, numClassify As Integer
    Dim Xnum, Ynum, Znum As Integer
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    '去除重复的系统编号
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP截止.Year - DTP开始.Year) * 12 + (DTP截止.Month - DTP开始.Month) + 1
    
    '初始化统计信息
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP截止.Value - DTP开始.Value < 0 Then MsgBox ("结束日期必须大于初始日期"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "人数：" & cgrdInfo.rows - 1
'    Label9.Caption = "系统编号不重复：" & TotalLines
    
    ReDim queryinfo(1 To TotalLines, 1 To 3) As String
    
    no结论 = TotalLines
    For i = 1 To TotalLines
        strSysNo = SysNo(i)
        strSQL = "select * from 职业病体检_科室结论表 where 科室='16' and 系统编号='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        On Error Resume Next
        queryinfo(i, 1) = "0"
        queryinfo(i, 2) = DTP截止.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub最终合格(IIf(IsNull(lobjRec("文字结论")), "", lobjRec("文字结论")))
            queryinfo(i, 2) = lobjRec("结论日期")
            If lobjRec("结论日期").RecordCount = 1 Then
                no结论 = no结论 - 1
                queryinfo(i, 3) = "1"   '有结论，覆盖结果
            End If
        Else
            queryinfo(i, 1) = ""
            queryinfo(i, 2) = ""
            queryinfo(i, 3) = "0"   '无结论
        End If
    Next
'    Label10.Caption = "有结果人数：" & (TotalLines - no结论)
    
    '计算合格人数和总人数
    numClassify = 2 + IIf(no结论 > 0, 1, 0) '包括“合格人数”、“不合格人数”，可能包括“无结果人数”
    ReDim StaticGrid(1 To Timex + 1, 1 To numClassify) As Double
    For i = 1 To TotalLines
        Month = (Mid(queryinfo(i, 2), 1, 4) - DTP开始.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP开始.Month) + 1
        curTimex = Round(Month / TimeInterval)
        If Round(Month / TimeInterval) < Month / TimeInterval Then
            curTimex = Round(Month / TimeInterval) + 1
        End If
        
        If queryinfo(i, 1) = "1" Then
            StaticGrid(curTimex, 1) = StaticGrid(curTimex, 1) + 1
        ElseIf queryinfo(i, 1) = "0" Then
            StaticGrid(curTimex, 2) = StaticGrid(curTimex, 2) + 1
        Else
            StaticGrid(curTimex, 3) = StaticGrid(curTimex, 3) + 1
        End If
'        If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, 3) = StaticGrid(curTimex, 3) + 1
    Next i
        
    '初始化cgrdStatic表格内容和表头格式
    With cgrdStatic
        .Clear
        .cols = numClassify + 1
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        .TextMatrix(0, 1) = "合格人数"
        .TextMatrix(0, 2) = "不合格人数"
        If no结论 > 0 Then .TextMatrix(0, 3) = "无结果人数"
        tmpDate = DTP开始.Value
        i = 1
        While tmpDate < DTP截止.Value
            '修改人：张令 2012.12.10
            '说明：将区间改为CombInterval的值，如：日，月，季，年等。↓↓
'            .TextMatrix(i, 0) = "第" & i & "区间"
            .TextMatrix(i, 0) = "第" & i & CombInterval.Text
            '2012.12.11   ↑↑
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '将计算结果填入cgrdStatic中
    Xnum = 0
    Ynum = 0
    Znum = 0
    With cgrdStatic
        For i = 1 To Timex
            For j = 1 To numClassify
                .TextMatrix(i, j) = StaticGrid(i, j)
            Next
            Xnum = Xnum + StaticGrid(i, 1)
            Ynum = Ynum + StaticGrid(i, 2)
            If numClassify = 3 Then Znum = Znum + StaticGrid(i, 3)
        Next
        .AutoSize 0, .cols - 1, 0, 0
    End With
    dafuncGetData ("update 职业病体检_业务设置信息表 set 设置值='" & Xnum & "',枚举来源='" & Ynum & "',说明='" & Znum & "' where 设置项目='统计内容-按体检情况'")
End Sub

'2012-08-08 于登淼
'按危害因素统计
Sub sub按危害因素统计(ByVal XSelected As Integer, ByVal YSelected As Integer)
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '根据时间划分了多少行，当前所在时间行
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no危害, no结论 As Integer
    Dim cur危害 As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    '去除重复的系统编号
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP截止.Year - DTP开始.Year) * 12 + (DTP截止.Month - DTP开始.Month) + 1
    
    '初始化统计信息
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP截止.Value - DTP开始.Value < 0 Then MsgBox ("结束日期必须大于初始日期"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "人数：" & cgrdInfo.rows - 1
'    Label9.Caption = "系统编号不重复：" & TotalLines
    
    
    no结论 = 0
    Set cur危害 = New Collection
    cur危害.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        strSQL = "select 危害因素 from 职业病体检_体检人员基本信息表 where 系统编号='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("危害因素"))
        
        ifAdd = True
        For j = 1 To cur危害.Count
            If cur危害.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then cur危害.Add queryinfo(i, 0)
        
        strSQL = "select * from 职业病体检_科室结论表 where 科室='16' and 系统编号='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP截止.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub最终合格(IIf(IsNull(lobjRec("文字结论")), "", lobjRec("文字结论")))
            queryinfo(i, 2) = lobjRec("结论日期")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '有结论，覆盖结果
            
            strSQL = "select * from 职业病体检_体检基本信息表 where 系统编号='" & strSysNo & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("收费金额")), "0", lobjRec("收费金额"))
        Else
            no结论 = no结论 + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '无结论
            queryinfo(i, 4) = "0"
        End If
    Next
    cur危害.Remove (1)
    cur危害.Add ""  '无危害名称为未分类，放在集合最后面
'    Label10.Caption = "有结果人数：" & (TotalLines - no结论)
    
    '计算合格人数和总人数
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur危害.Count + 1) * 4) As Double
    no危害 = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP开始.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP开始.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        
        For j = 0 To cur危害.Count - 1
            If queryinfo(i, 0) = cur危害.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then no危害 = no危害 + 1
        Next j
                
    Next i
        
        
    '初始化cgrdStatic表格内容和表头格式
    With cgrdStatic
        .Clear
        .cols = cur危害.Count + IIf(no危害 > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To cur危害.Count - 1
            .TextMatrix(0, i) = cur危害.Item(i)
        Next
        If no危害 > 0 Then .TextMatrix(0, i) = "未分类"
        tmpDate = DTP开始.Value
        i = 1
        While tmpDate < DTP截止.Value
            '修改人：张令 2012.12.10
            '说明：将区间改为CombInterval的值，如：日，月，季，年等。↓↓
'            .TextMatrix(i, 0) = "第" & i & "区间"
            .TextMatrix(i, 0) = "第" & i & CombInterval.Text
            '2012.12.11   ↑↑
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '将计算结果填入cgrdStatic中
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur危害.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur危害.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur危害.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur危害.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '修改人：张令 2012.12.10
                '说明：此处cur工种的条数用表格中获取的条数代替。↓↓
'                For j = 0 To cur危害.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ↑↑
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    dafuncGetData ("update 职业病体检_业务设置信息表 set 设置项目='统计内容_按危害',设置值='统计类别_" & combYAxis.Text & "',枚举来源='" & Xnum & "',说明='" & Ynum & "' where left(设置项目,5)='统计内容_'")

End Sub

'2012-06-06 于登淼
'添加函数，打印统计报告
Sub sub打印统计报告(ByVal para预览报告 As Boolean, ByVal para打印报告 As Boolean)
    Dim lcolID As Collection
    Dim para报告名称 As String
    
    Set lcolID = New Collection
    lcolID.Add "27"
    
    Dim tmplobj As Object
    Set tmplobj = CreateObject("职业病文书.cls文书")
    
    VB.SavePicture picChart.Picture, "C:\统计图表.bmp"
    
    cgrdStatic.PictureType = flexPictureColor
    tmpPicture.AutoSize = True
    Clipboard.Clear
    Clipboard.SetData cgrdStatic.Picture
    tmpPicture.Picture = Clipboard.GetData
    VB.SavePicture tmpPicture.Picture, "C:\统计数据.bmp"
    
    If combXAxis.ListIndex <> 3 And combYAxis.ListIndex <> 2 And combYAxis.ListIndex <> 4 Then
        para报告名称 = "职业病体检_统计报告(人数)"
    ElseIf combXAxis.ListIndex <> 3 And combYAxis.ListIndex = 2 Then
        para报告名称 = "职业病体检_统计报告(合格率)"
    ElseIf combXAxis.ListIndex <> 3 And combYAxis.ListIndex = 4 Then
        para报告名称 = "职业病体检_统计报告(金额)"
    ElseIf combXAxis.ListIndex = 3 Then
        para报告名称 = "职业病体检_统计报告(按体检情况)"
    End If
    
    If para预览报告 = True Then
        tmplobj.Sub打印文书 para报告名称, lcolID, True, True, "tmp体检条码号", False
    Else
        tmplobj.Sub打印文书 para报告名称, lcolID, False, False, "tmp体检条码号", False
    End If
    
    Set lcolID = Nothing

End Sub

'2013-03-31 刘云乐
'按时间人数统计
Sub sub按时间人数统计(ByVal XSelected As Integer, ByVal YSelected As Integer)
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '根据时间划分了多少行，当前所在时间行
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim TotalLinesF As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no危害, no结论 As Integer
    Dim cur危害 As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 5) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim SysNoF(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j, k As Integer
    
    Dim Month As String
    Dim peopleCount As Double
    '去除重复的系统编号
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    TotalLinesF = 1
    'For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
'        If flag(i) = True Then
'            For j = i + 1 To cgrdInfo.rows - 1
'                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
'            Next j
'        End If
        'If flag(i) = True Then
        'If Right(cgrdInfo.TextMatrix(i, 0), 1) <> "F" Then
            SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
        'Else
        '    SysNoF(TotalLinesF) = cgrdInfo.TextMatrix(i, 0): TotalLinesF = TotalLinesF + 1
        'End If
    Next i
    
    TotalLines = TotalLines - 1
    TotalLinesF = TotalLinesF - 1
    
    Month = (DTP截止.Year - DTP开始.Year) * 12 + (DTP截止.Month - DTP开始.Month) + 1
    
    '初始化统计信息
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP截止.Value - DTP开始.Value < 0 Then MsgBox ("结束日期必须大于初始日期"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "人数：" & cgrdInfo.rows - 1
'    Label9.Caption = "系统编号不重复：" & TotalLines
    
    
    'no结论 = 0
    Set cur危害 = New Collection
    cur危害.Add "初检"
    cur危害.Add "复检"
    For k = 1 To cgrdInfo.cols
        If cgrdInfo.TextMatrix(0, k) = "体检日期" Then
            Exit For
        End If
    Next k
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        If Right(strSysNo, 1) = "F" Then
            queryinfo(i, 0) = "复检"
        Else
            queryinfo(i, 0) = "初检"
        End If
        
        If Len(Trim(cgrdInfo.TextMatrix(i, k))) = 10 Then
            queryinfo(i, 2) = cgrdInfo.TextMatrix(i, k)
        Else
            queryinfo(i, 2) = "2013-01-01"
        End If
        
        
    Next
   
    '计算合格人数和总人数
    'ReDim StaticGrid(1 To Timex + 1, 1 To (cur危害.Count + 1) * 4) As Double
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur危害.Count + 1) * 5) As Double
    'no危害 = 0
    'peopleCount = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
'        If queryinfo(i, 2) = "0" Then
'            curTimex = 1
'        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP开始.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP开始.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
'        End If
        
        For j = 0 To cur危害.Count - 1
        'j = 0
            If queryinfo(i, 0) = cur危害.Item(j + 1) Then
                'If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then
'                StaticGrid(curTimex, j * 5 + 2) = StaticGrid(curTimex, j * 5 + 2) + 1
'                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 5 + 1) = StaticGrid(curTimex, j * 5 + 1) + 1
'                StaticGrid(curTimex, j * 5 + 3) = StaticGrid(curTimex, j * 5 + 3) + 1
'                StaticGrid(curTimex, j * 5 + 4) = StaticGrid(curTimex, j * 5 + 4) + CDbl(queryinfo(i, 4))
                'If Right(SysNo(i), 1) = "F" Then
                '    StaticGrid(curTimex, j * 5 + 0) = StaticGrid(curTimex, j * 5 + 0) + 1
                'Else
                    StaticGrid(curTimex, j * 5 + 5) = StaticGrid(curTimex, j * 5 + 5) + 1
                'End If
                
                
            End If
            'If queryinfo(i, 0) = "" Then no危害 = no危害 + 1
            'If queryinfo(i, 5) = "1" Then peopleCount = peopleCount + 1
        Next j
                
    Next i
        
        
    '初始化cgrdStatic表格内容和表头格式
    With cgrdStatic
        .Clear
'        .cols = cur危害.Count + IIf(no危害 > 0, 1, 0)
        .cols = 3
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
'        For i = 1 To cur危害.Count - 1
'            .TextMatrix(0, i) = cur危害.Item(i)
'        Next
        .TextMatrix(0, 1) = "初检"
        .TextMatrix(0, 2) = "复检"
        'If no危害 > 0 Then .TextMatrix(0, i) = "未分类"
        tmpDate = DTP开始.Value
        i = 1
        While tmpDate < DTP截止.Value
            '修改人：张令 2012.12.10
            '说明：将区间改为CombInterval的值，如：日，月，季，年等。↓↓
'            .TextMatrix(i, 0) = "第" & i & "区间"
            .TextMatrix(i, 0) = "第" & i & CombInterval.Text
            '2012.12.11   ↑↑
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '将计算结果填入cgrdStatic中
    Xnum = 0
    Ynum = 0
    Select Case YSelected

    '2013-03-31 刘云乐
    Case 5
        With cgrdStatic
            For i = 1 To Timex
                For j = 0 To cgrdStatic.cols - 2
                    .TextMatrix(i, j + 1) = StaticGrid(i, 5 * j + 5)
                    
                    'Xnum = Xnum + StaticGrid(i, 5 * j + 4)
                    'Ynum = Ynum + StaticGrid(i, 5 * j + 3)
                Next
            Next
        End With
    '2013-03-31
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    
    dafuncGetData ("update 职业病体检_业务设置信息表 set 设置项目='统计内容_按危害',设置值='统计类别_" & combYAxis.Text & "',枚举来源='" & Xnum & "',说明='" & Ynum & "' where left(设置项目,5)='统计内容_'")

End Sub
