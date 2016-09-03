VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAssessDeformity 
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   14910
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "录入内容"
      Height          =   8295
      Left            =   10080
      TabIndex        =   15
      Top             =   840
      Width           =   4455
      Begin VB.Frame Frame历史评残 
         Caption         =   "上次评残"
         Height          =   3135
         Left            =   120
         TabIndex        =   31
         Top             =   5040
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox Text历史结论 
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   2520
            Width           =   3975
         End
         Begin VB.TextBox Text历史结果 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox Text历史时间 
            Height          =   375
            Left            =   720
            TabIndex        =   35
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox Text历史等级 
            Height          =   375
            Left            =   720
            TabIndex        =   33
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label14 
            Caption         =   "体检结论："
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "体检结果："
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "时间："
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "等级："
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
      End
      Begin MSComCtl2.DTPicker DTPicker时间 
         Height          =   375
         Left            =   840
         TabIndex        =   26
         Top             =   4560
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59637761
         CurrentDate     =   42520
      End
      Begin VB.TextBox Text等级 
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox Text资料 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox Text结论 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox Text结果 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label8 
         Caption         =   "时间："
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "等级："
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "评残："
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "提供资料："
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "体检结论："
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "体检结果："
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "查询条件"
      Height          =   6495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      Begin VB.CheckBox Check调残 
         Caption         =   "调残"
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check评残 
         Caption         =   "评残"
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   360
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.OptionButton coptType 
         Caption         =   "提供证明"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   4440
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "已复核"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   38
         Top             =   4920
         Width           =   975
      End
      Begin VB.ComboBox Combo类别 
         Height          =   300
         Left            =   1440
         TabIndex        =   30
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox Check类别 
         Caption         =   "人员类别："
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Com查询 
         Caption         =   "查询"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   5640
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "评残未复核"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   4920
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "待评残"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text单位 
         Height          =   270
         Left            =   1440
         TabIndex        =   11
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox Check单位 
         Caption         =   "单位名称："
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text编号 
         Height          =   270
         Left            =   1440
         TabIndex        =   9
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CheckBox Check编号 
         Caption         =   "体检编号："
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker结束 
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Format          =   59637760
         CurrentDate     =   42520
      End
      Begin MSComCtl2.DTPicker DTPicker开始 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Format          =   59637760
         CurrentDate     =   42520
      End
      Begin VB.CheckBox Check日期 
         Caption         =   "体检日期："
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ComboBox Combo类型 
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox Check类型 
         Caption         =   "人员类型："
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "到"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   2400
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   9600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdInfo 
      Height          =   6015
      Left            =   4200
      TabIndex        =   27
      Top             =   1440
      Width           =   5775
      _cx             =   10186
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
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
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1058
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   240
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label12 
      Caption         =   "总人数："
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "这是人数"
      Height          =   255
      Left            =   5040
      TabIndex        =   36
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "这是标题"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   960
      Width           =   5775
   End
End
Attribute VB_Name = "frmAssessDeformity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mobjQueryResult As Object

Private Sub cgrdInfo_Click()
subClear
'Com查询_Click
Dim SyNo As String
Dim obj As Object
SyNo = cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号"))
Set obj = dafuncGetData("select * from 职业病体检_评残信息表 where 系统编号='" & SyNo & "'")
If obj.RecordCount > 0 Then
    If Not IsNull(obj("提供资料信息")) Then
        Text资料.Text = obj("提供资料信息")
    End If
    If Not IsNull(obj("评残等级")) Then
        Text等级.Text = obj("评残等级")
    End If
    If Not IsNull("评残时间") Then
        DTPicker时间.Value = obj("评残时间")
    End If

'    '判断是否已经评残
'    If Not IsNull(obj("是否已评残")) Then
'        If obj("是否已评残") = 1 And coptType(1).Value = True Then '1代表历史评过残，0代表没有评过残
'            Frame历史评残.Visible = True
'            Text历史等级.Text = IIf(Not IsNull(obj("历史评残等级")), obj("历史评残等级"), "")
''            Text历史等级.Text = obj("历史评残等级")
'            Text历史时间.Text = IIf(Not IsNull(obj("历史评残时间")), obj("历史评残时间"), "")
''            Text历史时间.Text = obj("历史评残时间")
'        End If
'    Else
'        Frame历史评残.Visible = False
'    End If
End If
'如果是调残，显示下面内容
'If Check调残.Value = 1 Then
    '身份证号判断以前是否有过评残历史
    Dim obj1 As Object
    Set obj1 = dafuncGetData("select a.评残等级,a.评残时间,a.体检结果,a.体检结论 from 职业病体检_评残详细历史信息表 a left join 职业病体检_体检人员基本信息表 b on a.身份证号=b.公民身份号码 where b.公民身份号码='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("身份证号")) & "' order by a.体检时间 desc")
    If obj1.RecordCount > 0 Then
        Frame历史评残.Visible = True
        Text历史等级.Text = obj1("评残等级")
        Text历史时间.Text = obj1("评残时间")
        Text历史结果.Text = obj1("体检结果")
        Text历史结论.Text = obj1("体检结论")
    Else
        Frame历史评残.Visible = False
    End If
'End If
Text结果.Text = cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("体检结果"))
Text结论.Text = cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("体检结论"))
End Sub

Private Sub Check类别_Click()
    If Check类别.Value = 0 Then
        Combo类别.Text = ""
    End If
End Sub

Private Sub Check类型_Click()
    If Check类型.Value = 0 Then
        Combo类型.Text = ""
    End If
End Sub

Private Sub Com查询_Click()
On Error Resume Next
Dim lsql类型 As String
Dim lsql类别 As String
Dim lsql开始日期 As Date
Dim lsql结束日期 As Date
Dim lsql编号 As String
Dim lsql单位 As String
Dim lsqlwhere As String
Dim lsql查询 As String

lsqlwhere = ""
'根据意见查找要评残人员
If Check评残.Value = 1 Then
    If coptType(1).Value = True Or coptType(2).Value = True Then
        lsqlwhere = lsqlwhere + " b.诊断和处理意见 like '%不排除%' and a.系统编号=b.系统编号 "
    ElseIf coptType(0).Value = True Then
        lsqlwhere = lsqlwhere + "(b.体检状态='11' or b.诊断和处理意见 like '%不排除%') and a.系统编号=b.系统编号 "
    Else
        lsqlwhere = lsqlwhere + "(b.体检状态='6' or b.体检状态='7' or b.体检状态='8') and a.系统编号=b.系统编号 "
    End If
Else   '调残   暂时没法取值，未完成
    If coptType(1).Value = True Or coptType(2).Value = True Then
        lsqlwhere = lsqlwhere + " b.诊断和处理意见 like '%不排除%' and a.系统编号=b.系统编号 "
    ElseIf coptType(0).Value = True Then
        lsqlwhere = lsqlwhere + "(b.体检状态='11' or b.诊断和处理意见 like '%不排除%') and a.系统编号=b.系统编号 "
    Else
        lsqlwhere = lsqlwhere + "(b.体检状态='6' or b.体检状态='7' or b.体检状态='8') and a.系统编号=b.系统编号 "
    End If
End If
'组装查询条件
'1.体检类型
    If Check类型.Value = 1 Then
        lsql类型 = Combo类型.Text
        lsqlwhere = lsqlwhere + "and a.体检表类型='" & lsql类型 & "' "
    Else
        Combo类型.Text = ""
        lsql类型 = ""
    End If
'2.体检类别
    If Check类别.Value = 1 Then
        lsql类别 = Combo类别.Text
        lsqlwhere = lsqlwhere + " and a.体检表类别='" & lsql类别 & "' "
    Else
        Combo类别.Text = ""
        lsql类别 = ""
    End If
'3.编号
    If Check编号.Value = 1 Then
        lsql编号 = Text编号.Text
        lsqlwhere = lsqlwhere + "and b.系统编号='" & lsql编号 & "' "
    Else
        Text编号.Text = ""
        lsql编号 = ""
    End If
'4.单位
    If Check单位.Value = 1 Then
        lsql单位 = Text单位.Text
        lsqlwhere = lsqlwhere + "and a.单位名称='" & lsql单位 & "' "
    Else
        Text单位.Text = ""
        lsql单位 = ""
    End If
'5.体检日期
    If Check日期.Value = 1 Then
        lsql开始日期 = DTPicker开始.Value
        lsql结束日期 = DTPicker结束.Value
        lsqlwhere = lsqlwhere + " and b.体检日期 between '" & DTPicker开始.Value & " 00:00:00' and '" & DTPicker结束.Value & " 23:59:59'"
    Else
        lsql开始日期 = ""
        lsql结束日期 = ""
    End If
If Check评残.Value = 1 Then
    If coptType(0).Value = True Then
        ctlb工具栏.Buttons(3).Visible = False
        ctlb工具栏.Buttons(4).Visible = True
        ctlb工具栏.Buttons(5).Visible = False
        ctlb工具栏.Buttons(6).Visible = False
        ctlb工具栏.Buttons(8).Visible = False
'        lsqlwhere = lsqlwhere + " and c.系统编号 is null"
        lsqlwhere = lsqlwhere + " and (c.系统编号 is null or c.评残状态='1')"
        lsql查询 = "select b.系统编号 as 编号,a.公民身份号码 as 身份证号,convert(varchar(10),b.体检日期,111) 体检时间 ,a.单位名称 as 地区,a.姓名,a.性别,a.年龄,b.体检结论 as 体检结果,b.诊断和处理意见 as 体检结论 from 职业病体检_体检人员基本信息表 a ,职业病体检_体检基本信息表 b left join 职业病体检_评残信息表 c on b.系统编号=c.系统编号 where " & lsqlwhere & ""
    ElseIf coptType(1).Value = True Then
        ctlb工具栏.Buttons(3).Visible = True
        ctlb工具栏.Buttons(5).Visible = True
        ctlb工具栏.Buttons(6).Visible = True
        ctlb工具栏.Buttons(8).Visible = False
        lsqlwhere = lsqlwhere + " and c.系统编号=b.系统编号 and 评残状态='2'"
        lsql查询 = "select b.系统编号 as 编号,a.公民身份号码 as 身份证号,convert(varchar(10),b.体检日期,111) 体检时间,a.单位名称 as 地区,a.姓名,a.性别,a.年龄,b.体检结论 as 体检结果,b.诊断和处理意见 as 体检结论 from 职业病体检_体检人员基本信息表 a ,职业病体检_体检基本信息表 b, 职业病体检_评残信息表 c where " & lsqlwhere & ""
    ElseIf coptType(2).Value = True Then
        ctlb工具栏.Buttons(4).Visible = False
        ctlb工具栏.Buttons(5).Visible = False
        ctlb工具栏.Buttons(6).Visible = False
        ctlb工具栏.Buttons(8).Visible = True
        lsqlwhere = lsqlwhere + " and c.系统编号=b.系统编号 and 评残状态='3'"
        lsql查询 = "select b.系统编号 as 编号,a.公民身份号码 as 身份证号,convert(varchar(10),b.体检日期,111) 体检时间,a.单位名称 as 地区,a.姓名,a.性别,a.年龄,b.体检结论 as 体检结果,b.诊断和处理意见 as 体检结论 from 职业病体检_体检人员基本信息表 a ,职业病体检_体检基本信息表 b, 职业病体检_评残信息表 c where " & lsqlwhere & ""
    End If
Else   '调残  未完成
    '待补内容
End If
    Set mobjQueryResult = dafuncGetData(lsql查询)
    Set cgrdInfo.DataSource = mobjQueryResult
    '隐藏身份证号列
    cgrdInfo.ColHidden(cgrdInfo.ColIndex("身份证号")) = True
    '取标题
    Label1.Caption = Format(Now, "yyyy") + "年" + IIf(Combo类型.Text = "涉核部队YK", "铀矿部队", IIf(Combo类型.Text = "8023部队", "原8023部队", Combo类型.Text)) + "退役人员" + Combo类别.Text + "不排除核辐射影响人员一览表"
    Label11.Caption = cgrdInfo.rows - 1
    'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
    cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    cgrdInfo.Col = 0
    cgrdInfo.Sort = flexSortGenericDescending
    subClear
    ctlb工具栏.Buttons(3).Visible = False
    Frame历史评残.Visible = False
End Sub

Private Sub coptType_Click(Index As Integer)
    Com查询_Click
    If cgrdInfo.rows > 1 Then
        cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
        cgrdInfo.ExplorerBar = flexExSort
        cgrdInfo.DataMode = flexDMFree
        cgrdInfo.Col = 0
        cgrdInfo.Sort = flexSortGenericDescending
    End If
End Sub

Private Sub Form_Load()
 On Error Resume Next
Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
 Set mobjGUI = New cls界面通用对象
     '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
        '设置窗体正在使用的标志。
    mblnInUse = True
    
    '设置工具栏上所需要的各种按钮。
    With lcol工具栏按钮
        .Add "导出Excel(&O)113"     '1
        .Add "|"
        .Add "删除"            '3
        .Add "保存(&S)101"     '4
        .Add "取消保存(&E)111"     '5
        .Add "复核(&F)109"     '6
        .Add "|"
        .Add "取消复核(&E)111" '8
        .Add "|"
        .Add "退出"            '10
    End With
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlb工具栏
    End With
    mobjGUI.subInitialize lcol工具栏按钮, ""
    DoEvents
    
    With cgrdInfo
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "编号"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "身份证号"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检时间"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "地区"       '后面地区的值是暂时取的单位
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "姓名"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "性别"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "年龄"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检结果"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检结论"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    
    '初始化界面查询条件
    listcombox
    Check类型.Value = 1
    Check类别.Value = 1
    Check日期.Value = 1
'    coptType(0).Value = True
    DTPicker结束.Value = Format(Now, "yyyy-MM-dd")
    DTPicker开始.Value = Format(DateAdd("M", -5, Now()), "yyyy/MM/dd")
    DTPicker时间.Value = Now
    Com查询_Click
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
 Dim lobj As Object
 On Error Resume Next
Cancel = True
Select Case Operate
    Case "导出Excel"
        If cgrdInfo.rows <= 1 Then
            MsgBox "没有需要导出的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        Dim lstrFile As String
        ccmdFile.Filter = "Excel文件 (*.xls)|*.xls|文本文件 (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            '认为第0列，为系统编号。设置其列保存时为string
            cgrdInfo.ColDataType(cgrdInfo.ColIndex("编号")) = flexDTString
            cgrdInfo.SaveGrid lstrFile, flexFileExcel, True   '导出excel系统编号为数字
            'cgrdInfo.SaveGrid lstrFile, flexFileTabText, True
        End If
        MsgBox "导出完成"
        
    Case "删除"
        If MsgBox("你确认要删除该条记录吗？", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
            dafuncGetData ("delete from 职业病体检_评残信息表 where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
            Com查询_Click
            MsgBox "已成功删除该条记录。"
        End If
        
    Case "保存"
'        Dim lobj As Object
        Set lobj = dafuncGetData("select * from 职业病体检_评残信息表 where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
        If lobj.RecordCount > 0 Then
            dafuncGetData ("update 职业病体检_评残信息表 set 身份证号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("身份证号")) & "',体检日期='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("体检时间")) & "',体检结果='" & Text结果.Text & "',  体检结论='" & Text结论.Text & "',提供资料信息='" & Text资料.Text & "',评残等级='" & Text等级.Text & "',评残时间='" & DTPicker时间.Value & "',评残状态='2' where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
            Com查询_Click
            MsgBox "保存成功！"
        Else
            dafuncGetData ("insert into 职业病体检_评残信息表(系统编号,身份证号,体检日期,体检结果,体检结论,提供资料信息,评残等级,评残时间,评残状态) values('" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'," & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("身份证号")) & ",'" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("体检时间")) & "','" & Text结果.Text & "','" & Text结论.Text & "','" & Text资料.Text & "','" & Text等级.Text & "','" & DTPicker时间.Value & "','2')")
            Com查询_Click
            MsgBox "保存成功！"
        End If
        
    Case "取消保存"
        If MsgBox("你确认要取消该保存结论吗？", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
            dafuncGetData ("update 职业病体检_评残信息表 set 评残状态='1' where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
            Com查询_Click
            MsgBox "已取消保存"
        End If
'        dafuncGetData ("update 职业病体检_评残信息表 set 评残状态='1' where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
    
    Case "复核"
            dafuncGetData ("update 职业病体检_评残信息表 set 评残状态='3',是否已评残='1',历史评残等级='" & Text等级.Text & "',历史评残时间='" & DTPicker时间.Value & "',历史资料信息='" & Text资料.Text & "'where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
            Dim detail As Object
            Set detail = dafuncGetData("select * from 职业病体检_评残详细历史信息表 where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
            If detail.RecordCount > 0 Then
                dafuncGetData ("update 职业病体检_评残详细历史信息表 set 身份证号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("身份证号")) & "',体检日期='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("体检时间")) & "',体检结果='" & Text结果.Text & "',  体检结论='" & Text结论.Text & "',提供资料信息='" & Text资料.Text & "',评残等级='" & Text等级.Text & "',评残时间='" & DTPicker时间.Value & "' where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
            Else
                dafuncGetData ("insert into 职业病体检_评残详细历史信息表(身份证号,系统编号,体检时间,体检结果,体检结论,资料信息,评残等级,评残时间) values('" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("身份证号")) & "'," & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & ",'" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("体检时间")) & "','" & Text结果.Text & "','" & Text结论.Text & "','" & Text资料.Text & "','" & Text等级.Text & "','" & DTPicker时间.Value & "')")
            End If
            Com查询_Click
    
    Case "取消复核"
        If MsgBox("你确认要取消该复核记录吗？", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
            dafuncGetData ("update 职业病体检_评残信息表 set 评残状态='2',是否已评残='0' where 系统编号='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("编号")) & "'")
            Com查询_Click
            MsgBox "已取消复核"
        End If
        
    Case "退出"
        Unload Me
End Select
End Sub
'退出窗体时，清空部分变量
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub
'初始化下拉列表值
Sub listcombox()
'1人员类型
    Combo类型.Text = ""
    Combo类型.AddItem "8023部队": Combo类型.ItemData(Combo类型.NewIndex) = 0
    Combo类型.AddItem "涉核部队": Combo类型.ItemData(Combo类型.NewIndex) = 1
    Combo类型.AddItem "涉核部队YK": Combo类型.ItemData(Combo类型.NewIndex) = 2
    Combo类型.ListIndex = 0
 '体检类别
    Combo类别.Text = ""
    Combo类别.AddItem "初检": Combo类别.ItemData(Combo类别.NewIndex) = 0
    Combo类别.AddItem "复检": Combo类别.ItemData(Combo类别.NewIndex) = 1
    Combo类别.ListIndex = 1
End Sub
Sub subClear()
    '清空各个文本框
    Text结果.Text = ""
    Text结论.Text = ""
    Text资料.Text = ""
    Text等级.Text = ""
    DTPicker时间.Value = Now
End Sub
