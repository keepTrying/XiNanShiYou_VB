VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm直接收费 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "直接收费"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   12030
   ClipControls    =   0   'False
   Icon            =   "frm直接收费.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "票据号"
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   120
      TabIndex        =   52
      Top             =   8040
      Width           =   3615
      Begin VB.TextBox ctxt票据号 
         Height          =   375
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   53
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label clblCurNoArea 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   58
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "当前号段："
         Height          =   180
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "当前票号："
         Height          =   180
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "基本信息"
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   120
      TabIndex        =   35
      Top             =   720
      Width           =   11685
      Begin VB.ComboBox ccmb开户行 
         Height          =   300
         Left            =   5520
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox ccmb对应业务 
         Height          =   300
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton ccmd定位 
         Caption         =   "..."
         Height          =   375
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox ccmb片区 
         Height          =   300
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ccmb卫生种类 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox ccmb主管科室 
         Height          =   300
         Left            =   9480
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   3
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   2
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   1320
      End
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H00F0F0F0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开户行"
         Height          =   180
         Index           =   0
         Left            =   4800
         TabIndex        =   54
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "类型"
         Height          =   180
         Left            =   9000
         TabIndex        =   51
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "片区"
         Height          =   180
         Left            =   2760
         TabIndex        =   41
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卫生种类"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主管科室"
         Height          =   180
         Index           =   4
         Left            =   8640
         TabIndex        =   39
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费单位"
         Height          =   180
         Index           =   3
         Left            =   4680
         TabIndex        =   38
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费人"
         Height          =   180
         Index           =   2
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费编号"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   36
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   9840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      Caption         =   "费用计算"
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   3840
      TabIndex        =   18
      Top             =   8040
      Width           =   8010
      Begin VB.ComboBox cmb交费方式 
         Height          =   300
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   810
         Width           =   1365
      End
      Begin VB.TextBox ctxtInput 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-M-d"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   24
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   23
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   810
         Width           =   1455
      End
      Begin VB.TextBox ctxtInput 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   6360
         MaxLength       =   12
         TabIndex        =   11
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox ctxtInput 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   21
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   300
         Width           =   1365
      End
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   20
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费日期"
         Height          =   180
         Index           =   25
         Left            =   150
         TabIndex        =   30
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费方式"
         Height          =   180
         Index           =   24
         Left            =   2565
         TabIndex        =   29
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "找补金额"
         Height          =   180
         Index           =   23
         Left            =   5505
         TabIndex        =   26
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收金额"
         Height          =   180
         Index           =   22
         Left            =   5520
         TabIndex        =   25
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应收金额大写"
         Height          =   180
         Index           =   21
         Left            =   2565
         TabIndex        =   23
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应收金额"
         Height          =   180
         Index           =   20
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "打折情况"
      Height          =   930
      Left            =   1440
      TabIndex        =   17
      Top             =   8280
      Visible         =   0   'False
      Width           =   1920
      Begin VB.TextBox ctxtInput 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   300
         Index           =   19
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "1.00"
         Top             =   225
         Width           =   480
      End
      Begin VB.CheckBox cchk打印打折比率 
         Caption         =   "打印打折比率"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.UpDown cupd修改打折比率 
         Height          =   360
         Left            =   1500
         TabIndex        =   20
         Top             =   195
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打折比率"
         Height          =   180
         Index           =   19
         Left            =   165
         TabIndex        =   31
         Top             =   285
         Width           =   720
      End
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   979
      ButtonWidth     =   1455
      ButtonHeight    =   926
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchk预览 
         Caption         =   "打印前预览"
         Height          =   255
         Left            =   9120
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "费用修改清单 "
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   120
      TabIndex        =   33
      Top             =   1920
      Width           =   11700
      Begin VB.Frame Frame1 
         Caption         =   "双击选择收费项目！"
         Height          =   5595
         Left            =   5880
         TabIndex        =   42
         Top             =   120
         Width           =   5775
         Begin VB.ListBox clst收费标准 
            Height          =   4200
            Left            =   120
            TabIndex        =   55
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox ctxt数量 
            Height          =   270
            Left            =   4800
            TabIndex        =   10
            Top             =   5160
            Width           =   735
         End
         Begin VB.TextBox ctxt单价 
            Height          =   270
            Left            =   3360
            TabIndex        =   9
            Top             =   5160
            Width           =   735
         End
         Begin VB.TextBox ctxt收费项目 
            Height          =   270
            Left            =   1200
            TabIndex        =   8
            Top             =   5160
            Width           =   1335
         End
         Begin VB.ComboBox Ccbo收费项目大类 
            Height          =   300
            Left            =   2760
            TabIndex        =   44
            Top             =   600
            Width           =   2775
         End
         Begin VB.ListBox clst收费项目 
            Height          =   3840
            Left            =   2760
            TabIndex        =   43
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label clblName 
            Height          =   180
            Left            =   2520
            TabIndex        =   50
            Top             =   4800
            Width           =   1410
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "数量："
            Height          =   180
            Index           =   2
            Left            =   4200
            TabIndex        =   49
            Top             =   5160
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "单价："
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   48
            Top             =   5160
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "收费项目："
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   47
            Top             =   5160
            Width           =   900
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费标准"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Clab收费项目大类 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费项目大类"
            Height          =   180
            Left            =   2760
            TabIndex        =   45
            Top             =   360
            Width           =   1200
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   5460
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5685
         _cx             =   60368684
         _cy             =   60368287
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
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   -1  'True
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "按“Del”键可以删除当前选中的项目"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   3960
         Width           =   2970
      End
   End
End
Attribute VB_Name = "frm直接收费"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

'启动参数。
Public pstr收费编号 As String

Dim WithEvents mobj界面通用对象 As cls界面通用对象
Attribute mobj界面通用对象.VB_VarHelpID = -1


Private Const 收费_收费编号 = 1
Private Const 收费_交费人 = 2
Private Const 收费_交费单位 = 3

Private Const 打折比率 = 19
Private Const 应收金额 = 20
Private Const 应收金额大写 = 21
Private Const 实收金额 = 22
Private Const 找补金额 = 23
Private Const 交费日期 = 24

Private Const 费用清单_收费项目编号 = 0
Private Const 费用清单_收费项目名称 = 1
Private Const 费用清单_单价 = 2
Private Const 费用清单_数量 = 3
Private Const 费用清单_金额 = 4


Dim mstrUndoCount As String          '用于保存表格中原来的字符串,以便在输入不合法时能够还原
Dim mstrUndoMoney As String          '用于保存表格中原来的字符串,以便在输入不合法时能够还原
Dim mstrUndoItemName As String

Dim mcur最小单价 As Currency
Dim mcur最大单价 As Currency

Dim mstr交费单位编号 As String  '从单位定位接口得到的交费单位的编号
Dim mint交费方式编号 As Integer '交费方式的编号

Dim mcur总金额 As Currency

Dim mint打折控制 As Integer
Dim mint科目级数 As Integer
Dim msng打折比率 As Single
'何嘉新增
Dim mbln使用 As Boolean         '是否能使用系统
Dim mint是否右键 As Integer     '在交费单位文本框上是否使用了右键
Dim mstr收费编号 As String      '定义变量记录收费编号
Dim mblntemp As Boolean         '判断增加项目函数是否执行过
Dim mbln控制号段 As Boolean     '是否使用收费员号段控制功能

'修改：2002-10-17（杨春）记忆打印前预览。
Private mobj记忆  As cls用户操作记忆


Private Sub Ccbo收费项目大类_Click()
    On Error GoTo errHandler
   
    Dim lobjRec As Object            '定义变量记录数据集
    
    '根据收费项目大类名称,获取收费编号前缀
    Set lobjRec = dafuncGetData("select 收费项目编号 from 收费管理_收费项目字典表 where 收费项目名称= '" & Ccbo收费项目大类.Text & "'")
    
    '获取下级收费项目
    Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where left(收费项目编号,3)='" & Left$(lobjRec("收费项目编号"), 3) & "' and len(收费项目编号)>3")
    clst收费项目.Clear
    Do While Not lobjRec.EOF
        clst收费项目.AddItem lobjRec("收费项目名称") & " " & lobjRec("收费项目编号")
        lobjRec.MoveNext
    Loop
    Exit Sub
errHandler:
    MsgBox "获取并显示指定大类的收费项目失败！" & Error, vbOKOnly + vbExclamation, "系统提示"
End Sub

Private Sub ccmb对应业务_Click()
'    If ccmb对应业务.ListIndex = 0 Then
'        ccmb开户行.Visible = True
'    Else
'        ccmb开户行.Visible = False
'    End If
    
End Sub

Private Sub ccmb收费标准_Click()
End Sub

Private Sub ccmd定位_Click()
    Dim lrds打折信息 As Object               '单位的打折信息
    Dim lrdsTemp As Object
    
    On Error GoTo errHandler
    
    '调用单位档案的定位接口获取单位信息
    Set lrdsTemp = pobj单位定位.func单位简单定位(100, 100)
    If Not (lrdsTemp Is Nothing) Then
        If lrdsTemp.RecordCount > 0 Then
            '显示单位名称`
            ctxtInput(收费_交费单位).Text = lrdsTemp("单位名称")
            '显示卫生种类、片区
            ccmb卫生种类.Text = lrdsTemp("卫生种类")
            ccmb片区.Text = IIf(IsNull(lrdsTemp("片区")), "", lrdsTemp("片区"))
            
            '保存单位的申请编号
            mstr交费单位编号 = lrdsTemp("申请编号")
            ctxtInput(收费_交费单位).SetFocus
        End If
    End If
    
    '查询打折信息
    ctxtInput(打折比率).Text = "1.00"
    Set lrds打折信息 = dafuncGetData("select * from 收费管理_打折信息表 where 单位编号='" & mstr交费单位编号 & "'")
    If Not (lrds打折信息.EOF) Then
        If mint打折控制 > 0 Then
            ctxtInput(打折比率).Text = IIf(IsNull(lrds打折信息("打折比率")), "1.00", lrds打折信息("打折比率"))
        End If
    End If

    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ccmd定位_Click", Err.Number, Err.Description, False
End Sub

Private Sub clst收费标准_DblClick()
    Dim lrds收费标准 As Object
    Dim i As Integer
    Dim lcurMoney As Currency
    
    On Error GoTo errHandler
    
    Set lrds收费标准 = dafuncGetData("select a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,金额=a.单价*a.数量 from 收费管理_收费标准信息表 a,收费管理_收费项目字典表 b where b.收费项目编号=a.收费项目编号 and 收费标准名称='" & clst收费标准.Text & "'")
    
    If lrds收费标准.EOF Then
        sffuncMsg "收费标准中无收费项目！", sf警告
        Exit Sub
    Else
        lrds收费标准.MoveFirst
        Dim llngItemCount As Long
        For i = 0 To lrds收费标准.RecordCount - 1
            If Not func检查项目是否已选(lrds收费标准("收费项目编号")) Then
                ctxt收费项目 = lrds收费标准("收费项目编号")
                ctxt收费项目_LostFocus
                sub添加项目
                llngItemCount = llngItemCount + 1
            End If
            lrds收费标准.MoveNext
        Next
        For i = 1 To cgrdDetail.Rows - 1
            lcurMoney = Format(lcurMoney + cgrdDetail.ValueMatrix(i, 费用清单_金额), "0.00")
        Next
        mcur总金额 = lcurMoney
        ctxtInput(应收金额) = lcurMoney * Val(ctxtInput(打折比率).Text)
        ctxtInput(应收金额大写) = FuncConvertToCapsStr(Val(ctxtInput(应收金额)))
        
'        If llngItemCount = lrds收费标准.RecordCount Then
'            MsgBox "收费标准中的所有收费项目(" & llngItemCount & "条)已添加到费用清单中！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
'        ElseIf llngItemCount = 0 Then
'            MsgBox "收费标准中的所有收费项目在费用清单中已添加！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
'        Else
'            MsgBox "收费标准中部分收费项目在费用清单中已添加,其余的 " & llngItemCount & " 条已添加到费用清单！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
'        End If
    End If
                
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ccmb收费标准_Click", Err.Number, Err.Description, False
End Sub

Private Sub clst收费项目_Click()
    Dim lobjRec As Object
    On Error GoTo errHandler
   ctxt收费项目 = Right(clst收费项目.List(clst收费项目.ListIndex), Len(clst收费项目.List(clst收费项目.ListIndex)) - InStr(clst收费项目.List(clst收费项目.ListIndex), " "))
    
    Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & ctxt收费项目 & "'")
    If lobjRec.RecordCount > 0 Then
        ctxt单价 = lobjRec("单价")
        ctxt数量 = 1
        mcur最小单价 = IIf(IsNull(lobjRec("最小单价").Value), 0, lobjRec("最小单价").Value)
        mcur最大单价 = IIf(IsNull(lobjRec("最大单价").Value), 99999999, lobjRec("最大单价").Value)
        clblName.Caption = lobjRec!收费项目名称
        ctxt数量.SelStart = 0
        ctxt数量.SelLength = Len(ctxt数量)
        ctxt数量.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "clst收费项目_Click", Err.Number, Err.Description, False
          
End Sub

Private Sub ctxtInput_Change(Index As Integer)
    On Error GoTo errhandle
    Static lcurMoney As Currency
    Static lintAge As Integer
    Static lsngR As Single

    Select Case Index
        Case 应收金额
            ctxtInput(应收金额大写).Text = FuncConvertToCapsStr(Val(ctxtInput(应收金额).Text))
            ctxtInput(找补金额).Text = Format(Val(ctxtInput(实收金额).Text) - Val(ctxtInput(应收金额).Text), "0.00")
            
        Case 打折比率
            If ctxtInput(打折比率).Text = vbNullString Then ctxtInput(打折比率).Text = "1.00"
            If Val(ctxtInput(打折比率).Text) > 1 Then ctxtInput(打折比率).Text = "1.00"
            If Val(ctxtInput(打折比率).Text) < 0 Then ctxtInput(打折比率).Text = "0.00"
            If Not IsNumeric(ctxtInput(打折比率).Text) Then ctxtInput(打折比率).Text = "1.00"
            
            
            ctxtInput(应收金额).Text = mcur总金额 * Val(ctxtInput(打折比率).Text)
            
            ctxtInput(应收金额大写).Text = FuncConvertToCapsStr(Val(ctxtInput(应收金额).Text))
            ctxtInput(找补金额).Text = Format(Val(ctxtInput(实收金额).Text) - Val(ctxtInput(应收金额).Text), "0.00")
            
        Case 实收金额
            If ctxtInput(实收金额).Text = vbNullString Then ctxtInput(实收金额).Text = 0
            If Not IsNumeric(ctxtInput(实收金额).Text) Then
                ctxtInput(实收金额).Text = CStr(lcurMoney)
            Else
                lcurMoney = Val(ctxtInput(实收金额).Text)
            End If
            ctxtInput(找补金额).Text = Format(Val(ctxtInput(实收金额).Text) - Val(ctxtInput(应收金额).Text), "0.00")
            
        Case 找补金额
            If Val(ctxtInput(找补金额).Text) < 0 Then
                ctxtInput(找补金额).ForeColor = &HFF
            Else
                ctxtInput(找补金额).ForeColor = &HFF0000
            End If
    End Select
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ctxtInput_Change", Err.Number, Err.Description, False
End Sub

Private Sub ctxtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errhandle
    Select Case KeyAscii
        Case vbKeyReturn
            If Index = 实收金额 And ctlb工具栏.Buttons(1).Enabled Then
                Call mobj界面通用对象_BeforeOperate("收费", False)
            ElseIf Index = 收费_交费单位 Then
                Ccbo收费项目大类.SetFocus
            End If
        Case Else
            If Index = 收费_交费单位 Then
                ctxtInput(打折比率).Text = "1.00"
            End If
        End Select
Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ctxtInput_KeyPress", Err.Number, Err.Description, False
End Sub




Private Sub cgrdDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim lcurMoney As Currency
    
    On Error GoTo errhandle
    'ctlb工具栏.Buttons("收费(&G)").Enabled = True
    Select Case cgrdDetail.TextMatrix(0, Col)
        Case "数量"
            '判断输入的是否数值
            If Len(cgrdDetail.TextMatrix(Row, Col)) > 4 Then
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
            Else
                If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) And Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    '是数值
                    '计算金额
                    cgrdDetail.TextMatrix(Row, 费用清单_金额) = cgrdDetail.TextMatrix(Row, 费用清单_单价) * cgrdDetail.TextMatrix(Row, 费用清单_数量)
                Else
                    '不是数值
                    'Undo
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
                End If
            End If
        Case "单价"
            Dim lcur单价 As Currency
            If mcur最小单价 = mcur最大单价 Then
                sffuncMsg "该收费项目单价已定,不可修改！", sf警告
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                Exit Sub
            End If
            
            If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) Then
                If Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    If Val(cgrdDetail.TextMatrix(Row, Col)) <= mcur最大单价 And Val(cgrdDetail.TextMatrix(Row, Col)) >= mcur最小单价 Then
                        cgrdDetail.TextMatrix(Row, 费用清单_金额) = cgrdDetail.TextMatrix(Row, 费用清单_单价) * cgrdDetail.TextMatrix(Row, 费用清单_数量)
                    Else
                        sffuncMsg "输入的单价超出范围！", sf警告
                        cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                    End If
                Else
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                End If
            Else
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
            End If
        Case "收费项目名称"
            If cgrdDetail.TextMatrix(Row, Col) = "" Then
                sffuncMsg "必须输入收费项目名称！", sf警告
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoItemName
                Exit Sub
            End If
            '判断项目名称是否重复。
            For i = 1 To cgrdDetail.Rows - 1
                If i <> Row And cgrdDetail.TextMatrix(Row, Col) = cgrdDetail.TextMatrix(i, 费用清单_收费项目名称) Then
                    sffuncMsg "收费项目名称不允许重复！", sf警告
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoItemName
                    Exit Sub
                End If
            Next
        Case Else
    End Select
    
    For i = 1 To cgrdDetail.Rows - 1
        lcurMoney = lcurMoney + cgrdDetail.ValueMatrix(i, 费用清单_金额)
    Next
    mcur总金额 = lcurMoney
    
    ctxtInput(应收金额) = lcurMoney * Val(ctxtInput(打折比率).Text)
    ctxtInput(应收金额大写) = FuncConvertToCapsStr(Val(ctxtInput(应收金额)))
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "cing费用清单_AfterEdit", Err.Number, Err.Description, False
    
End Sub

Private Sub cgrdDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
     
    Select Case Col
        Case 费用清单_数量
            ctlb工具栏.Buttons("收费(&G)").Enabled = False
            mstrUndoCount = cgrdDetail.TextMatrix(Row, Col)
            
        Case 费用清单_单价
            ctlb工具栏.Buttons("收费(&G)").Enabled = False
                        
            '获取最小单价,最大单价.
            Dim lobjRec As Object
            Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & cgrdDetail.TextMatrix(Row, 0) & "'")
            If lobjRec.RecordCount > 0 Then
                mcur最小单价 = IIf(IsNull(lobjRec("最小单价").Value), 0, lobjRec("最小单价").Value)
                mcur最大单价 = IIf(IsNull(lobjRec("最大单价").Value), 99999999, lobjRec("最大单价").Value)
            Else
                sffuncMsg "未找到该收费项目的设置信息，该设置信息可能已被修改或删除，请退出收费界面，重新进入！"
            End If
            mstrUndoMoney = cgrdDetail.TextMatrix(Row, Col)
        Case 费用清单_收费项目名称
            ctlb工具栏.Buttons("收费(&G)").Enabled = False
            mstrUndoItemName = cgrdDetail.TextMatrix(Row, Col)
        Case Else
            ctlb工具栏.Buttons("收费(&G)").Enabled = True
            Cancel = True
    End Select
End Sub



Private Sub cgrdDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete
            mobj界面通用对象_BeforeOperate "删除", False
    End Select

End Sub

Private Sub cgrdDetail_LostFocus()
    On Error Resume Next
    ctlb工具栏.Buttons("收费(&G)").Enabled = True
End Sub


Private Sub clst收费项目_DblClick()
    Dim lobjRec As Object
    On Error GoTo errHandler
   ctxt收费项目 = Right(clst收费项目.List(clst收费项目.ListIndex), Len(clst收费项目.List(clst收费项目.ListIndex)) - InStr(clst收费项目.List(clst收费项目.ListIndex), " "))
    
    Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & ctxt收费项目 & "'")
    If lobjRec.RecordCount > 0 Then
        ctxt单价 = lobjRec("单价")
        ctxt数量 = 1
        clblName.Caption = lobjRec!收费项目名称
        
        '添加收费项目
        If Not func检查项目是否已选(ctxt收费项目) Then
            sub添加项目
        End If
        
        ctxt收费项目 = ""
        clblName = ""
        ctxt单价 = ""
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "clst收费项目_DblClick", Err.Number, Err.Description, False
          
End Sub

Private Sub sub添加项目()
    Dim lcurMoney As Double
    Dim i As Long
    On Error GoTo errHandler
    
    cgrdDetail.AddItem ctxt收费项目 & vbTab & clblName & vbTab & _
                    ctxt单价 & vbTab & ctxt数量 & vbTab & Format(Val(ctxt单价) * Val(ctxt数量), "0.0")
    For i = 1 To cgrdDetail.Rows - 1
        lcurMoney = Format(lcurMoney + cgrdDetail.ValueMatrix(i, 费用清单_金额), "0.00")
    Next
    mcur总金额 = lcurMoney
    ctxtInput(应收金额) = lcurMoney * Val(ctxtInput(打折比率).Text)
    ctxtInput(应收金额大写) = FuncConvertToCapsStr(Val(ctxtInput(应收金额)))

    
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "sub添加项目", Err.Number, Err.Description, True
End Sub

Private Sub ctxt单价_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 And clblName <> "" Then
        Dim lcur单价 As Currency
        If mcur最小单价 = mcur最大单价 And ctxt单价 <> mcur最小单价 Then
            ctxt单价 = mstrUndoMoney
            ctxt收费项目.SetFocus
            sffuncMsg "该收费项目单价已定，不可修改！", sf警告
            Exit Sub
        Else
            If IsNumeric(ctxt单价) Then
                If Val(ctxt单价) > 0 Then
                    If Val(ctxt单价) <= mcur最大单价 And Val(ctxt单价) >= mcur最小单价 Then
                        
                    Else
                        ctxt单价 = mstrUndoMoney
                        ctxt收费项目.SetFocus
                        sffuncMsg "输入的单价超出范围！", sf警告
                        Exit Sub
                    End If
                Else
                    ctxt单价 = mstrUndoMoney
                End If
            Else
                ctxt单价 = mstrUndoMoney
            End If
            ctxt数量.SelStart = 0
            ctxt数量.SelLength = Len(ctxt数量)
            ctxt数量.SetFocus
        End If
        
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ctxt单价_KeyUp", Err.Number, Err.Description, False
    
End Sub

Private Sub ctxt票据号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        ctxt收费项目.SetFocus
    End If
End Sub

Private Sub ctxt票据号_LostFocus()
    Dim lstr票据号 As String
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    
    If Not IsNumeric(ctxt票据号) Then
        MsgBox "票据号必须是数字！", vbInformation, "系统提示"
        ctxt票据号.SetFocus
        Exit Sub
    End If
    If mbln控制号段 Then
        '检查该票据号是否在已分配的号段内
        Set lobjRec = dafuncGetData("select 起号,止号 from 收费管理_收费员号段信息表 where '" & ctxt票据号 & "' between 起号 and 止号 and 用户编号='" & um用户编号 & "' and 是否用完='否'")
        If lobjRec.RecordCount = 0 Then
            MsgBox "您设置的当前票据号不在尚未用完的票据号段范围内，不能进行收费，请重新进行设置！", vbInformation, "系统提示"
            ctlb工具栏.Buttons(1).Enabled = False
            Exit Sub
        End If
        clblCurNoArea = lobjRec(0) & "－" & lobjRec(1)
    End If
    lstr票据号 = Format(Val(ctxt票据号) - 1, String(Len(ctxt票据号), "0"))
    dafuncGetData "update 系统管理_系统编号生成记录表 set 当前值=" & lstr票据号 & " where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'"
    ctlb工具栏.Buttons(1).Enabled = True
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ctxt票据号_LostFocus", Err.Number, Err.Description, False
End Sub

Private Sub ctxt收费项目_GotFocus()
    On Error Resume Next
    If ctxt收费项目 = "无该项目！" Then
        ctxt收费项目 = ""
    End If
End Sub

Private Sub ctxt收费项目_LostFocus()
    Dim lobjRec As Object
    Dim pint科目级数 As Integer
    On Error GoTo errHandler
    
    '根据收费项目编号获取收费项目名称。
    clblName.Caption = ""
    pint科目级数 = Val(pobj收费管理.业务设置("科目级数"))
    If pint科目级数 = 0 Then pint科目级数 = 2
    
    If ctxt收费项目 <> "" And Len(ctxt收费项目) = 3 * pint科目级数 Then
    
        Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & ctxt收费项目 & "'")
        If lobjRec.RecordCount > 0 Then
            clblName.Caption = lobjRec!收费项目名称
            ctxt单价 = lobjRec("单价")
            mstrUndoMoney = lobjRec("单价")
            mcur最小单价 = IIf(IsNull(lobjRec("最小单价").Value), 0, lobjRec("最小单价").Value)
            mcur最大单价 = IIf(IsNull(lobjRec("最大单价").Value), 99999999, lobjRec("最大单价").Value)

            If ctxt数量 = "" Then ctxt数量 = 1
        Else
            ctxt收费项目 = "无该项目！"
        End If
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ctxt收费项目_LostFocus", Err.Number, Err.Description, False
    
End Sub

Private Sub ctxt数量_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        If clblName <> "" And ctxt收费项目 <> "" Then
            If Not func检查项目是否已选(ctxt收费项目.Text) Then
                sub添加项目
            End If
        End If
        ctxt收费项目 = ""
        clblName = ""
        ctxt单价 = ""
        ctxt收费项目.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "ctxt数量_KeyDown", Err.Number, Err.Description, False

End Sub

Private Sub cupd修改打折比率_DownClick()
    On Error Resume Next
    If Val(ctxtInput(打折比率).Text) > 0 Then
        ctxtInput(打折比率).Text = Format(CStr(Val(ctxtInput(打折比率).Text) - 0.01), "0.00")
    Else
        ctxtInput(打折比率).Text = "0.00"
    End If
End Sub

Private Sub cupd修改打折比率_UpClick()
On Error GoTo errhandle
    If Val(ctxtInput(打折比率).Text) < 1 Then
        ctxtInput(打折比率).Text = Format(CStr(Val(ctxtInput(打折比率).Text) + 0.01), "0.00")
    Else
        ctxtInput(打折比率).Text = "1.00"
    End If
Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "Form_UpClick", Err.Number, Err.Description, False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And ActiveControl.Name <> "ctxt单价" And ActiveControl.Name <> "ctxt数量" Then
        SendKeys Chr(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim lcol工具栏 As Collection
    Dim i As Long, lobjRec As Recordset
    On Error GoTo errhandle
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
    mbln使用 = True
        
    Set mobj界面通用对象 = New cls界面通用对象
    Set mobj界面通用对象.Form = Me
    Set mobj界面通用对象.c工具栏 = ctlb工具栏
    
    Set lcol工具栏 = New Collection
    
    lcol工具栏.Add "收费(&G)101"
    lcol工具栏.Add "|"
    lcol工具栏.Add "删除"
    lcol工具栏.Add "清空"
    lcol工具栏.Add "|"
    lcol工具栏.Add "退出"
    
    mobj界面通用对象.subInitialize lcol工具栏, ""
    
    mint打折控制 = Val(pobj收费管理.业务设置("打折控制"))
    mint科目级数 = Val(pobj收费管理.业务设置("科目级数"))
    
    sub初始化窗体
    
    If pstr收费编号 <> "" Then
        '内部收费,显示费用信息。
        sub显示费用信息
        
        '没有权限修改，则不能修改费用信息。
        If Not umfunc校验用户权限("收费管理_内部收费信息修改") Then
            Frame4.Enabled = False
            Frame1.Enabled = False
            cgrdDetail.Editable = False
'            Label4.Caption = "你没有权限修改内部收费信息！"
        End If
    End If
    ctxt票据号 = func获取票据号()
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
Private Function func获取票据号() As String
    Dim lstr票据号 As String
    Dim lLen As Integer
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select 当前值 from 系统管理_系统编号生成记录表 where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'")
    If mbln控制号段 Then
        If lobjRec.RecordCount = 0 Then
            '找出下一个未使用的最小号段
            Set lobjRec = dafuncGetData("select 起号,止号 from 收费管理_收费员号段信息表 where 用户编号='" & um用户编号 & "' and 是否用完='否' order by 起号")
            If lobjRec.RecordCount = 0 Then
                MsgBox "您当前没有尚未用完的票据号段信息，不能进行收费！", vbInformation, "系统提示"
                ctlb工具栏.Buttons(1).Enabled = False
                func获取票据号 = ""
                Exit Function
            Else
                clblCurNoArea = lobjRec(0) & "－" & lobjRec(1)
                lstr票据号 = lobjRec(0)
                lLen = Len(lstr票据号)
                dafuncGetData "insert into 系统管理_系统编号生成记录表(业务名称,编号名称,数据类型,当前值,长度,是否按年重编,当前年号) values('收费管理" & um用户编号 & "','收据号','C'," & Format(Val(lstr票据号) - 1, String(lLen, "0")) & ",9,'否',2008)"
    '            dafuncGetData "update 系统管理_系统编号生成记录表 set 当前值='" & Format(Val(lstr票据号) - 1, String(lLen, "0")) & "' where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'"
            End If
        Else
            lstr票据号 = IIf(IsNull(lobjRec(0)), "0", lobjRec(0))
            lLen = Len(lstr票据号)
            '检查该票据号是否正确
            Set lobjRec = dafuncGetData("select 起号,止号 from 收费管理_收费员号段信息表 where '" & lstr票据号 & "' between 起号 and 止号 and 用户编号='" & um用户编号 & "' and 是否用完='否'")
            If lobjRec.RecordCount = 0 Then
                '检查该号是否是新的号段的起始号
                Set lobjRec = dafuncGetData("select 起号,止号 from 收费管理_收费员号段信息表 where '" & Format(Val(lstr票据号) + 1, String(lLen, "0")) & "' between 起号 and 止号 and 用户编号='" & um用户编号 & "' and 是否用完='否'")
                If lobjRec.RecordCount = 0 Then
                    MsgBox "您设置的当前票据号不在尚未用完的票据号段范围内，不能进行收费，请重新进行设置！", vbInformation, "系统提示"
                    ctlb工具栏.Buttons(1).Enabled = False
                    func获取票据号 = ""
                    Exit Function
                Else
                    clblCurNoArea = lobjRec(0) & "－" & lobjRec(1)
                End If
            Else
                clblCurNoArea = lobjRec(0) & "－" & lobjRec(1)
            End If
            '检查该号段是否已经用完，需要换发票了
            Set lobjRec = dafuncGetData("select ID from 收费管理_收费员号段信息表 where 用户编号='" & um用户编号 & "' and 是否用完='否' and 止号='" & lstr票据号 & "'")
            If lobjRec.RecordCount Then
                '将该号段设为已用完
                dafuncGetData "update 收费管理_收费员号段信息表 set 是否用完='是' where ID=" & lobjRec(0)
                '找出下一个未使用的最小号段
                Set lobjRec = dafuncGetData("select 起号,止号 from 收费管理_收费员号段信息表 where 用户编号='" & um用户编号 & "' and 是否用完='否' order by 起号")
                If lobjRec.RecordCount = 0 Then
                    MsgBox "您当前没有尚未用完的票据号段信息，不能进行收费！", vbInformation, "系统提示"
                    ctlb工具栏.Buttons(1).Enabled = False
                    '清除其当前票据号设置，让其重新开始
                    dafuncGetData "delete 系统管理_系统编号生成记录表 where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'"
                    func获取票据号 = ""
                    Exit Function
                Else
                    clblCurNoArea = lobjRec(0) & "－" & lobjRec(1)
                    lstr票据号 = lobjRec(0)
                    lLen = Len(lstr票据号)
                    dafuncGetData "update 系统管理_系统编号生成记录表 set 当前值=" & Format(Val(lstr票据号) - 1, String(lLen, "0")) & " where 业务名称='收费管理" & um用户编号 & "' and 编号名称='收据号'"
                    MsgBox "当前票据号段已经用完，请在打印机上安装正确的新票据！", vbInformation, "系统提示"
                End If
            Else
                lstr票据号 = Format(Val(lstr票据号) + 1, String(lLen, "0"))
                '检查收据号是否重复
                Set lobjRec = dafuncGetData("select * from 收费管理_费用信息表 where 收据号='" & lstr票据号 & "'")
                If lobjRec.RecordCount Then
                    MsgBox "系统中已经存在该票据号了，请注意检查！", vbInformation, "系统提示"
                End If
            End If
        End If
    Else
        If lobjRec.RecordCount = 0 Then
            lstr票据号 = "1"
        ElseIf ctxt票据号 = "" Then
            lstr票据号 = Format(Val(lobjRec(0)) + 1, String(Len(lobjRec(0)), "0"))
        Else
            lstr票据号 = Format(Val(ctxt票据号) + 1, String(Len(ctxt票据号), "0"))
        End If
        '检查收据号是否重复
        Set lobjRec = dafuncGetData("select * from 收费管理_费用信息表 where 收据号='" & lstr票据号 & "'")
        If lobjRec.RecordCount Then
            MsgBox "系统中已经存在该票据号了，请注意检查！", vbInformation, "系统提示"
        End If
    End If
    func获取票据号 = lstr票据号
End Function

Private Sub sub显示费用信息()
    Dim lobjRec As Object
    Dim i As Long
    
    On Error GoTo errHandler
    
    If pstr收费编号 <> "" Then
        '修改收费记录。
        Set lobjRec = dafuncGetData("select a.收费批号,a.收费编号,a.收费项目编号,收费项目名称=(select 收费项目名称 from 收费管理_收费项目字典表 where 收费项目编号=a.收费项目编号),a.数量,计量单位=(select 计量单位 from 收费管理_收费项目字典表 where 收费项目编号=a.收费项目编号),a.单价,a.金额,a.收费状态,a.交费方式,a.交费人,a.交费单位编号,交费单位名称 ,a.交费日期,a.退费日期,收费人编号=a.收费人,收费人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.收费人),退费人编号=a.退费人,退费人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.退费人) ,主管科室经手人编号=a.主管科室经手人,主管科室经手人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.主管科室经手人),主管科室编号,主管科室=(select 名称 from 系统管理_科室字典表 where 编号=a.主管科室编号),打折比率,备注1,备注2  from 收费管理_费用信息表 a where 收费编号='" & pstr收费编号 & "'  and 收费状态=0")
        
        cgrdDetail.Rows = 1
    
        Do While Not lobjRec.EOF
            cgrdDetail.AddItem lobjRec("收费项目编号") & vbTab & _
                lobjRec("收费项目名称") & vbTab & _
                lobjRec("单价") & vbTab & _
                lobjRec("数量") & vbTab & _
                lobjRec("金额")
            lobjRec.MoveNext
        Loop
        If lobjRec.RecordCount > 0 Then
            lobjRec.MoveFirst
            mstr交费单位编号 = IIf(IsNull(lobjRec("交费单位编号").Value), "", lobjRec("交费单位编号").Value)
        
            ctxtInput(收费_收费编号).Text = lobjRec("收费编号")
            
            If IIf(IsNull(lobjRec("主管科室")), "", lobjRec("主管科室")) <> "" Then
                For i = 0 To ccmb主管科室.ListCount - 1
                    If ccmb主管科室.List(i) = IIf(IsNull(lobjRec("主管科室")), "", lobjRec("主管科室")) Then
                        ccmb主管科室.ListIndex = i
                        Exit For
                    End If
                Next
            Else
                ccmb主管科室.ListIndex = -1
            End If
            
            ccmb卫生种类.Text = IIf(IsNull(lobjRec("备注1").Value), "", lobjRec("备注1").Value)
            ccmb片区.Text = IIf(IsNull(lobjRec("备注2").Value), "", lobjRec("备注2").Value)
        
        
            ctxtInput(收费_交费人).Text = lobjRec("交费人")
            ctxtInput(收费_交费单位).Text = IIf(IsNull(lobjRec("交费单位名称").Value), "", lobjRec("交费单位名称").Value)
        
            Set lobjRec = dafuncGetData("select 打折比率 from 收费管理_打折信息表 where 单位编号='" & mstr交费单位编号 & "'")
            
            If lobjRec.EOF Then
                ctxtInput(打折比率).Text = "1.00"
            Else
                ctxtInput(打折比率).Text = Format(lobjRec("打折比率").Value, "0.00")
            End If
            
        End If
        
        Dim lcurMoney As Currency
        lcurMoney = 0
        For i = 1 To cgrdDetail.Rows - 1
            lcurMoney = lcurMoney + cgrdDetail.ValueMatrix(i, 费用清单_金额)
        Next
        mcur总金额 = lcurMoney
        
        ctxtInput(应收金额) = lcurMoney * Val(ctxtInput(打折比率).Text)
        ctxtInput(应收金额大写) = FuncConvertToCapsStr(Val(ctxtInput(应收金额)))
        
    End If

    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm直接收费", "sub显示费用信息", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mobj界面通用对象 = Nothing
    
End Sub


Private Sub sub清除界面()
    Dim i As Integer
    
    On Error GoTo errhandle
    mcur总金额 = 0
    
    '在保存后需要清空收费编号;徐冀川;2002/9/30
    mstr交费单位编号 = ""
  
    Dim lobjCtrl As Control
    For Each lobjCtrl In ctxtInput
        lobjCtrl.Text = ""
    Next
    cgrdDetail.Rows = 1
    
    ctxtInput(打折比率).Text = "1.00"
    
    For i = 应收金额 To 交费日期
        ctxtInput(i).Text = ""
    Next
    
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select getdate()")
    ctxtInput(交费日期).Text = Format(lobjRec(0), "yyyy-mm-dd")
    
'    ctxtInput(交费日期).Text = Date
    ctxtInput(打折比率).Text = "1.00"
    ccmb主管科室.Text = um用户所属科室
    
    If ctxtInput(收费_收费编号).Enabled Then
        ctxtInput(收费_收费编号).SetFocus
    Else
        ctxtInput(收费_交费人).SetFocus
    End If
    
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "sub清除收费界面", Err.Number, Err.Description, True
End Sub


Private Sub sub初始化窗体()
    
    On Error GoTo errhandle
    
    Dim lobj收费标准 As Object
    Dim lobj科室 As Object
    Dim lobj交费方式 As Object

    mstrUndoCount = ""
    mstrUndoMoney = ""
    mstr交费单位编号 = ""
    mint交费方式编号 = 0
    mcur总金额 = 0
    
    Dim i As Long
    Dim j As Long
    
    Set lobj收费标准 = dafuncGetData("select 收费标准名称,助记符 from 收费管理_收费标准信息表 group by 助记符,收费标准名称")
    Set lobj科室 = dafuncGetData("select * from 系统管理_科室字典表")
    Set lobj交费方式 = dafuncGetData("select * from 收费管理_交费方式字典表")
    
    
    '初始化 "cing费用清单"
    With cgrdDetail
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, 费用清单_收费项目编号) = "收费项目编号"
        .ColWidth(费用清单_收费项目编号) = 1310
        .ColAlignment(费用清单_收费项目编号) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_收费项目名称) = "收费项目名称"
        .ColWidth(费用清单_收费项目名称) = 1320
        
        .TextMatrix(0, 费用清单_单价) = "单价"
        .ColWidth(费用清单_单价) = 480
        
        .TextMatrix(0, 费用清单_数量) = "数量"
        .ColWidth(费用清单_数量) = 500
        
        .TextMatrix(0, 费用清单_金额) = "金额"
        .ColWidth(费用清单_金额) = 570
    End With
    
    '初始化 "收费标准"
    clst收费标准.Clear
    Do While Not lobj收费标准.EOF
        clst收费标准.AddItem lobj收费标准("收费标准名称").Value
        lobj收费标准.MoveNext
    Loop
    
    '初始化 "主管科室"列表
    ccmb主管科室.Clear
    If Not (lobj科室 Is Nothing) Then
        Do While Not lobj科室.EOF
            ccmb主管科室.AddItem lobj科室("名称").Value
            ccmb主管科室.ItemData(ccmb主管科室.ListCount - 1) = "1" & lobj科室!编号
            lobj科室.MoveNext
        Loop
    End If
    
    ccmb主管科室.ListIndex = -1
    
    
    If umfunc校验用户权限("收费管理_打折") Then
        Frame6.Enabled = True
        Frame6.Caption = "打折情况"
        lblCaption(19).Enabled = True
        ctxtInput(19).Enabled = True
        cupd修改打折比率.Enabled = True
        cchk打印打折比率.Enabled = True
        Select Case mint打折控制
            Case 0
                cupd修改打折比率.Enabled = False
                ctxtInput(19).Enabled = False
            Case 1
                cupd修改打折比率.Enabled = True
                ctxtInput(19).Enabled = True
            Case 2
                cupd修改打折比率.Enabled = False
                ctxtInput(19).Enabled = False
            Case Else
        End Select
        
    Else
        Frame6.Caption = "打折情况(无权限)"
        Frame6.Enabled = False
        lblCaption(19).Enabled = False
        ctxtInput(打折比率).Enabled = False
        cupd修改打折比率.Enabled = False
        cchk打印打折比率.Enabled = False
    End If
        
    '初始化交费方式列表
    If Not (lobj交费方式 Is Nothing) Then
        Do While Not lobj交费方式.EOF
            cmb交费方式.AddItem lobj交费方式("名称").Value
            cmb交费方式.ItemData(cmb交费方式.ListCount - 1) = "1" & lobj交费方式("编号")
            lobj交费方式.MoveNext
        Loop
        cmb交费方式.ListIndex = 0
    End If
    
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select getdate()")
    ctxtInput(交费日期).Text = Format(lobjRec(0), "yyyy-mm-dd")
    
    '获取收费项目大类。
    Set lobjRec = dafuncGetData("select 收费项目编号,收费项目名称 from 收费管理_收费项目字典表 where len(收费项目编号)=3  order by 收费项目编号 ")
    Do While Not lobjRec.EOF
        Ccbo收费项目大类.AddItem lobjRec("收费项目名称")
        lobjRec.MoveNext
    Loop
    
    Ccbo收费项目大类.ListIndex = 0
    
    '获取卫生种类
    Set lobjRec = dafuncGetData("select * from 系统管理_卫生种类字典视图 order by 编号")
    ccmb卫生种类.Clear
    ccmb卫生种类.AddItem ""
    Do While Not lobjRec.EOF
        ccmb卫生种类.AddItem lobjRec("名称").Value
        lobjRec.MoveNext
    Loop
    
    '获取片区
    Set lobjRec = dafuncGetData("select * from 系统管理_片区字典视图 order by 编号")
    ccmb片区.Clear
    ccmb片区.AddItem ""
    Do While Not lobjRec.EOF
        ccmb片区.AddItem lobjRec("名称").Value
        lobjRec.MoveNext
    Loop
    
    '获取开户行帐号.
    Set lobjRec = dafuncGetData("select 开户行+' '+帐号 from 收费管理_银行开户行设置表")
    ccmb开户行.Clear
    Do While Not lobjRec.EOF
        ccmb开户行.AddItem lobjRec(0)
        
        lobjRec.MoveNext
    Loop
    If ccmb开户行.ListCount > 0 Then
        ccmb开户行.ListIndex = 0
    End If
    
    ccmb对应业务.AddItem "一般", 0
    ccmb对应业务.AddItem "门诊", 1
    ccmb对应业务.ListIndex = 0
    
    '获取是否使用收费员票据的号段控制功能.
    Set lobjRec = dafuncGetData("select 设置值 from 收费管理_业务配置表 where 设置项目='控制号段'")
    If lobjRec.RecordCount = 0 Then
        mbln控制号段 = False
    ElseIf lobjRec(0) = "0" Then
        mbln控制号段 = False
    Else
        mbln控制号段 = True
    End If
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "sub初始化窗体", Err.Number, Err.Description, True
End Sub


Private Sub mobj界面通用对象_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long, j As Long
    Dim lobjRec As Recordset
    
    On Error GoTo errhandle
    Select Case Operate
        Case "收费"
            '校验数据合法性。
            If Not ValidateData Then Exit Sub
            
            If ctxt票据号 = "" Then
                MsgBox "票据号设置不正确，不能收费！", vbInformation, "系统提示"
                Exit Sub
            End If
            '检查收据号是否重复
            Set lobjRec = dafuncGetData("select * from 收费管理_费用信息表 where 收据号='" & ctxt票据号 & "'")
            If lobjRec.RecordCount Then
                MsgBox "系统中已经存在该票据号了，请重新录入其它票据号！", vbInformation, "系统提示"
                'ctxt票据号.SetFocus
                Exit Sub
            End If
            
            mint交费方式编号 = Right(cmb交费方式.ItemData(cmb交费方式.ListIndex), Len(Trim(Str(cmb交费方式.ItemData(cmb交费方式.ListIndex)))) - 1)
            
            '收集要保存的费用信息。
            Dim lstr主管科室编号 As String
            Dim lcol记录 As Collection
            Dim lcol数据 As Collection
            Dim lstr收费编号 As String
            
            If ccmb主管科室.ListIndex >= 0 Then
                lstr主管科室编号 = ccmb主管科室.ItemData(ccmb主管科室.ListIndex)
                lstr主管科室编号 = Right(lstr主管科室编号, Len(lstr主管科室编号) - 1)
            Else
                lstr主管科室编号 = um用户所属科室编号
            End If
            Set lcol数据 = New Collection
            For i = 1 To cgrdDetail.Rows - 1
                Set lcol记录 = New Collection
                For j = 0 To cgrdDetail.Cols - 1
                    lcol记录.Add cgrdDetail.TextMatrix(i, j), cgrdDetail.TextMatrix(0, j)
                Next
                '添加收费其他字段
                lcol记录.Add ctxtInput(收费_交费人).Text, "交费人"
                lcol记录.Add mstr交费单位编号, "交费单位编号"
                lcol记录.Add ctxtInput(收费_交费单位).Text, "交费单位名称"
                lcol记录.Add lstr主管科室编号, "主管科室编号"
                lcol记录.Add um用户编号, "主管科室经手人"
                lcol记录.Add ccmb卫生种类.Text, "备注1"
                lcol记录.Add ccmb片区.Text, "备注2"
                lcol数据.Add lcol记录
            Next
            
            '保存划价信息。
            lstr收费编号 = pobj收费管理.func划价保存(lcol数据, pstr收费编号)
            
            '保存收费确信信息。
            Dim lcol收费编号集 As Collection
            
            Set lcol收费编号集 = New Collection
            lcol收费编号集.Add lstr收费编号
            
            Dim lcol收费确认信息 As Collection
            
            Set lcol收费确认信息 = New Collection
            With lcol收费确认信息
                .Add Val(ctxtInput(打折比率).Text), "打折比率"
                .Add mint交费方式编号, "收费方式"
                .Add CDate(ctxtInput(交费日期).Text), "交费日期"
                .Add um用户编号, "收费人"
                
                '2006-5-15
                If ccmb开户行.Visible Then
                    .Add ccmb开户行.Text, "开户银行"
                Else
                    .Add "", "开户银行"
                End If
                .Add ccmb对应业务.Text, "对应业务"
                
            End With
            
            Call pobj收费管理.sub收费确认(lcol收费编号集, lcol收费确认信息)
            
            mcur总金额 = 0
            ctxtInput(应收金额) = "0"
            
            sub清除界面
                        
            '打印票据。
            'Call func录入票据号
            
            pobj收费管理.sub打印票据 lstr收费编号, IIf(cchk预览.Value = 1, True, False), True
            
            ctxt票据号 = func获取票据号()
            If ctxt票据号 <> "" And mbln控制号段 Then
                If CLng(ctxt票据号) < 100 Then MsgBox "票据号被复位，请检查是否正确！", vbInformation, "系统提示"
            End If
            ctxtInput(收费_交费单位).SetFocus
            
        Case "删除"
            Dim lcurMoney As Currency
            
            If cgrdDetail.Row > 0 Then
                cgrdDetail.RemoveItem cgrdDetail.Row
                For i = 1 To cgrdDetail.Rows - 1
                    lcurMoney = lcurMoney + cgrdDetail.ValueMatrix(i, 费用清单_金额)
                Next
                mcur总金额 = lcurMoney
                
                ctxtInput(应收金额) = lcurMoney * Val(ctxtInput(打折比率).Text)
                ctxtInput(应收金额大写) = FuncConvertToCapsStr(Val(ctxtInput(应收金额)))
            End If
            
        Case "清空"
            sub清除界面
            
        Case Else
    End Select
    Exit Sub
    
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", "mobj界面通用对象_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Function func检查项目是否已选(ByVal para收费项目编号 As String) As Boolean
On Error GoTo errhandle
    Dim i As Long
    func检查项目是否已选 = False
    If cgrdDetail.Rows = 1 Then
        func检查项目是否已选 = False
        Exit Function
    End If
    
    For i = 1 To cgrdDetail.Rows - 1
        If para收费项目编号 = cgrdDetail.TextMatrix(i, 费用清单_收费项目编号) Then
            func检查项目是否已选 = True
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", " func检查项目是否已选()", Err.Number, Err.Description
End Function

Private Function ValidateData() As Boolean
    On Error GoTo errhandle
    ValidateData = False
    If ctxtInput(收费_交费人).Text = vbNullString And ctxtInput(收费_交费单位) = vbNullString Then
        sffuncMsg """交费人"" 和 ""交费单位"" 必须输入其中之一！", sf警告
        Exit Function
    End If
    If cgrdDetail.Rows = 1 Then
        sffuncMsg "无费用信息可以保存！", sf警告
        Exit Function
    End If
    
    If IsNumeric(ctxtInput(19).Text) Then
        If CDbl(ctxtInput(19).Text) = 0 Then
            sffuncMsg "打折比率不能为0！", sf警告
            ctxtInput(19).Text = "1.00"
            Exit Function
        End If
    Else
        sffuncMsg "打折比率录入不正确！", sf警告
        Exit Function
    End If
    ValidateData = True
    Exit Function
errhandle:
    sfsub错误处理 "收费界面部件", "frm直接收费", " ValidateData()", Err.Number, Err.Description, True
End Function


