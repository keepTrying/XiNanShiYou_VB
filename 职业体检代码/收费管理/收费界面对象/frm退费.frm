VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "录入控件.ocx"
Begin VB.Form frm退费 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "其它管理"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Cchk退费打印标识 
      Caption         =   "退费时打印票据"
      Height          =   255
      Left            =   9000
      TabIndex        =   31
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox ctxt提示 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   60
      TabIndex        =   29
      Text            =   "请稍后..."
      Top             =   7515
      Visible         =   0   'False
      Width           =   900
   End
   Begin MSComctlLib.ProgressBar cprg进度 
      Height          =   120
      Left            =   960
      TabIndex        =   28
      Top             =   7575
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Max             =   50
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   4875
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      TabIndex        =   13
      Top             =   6915
      Width           =   10995
      Begin VB.Timer ctmr定时 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   165
         Top             =   240
      End
      Begin VB.TextBox ctxt退费人 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6420
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox ctxt退费批号 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   270
         Width           =   1905
      End
      Begin VB.TextBox ctxt总金额 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         Height          =   330
         Left            =   4125
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   285
         Width           =   1335
      End
      Begin VB.TextBox ctxt退费日期 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   1365
      End
      Begin 录入控件.ctlInputBox cinb退费输入 
         Height          =   360
         Index           =   2
         Left            =   5805
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   580
         Text            =   ""
         Label           =   "退费人"
         Enabled         =   0   'False
         名称            =   ""
         长度            =   0
         允许等于最大值  =   0   'False
         允许等于最小值  =   0   'False
         允许多选        =   0   'False
      End
      Begin MSComCtl2.DTPicker cdtp日期 
         Height          =   345
         Index           =   2
         Left            =   3300
         TabIndex        =   9
         Top             =   735
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   36951
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "退费人"
         Height          =   210
         Left            =   5640
         TabIndex        =   25
         Top             =   375
         Width           =   600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "收费批号"
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "总金额"
         Height          =   180
         Left            =   3330
         TabIndex        =   21
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退费日期"
         Height          =   240
         Left            =   8145
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar cstu状态栏 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   7380
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "退费"
            TextSave        =   "退费"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "费用信息"
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   11055
      Begin VSFlex6Ctl.vsFlexGrid cgrd费用信息 
         Height          =   4245
         Left            =   75
         TabIndex        =   10
         Top             =   540
         Width           =   10815
         _cx             =   23743108
         _cy             =   23731520
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648447
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "收费批号        |收费编号       |收费项目     |数量   |金额    |交费人     |交费单位               |收据号 |打折比率   "
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         DataMember      =   ""
         Begin VB.TextBox ctxt同一批号 
            Appearance      =   0  'Flat
            Height          =   1650
            Left            =   0
            TabIndex        =   30
            Top             =   255
            Visible         =   0   'False
            Width           =   1710
         End
      End
      Begin VB.Label clab退费记录数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   9900
         TabIndex        =   38
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退费记录数："
         Height          =   180
         Left            =   8760
         TabIndex        =   37
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label clab记录数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   8200
         TabIndex        =   36
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费用记录数："
         Height          =   180
         Left            =   7080
         TabIndex        =   35
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退费信息"
         Height          =   180
         Left            =   5640
         TabIndex        =   34
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Height          =   135
         Left            =   5280
         TabIndex        =   33
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "查询要退费的费用"
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   15
      TabIndex        =   6
      Top             =   720
      Width           =   10995
      Begin VB.CheckBox cchk是否通过单位接口查询 
         Height          =   225
         Left            =   10380
         TabIndex        =   27
         Top             =   855
         Width           =   210
      End
      Begin VB.CheckBox cchk按时间查询 
         Height          =   285
         Left            =   6600
         TabIndex        =   0
         Top             =   810
         Value           =   1  'Checked
         Width           =   225
      End
      Begin 录入控件.ctlInputBox cinb退费输入 
         Height          =   360
         Index           =   1
         Left            =   5520
         TabIndex        =   2
         Top             =   360
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   850
         Text            =   ""
         Label           =   "交费人(&F)"
         名称            =   ""
         长度            =   0
         允许等于最大值  =   0   'False
         允许等于最小值  =   0   'False
         允许多选        =   0   'False
      End
      Begin 录入控件.ctlInputBox cinb退费输入 
         Height          =   360
         Index           =   0
         Left            =   8040
         TabIndex        =   3
         Top             =   360
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   1030
         Text            =   ""
         Label           =   "交费单位(&J)"
         名称            =   ""
         长度            =   0
         允许等于最大值  =   0   'False
         允许等于最小值  =   0   'False
         允许多选        =   0   'False
      End
      Begin 录入控件.ctlInputBox cinb退费输入 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   1030
         Text            =   ""
         Label           =   "收费批号(&N)"
         名称            =   ""
         长度            =   0
         允许等于最大值  =   0   'False
         允许等于最小值  =   0   'False
         允许多选        =   0   'False
      End
      Begin MSComCtl2.DTPicker cdtp日期 
         Height          =   300
         Index           =   0
         Left            =   3675
         TabIndex        =   5
         Top             =   810
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   36951
      End
      Begin MSComCtl2.DTPicker cdtp日期 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   810
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   36951
      End
      Begin 录入控件.ctlInputBox cinb退费输入 
         Height          =   360
         Index           =   4
         Left            =   2880
         TabIndex        =   32
         Top             =   360
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   850
         Text            =   ""
         Label           =   "收据号(&S)"
         Length          =   8
         名称            =   ""
         长度            =   0
         允许等于最大值  =   0   'False
         允许等于最小值  =   0   'False
         允许多选        =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "通过单位档案接口查询定位"
         Height          =   195
         Left            =   8055
         TabIndex        =   26
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "按时间查询"
         Height          =   240
         Left            =   5520
         TabIndex        =   18
         Top             =   855
         Width           =   990
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "日期范围(B)"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   885
         Width           =   990
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "至(E)"
         Height          =   180
         Left            =   2880
         TabIndex        =   11
         Top             =   870
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker cdtp日期 
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   529
      _Version        =   393216
      Format          =   20185089
      CurrentDate     =   36951
   End
End
Attribute VB_Name = "frm退费"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************frm退费***********************************************************************************
'创建时间：                 2001-3-29
'创建人：                   林涛
'修改时间：
'修改人：
'***************************BEGIN*****************************************************************************************
Option Explicit
Public pblnInUse As Boolean
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mstr单位编号 As String

'Private pobj收费管理 As Object
'Private pobj业务设置 As Object
'Private pobj单位定位 As Object  '单位档案接口

Private rs查找记录 As ADODB.Recordset

Private mstrSQL As String  '条件字符串
'定义常量
Private Const 收费批号 = 3
Private Const 交费人 = 1
Private Const 交费单位 = 0
Private Const 退费人 = 2
Private Const 收据号 = 4

Private Const 开始日期 = 1
Private Const 结束日期 = 0
Private Const 退费日期 = 2
  

'功能：选择是否通过单位定位接口进行查询单位名称
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cchk是否通过单位接口查询_Click()
    cinb退费输入(交费单位).Text = ""
End Sub



Private Sub cdtp日期_Change(Index As Integer)
    On Error Resume Next
    cdtp日期(退费日期).Value = CDate(Now)
End Sub

'功能：点击费用信息表中的一行，刷新显示所选择的记录
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cgrd费用信息_Click()
    
    Call sub选择同批费用
    Call 动态调整TextBox
End Sub

 



'功能：在费用信息表中屏蔽部分按键
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cgrd费用信息_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
    End If
End Sub

'功能：交费单位输入框内容变化，其对应的交费单位编号也相应改变
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cinb退费输入_Change(Index As Integer)
    On Error Resume Next
    If Index = 交费单位 Then
        mstr单位编号 = ""
    End If
End Sub

'功能：双击 交费单位输入框，弹出交费单位定位接口界面
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cinb退费输入_DblClick(Index As Integer)
    On Error Resume Next
    Dim lobj单位信息 As Recordset
    
    If Index = 交费单位 Then
         If cchk是否通过单位接口查询 Then
            Set lobj单位信息 = pobj单位定位.func单位简单定位(Screen.Width / 2, Screen.Height / 2)
            If Not (lobj单位信息 Is Nothing) Then
                cinb退费输入(交费单位).Text = lobj单位信息.Fields("单位名称").Value
                mstr单位编号 = lobj单位信息.Fields("申请编号").Value
                Set lobj单位信息 = Nothing
            End If
         End If
    End If
 
End Sub

'功能：输入框获得焦点,弹出单位档案接口界面，以获取单位信息
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-4-12
Private Sub cinb退费输入_GotFocus(Index As Integer)
    On Error Resume Next
    Dim lobj单位信息 As Recordset
    
    If Index = 交费单位 Then
         If cchk是否通过单位接口查询 Then
            Set lobj单位信息 = pobj单位定位.func单位简单定位(Screen.Width / 2, Screen.Height / 2)
            If Not (lobj单位信息 Is Nothing) Then
                cinb退费输入(交费单位).Text = lobj单位信息.Fields("单位名称").Value
                mstr单位编号 = lobj单位信息.Fields("申请编号").Value
                Set lobj单位信息 = Nothing
            End If
         End If
    End If
    
End Sub

'功能：弹出单位档案接口界面，以获取单位信息
'输入：用户按下任何键
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cinb退费输入_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    Dim lobj单位信息 As Recordset

    If Index = 交费单位 Then
         If KeyAscii = 8 Then
             If cchk是否通过单位接口查询 Then
                cinb退费输入(交费单位).Text = ""
                Exit Sub
             End If
         End If
         If cchk是否通过单位接口查询 = 0 Then Exit Sub
         Set lobj单位信息 = pobj单位定位.func单位简单定位(Screen.Width / 2, Screen.Height / 2)
         If Not (lobj单位信息 Is Nothing) Then
             cinb退费输入(交费单位).Text = lobj单位信息.Fields("单位名称").Value
             mstr单位编号 = lobj单位信息.Fields("申请编号").Value
             KeyAscii = 0
             Set lobj单位信息 = Nothing
             
         End If
    Else
         If KeyAscii = 13 Then
                cgrd费用信息.Clear
                cgrd费用信息.Rows = 1
                'cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |打折比率"
                cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |收据号 |打折比率"
                Call func查找记录(func查询条件)
                Call sub填充表格
         ElseIf KeyAscii = 39 Then
              KeyAscii = 0
         End If
    End If
    
End Sub


'功能：在状态栏上显示操作提示
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub ctlb工具栏_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lsngW, lsngH, lsngSepW As Single
    
    On Error Resume Next
    lsngW = ctlb工具栏.ButtonWidth
    lsngH = ctlb工具栏.ButtonHeight
    lsngSepW = ctlb工具栏.Buttons(2).Width
    
    With cstu状态栏
    If X <= lsngW And Y <= lsngH Then
       .Panels(1).Text = " 查找记录"
    Else
       .Panels(1).Text = ""
    End If
    
    If X <= 2 * lsngW + lsngSepW And X > lsngW + lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "清除表格上的数据"
    End If
    
    If X <= 3 * lsngW + 2 * lsngSepW And X > 2 * lsngW + 2 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "将表格上所选的费用进行退费"
    End If
    
    If X <= 4 * lsngW + 3 * lsngSepW And X > 3 * lsngW + 3 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "打印票据"
    End If
    
    If X <= 5 * lsngW + 4 * lsngSepW And X > 4 * lsngW + 4 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "关闭退费窗口"
    End If
    
    End With
End Sub


'功能：定时器，以刷新显示进度条
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29

Private Sub ctmr定时_Timer()
    On Error Resume Next
    If cprg进度.Value < cprg进度.Max Then
       cprg进度.Value = cprg进度.Value + 5
    Else
       ctmr定时.Enabled = False
    End If
   Me.Refresh
End Sub


'功能：增加一些快捷键
'作者：林涛
'创建时间：2001-3-29
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyB
            cdtp日期(开始日期).SetFocus
        Case vbKeyE
            cdtp日期(结束日期).SetFocus
                    
        End Select
    End If
End Sub


'功能：装载窗体，初始化界面，给部分控件属性付值
'作者：林涛
'创建时间：2001-3-29
Private Sub Form_Load()
    If pblnInUse Then Exit Sub
    Dim lcol工具栏按钮 As Collection
    
    On Error GoTo errhandler
    pblnInUse = True                              '指示窗体已启动
    
'    Set pobj收费管理 = CreateObject("收费业务对象.cls收费管理")
'    Set pobj业务设置 = CreateObject("收费业务对象.cls业务设置")
'    Set pobj单位定位 = CreateObject("单位档案业务.ClsUnitInterface")
    
    动态调整TextBox
    
    '初始化工具栏
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.c工具栏 = ctlb工具栏
    Set lcol工具栏按钮 = New Collection
    lcol工具栏按钮.Add "查询(&Q)105"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "清空"
    lcol工具栏按钮.Add "退费(&T)122"
    lcol工具栏按钮.Add "打印"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    mobjGUI.subInitialize lcol工具栏按钮, ""
    Set lcol工具栏按钮 = Nothing
    
    '功能：获取是否有退费的权限。时间：2002/02/20 作者：徐冀川
    If umfunc校验用户权限("收费管理_退费") Then
        ctlb工具栏.Buttons(4).Enabled = True
        ctxt退费人.Text = um用户名                    '获取当前用户名
        ctxt退费日期.Text = Date
    Else
        ctlb工具栏.Buttons(4).Enabled = False
    End If
    '功能：获取是否有打印票据的权限。时间：2002/02/20 作者：徐冀川
    If umfunc校验用户权限("收费管理_票据打印") Then
        ctlb工具栏.Buttons(5).Enabled = True
    Else
        ctlb工具栏.Buttons(5).Enabled = False
    End If
    
    cdtp日期(开始日期).Value = Date               '初始化开始日期输入框为本机日期
    cdtp日期(结束日期).Value = Date               '初始化结束日期输入框为本机日期
    'cgrd费用信息.Cols = 7
    cgrd费用信息.Rows = 1
    ' 设退费时标识
    Cchk退费打印标识.Value = 0
    Exit Sub
errhandler:
    Call sfsub错误处理("收费界面对象", "frm退费", "Form_Load", Err.Number, Err.Description, False)
End Sub

'功能：显示查询结果于表格中
'输入：无
'输出：无
'返回：如果查询到结果则显示其内容于表格中，否则表格内容为空白
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Function func查找记录(ByVal strSQL As String) As ADODB.Recordset
    On Error GoTo errhandle
    'Set func查找记录 = pobj收费管理.func查询费用信息(strSQL)
    Set func查找记录 = dafuncGetData("select * from 收费管理_打印费用信息 where " & strSQL)
    Exit Function
errhandle:
    Call sfsub错误处理("收费界面对象", "frm退费", "func查找记录", Err.Number, Err.Description, True)
End Function
'功能：得到用户输入查询条件
'输入：无
'输出：无
'返回：如果输入的查询条件不为空，则返回含“Where”的查询条件字符串，反之返回""
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Function func查询条件() As String
On Error Resume Next
Dim strSQL As String

    strSQL = ""
    '收费批号条件
    If Trim(cinb退费输入(收费批号).Text) <> "" Then
       strSQL = "'" & Trim(cinb退费输入(收费批号).Text) & "',"
    Else
        strSQL = "'',"
    End If
    
    '修改：2002-6-23（杨春）按数据号查询。
    If Trim(cinb退费输入(收据号).Text) <> "" Then
        strSQL = strSQL & "'" & Trim(cinb退费输入(收据号).Text) & "',"
    Else
        strSQL = strSQL & "'',"
    End If
    
    '交费人条件
    If Trim(cinb退费输入(交费人).Text) <> "" Then
        strSQL = strSQL & "'" & Trim(cinb退费输入(交费人).Text) & "',"
    Else
        strSQL = strSQL & "'',"
    End If
    
    '交费单位
    If Trim(cinb退费输入(交费单位).Text) <> "" Then
        strSQL = strSQL & "'" & Trim(cinb退费输入(交费单位).Text) & "',"
    Else
        strSQL = strSQL & "'',"
    End If
    
     '交费日期范围
    If cchk按时间查询.Value = 1 Then
    
        If Trim(cdtp日期(开始日期)) <> "" Then
            strSQL = strSQL & "'" & Trim(cdtp日期(开始日期)) & "',"
        Else
            strSQL = strSQL & "'',"
        End If
        
        If Trim(cdtp日期(结束日期)) <> "" Then
            strSQL = strSQL & "'" & Trim(cdtp日期(结束日期)) & "'"
        Else
            strSQL = strSQL & "''"
        End If
    Else
        strSQL = strSQL & "'',''"
    End If
    
    func查询条件 = strSQL

End Function
'功能：在其它管理中，提供对收费票据的打印功能.
'时间: 2002/02/20
'作者：徐冀川
Private Sub sub打印票据()
On Error GoTo errhandle
    Dim lcol费用信息 As Collection         '将费用信息所有字段信息写入集合中
    Dim lcol费用打印信息集 As Collection   '存放费用信息的集合
    Dim lrec查找记录 As Object             '存放查询出的费用信息
    Dim lstr格式文件名 As String           '记录打印格式的文件名
    Dim lrec格式文件名对象 As Object           '记录打印格式的文件名对象
    Dim lrec费用信息 As Object             '记录详的费用信息
    Dim lrec费用票据信息 As Object         ' 记录与票据有关的信息
    Dim i As Long                         '循环变量
    Dim j As Long                         '循环变量
    Dim k As Long                         '循环变量
    Dim lstr交费人 As String               '记录交费人姓名
    Dim lstr交费单位 As String             ' 记录交费单位姓名
    Dim lsge打折比率 As Single            '记录打折比率
    Dim lobj汇总记录 As Object
    
    '判断是否选中记录
    With cgrd费用信息
    If .Row = 0 Then
        MsgBox "请选择要打印的费用信息。", vbInformation, "打印票据"
        Exit Sub
    End If
    '设置退出按钮不可用
    ctlb工具栏.Buttons(7).Enabled = False
    
    
    '更新查询数据接口
    '时间：2002/08/05
    '作者：徐冀川
    Dim lstr存储过程 As String
    mstrSQL = func查询条件
    lstr存储过程 = "exec 收费管理_返回收费信息 " + mstrSQL
    Set lrec查找记录 = dafuncGetData(lstr存储过程)
    
    'Set lrec查找记录 = func查找记录(mstrSQL)
    
    '判断是否选则同一批号数据
    Dim lstrtemp As String
    lstrtemp = .TextMatrix(.Row, 0)
    For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                 lrec查找记录.MoveFirst
                 lrec查找记录.Move i - 1
                If lstrtemp <> lrec查找记录("收费批号") Then
                    MsgBox "你已选择了不同的收费批号！请重新选择。", vbOKOnly, "打印票据"
                    Call sub选择同批费用
                    Exit Sub
                End If
             End If
    Next i
    End With
    '获取与费用信息相关的票据信息
    Set lrec费用票据信息 = pobj收费管理.funcExecute("select b.票据类型编号 from 收费管理_收费项目字典表 b, 收费管理_费用信息表 c " & _
                                               "Where b.收费项目编号 = c.收费项目编号 and c.收费批号 ='" & _
                                               lrec查找记录("收费批号") & "' group by b.票据类型编号", "cls费用信息")
    '校检与费用信息相关的票据信息
    If (lrec费用票据信息 Is Nothing) Or (lrec费用票据信息.BOF And lrec费用票据信息.EOF) Then
        sffuncMsg "未检索到收费项目的票据类型信息,无法进行打印!", sf警告
        Exit Sub
    Else
        lrec费用票据信息.MoveFirst
    End If
                                                                                                                             
    '按票据类型取出费用信息
    For i = 0 To lrec费用票据信息.RecordCount - 1
        '获取打印费用信息
        Set lrec费用信息 = pobj收费管理.funcExecute("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & lrec查找记录("收费批号") & "'", "cls费用信息")
        '校检费用信息
        If (lrec费用信息 Is Nothing) Or (lrec费用信息.BOF And lrec费用信息.EOF) Then
            sffuncMsg "无可打印信息!", sf警告
            Exit Sub
        End If
        '处理费用信息中交费人和交费单位为空值的情况
        If IIf(IsNull(lrec费用信息("交费单位名称").Value), "", lrec费用信息("交费单位名称")) <> "" Then
            lstr交费单位 = lrec费用信息("交费单位名称").Value
        Else
            lstr交费单位 = ""
        End If
        If IIf(IsNull(lrec费用信息("交费人").Value), "", lrec费用信息("交费人")) <> "" Then
            lstr交费人 = lrec费用信息("交费人").Value
        Else
            lstr交费人 = ""
        End If
        '初始化打折比率值
        lsge打折比率 = 1
        Set lcol费用打印信息集 = New Collection
        
        '修改：2002-9-29（杨春）合并打印。
        Set lobj汇总记录 = pobj收费管理.funcExecute("select 收费项目编号,单价=avg(单价),数量=sum(数量),金额=sum(金额) from 收费管理_打印费用信息 " _
                            & "where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & lrec查找记录("收费批号") _
                            & "' group by 收费批号,收费项目编号", "cls费用信息")
        
        '将费用信息加入到合对象中
        For j = 0 To lobj汇总记录.RecordCount - 1
            '修改：2002-9-29（杨春）获取当前项目的详细信息。
            Set lrec费用信息 = pobj收费管理.funcExecute("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & lrec查找记录("收费批号") & "' and 收费项目编号='" & lobj汇总记录("收费项目编号") & "'", "cls费用信息")
            
            Set lcol费用信息 = New Collection
            For k = 0 To lrec费用信息.Fields.Count - 1
                If lrec费用信息.Fields(k).Name = "交费单位名称" Or lrec费用信息.Fields(k).Name = "交费人" Or lrec费用信息.Fields(k).Name = "打折比率" Then
                    If lrec费用信息.Fields(k).Name = "交费单位名称" Then lcol费用信息.Add lstr交费单位, "交费单位名称"
                    If lrec费用信息.Fields(k).Name = "交费人" Then lcol费用信息.Add lstr交费人, "交费人"
                    If lrec费用信息.Fields(k).Name = "打折比率" Then
                        lsge打折比率 = lrec费用信息(k).Value
                        lcol费用信息.Add lsge打折比率, "打折比率"
                    End If
                ElseIf lrec费用信息.Fields(k).Name <> "单价" And lrec费用信息.Fields(k).Name <> "数量" And lrec费用信息.Fields(k).Name <> "金额" Then
                    '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
                    lcol费用信息.Add lrec费用信息(k).Value, lrec费用信息.Fields(k).Name
                End If
            Next k
            '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
            lcol费用信息.Add Format(lobj汇总记录("单价").Value, "0.00"), "单价"
            lcol费用信息.Add lobj汇总记录("数量").Value, "数量"
            lcol费用信息.Add Format(lobj汇总记录("金额").Value, "0.00"), "金额"
            
            lcol费用信息.Add "年龄值", "年龄"
            lcol费用信息.Add "性别值", "性别"
            lcol费用信息.Add "住院号值", "住院号"
            lcol费用信息.Add "病种值", "病种"
            lcol费用信息.Add "2002", "入院日期"
            lcol费用信息.Add "2002", "出院日期"
            lcol费用信息.Add "入院操作员值", "入院操作员"
            lcol费用信息.Add "经治医生值", "经治医生"
            
            lcol费用打印信息集.Add lcol费用信息
            
            'If Not lrec费用信息.EOF Then lrec费用信息.MoveNext
            If Not lobj汇总记录.EOF Then lobj汇总记录.MoveNext
        Next j
        '获取格式文件名
        Set lrec格式文件名对象 = pobj收费管理.funcExecute("select * from 收费管理_票据设置信息表 where 票据类型编号='" & lrec费用票据信息("票据类型编号") & "' and 对应业务='一般'", "cls费用信息")
        If lrec格式文件名对象 Is Nothing Then
            sffuncMsg "未查找到票据格式文件!", sf警告
        End If
        If lrec格式文件名对象.BOF And lrec格式文件名对象.EOF Then
            sffuncMsg "未查找到票据格式文件!", sf警告
        Else
            lstr格式文件名 = lrec格式文件名对象("票据格式文件名称")
            Call pobj收费管理.sub打印票据(lcol费用打印信息集, App.Path & "\" & lstr格式文件名, , lsge打折比率, lrec格式文件名对象("最大项数").Value)
        End If
        '判断记录集
        If Not lrec费用票据信息.EOF Then lrec费用票据信息.MoveNext
    Next i
    '设置退出按钮可用
    ctlb工具栏.Buttons(7).Enabled = True
Exit Sub
errhandle:
    sfsub错误处理 "收费界面对象", "frm退费", "sub打印票据", Err.Number, Err.Description, True
End Sub
'功能：退费，改变收费状态为“2:已退费”
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub sub退费()
On Error GoTo errhandle
    Dim mcol费用信息 As Collection  '将费用信息所有字段信息写入集合中：
    Dim lcol费用打印信息集 As Collection    '存放费用信息的集合
    Dim lcol费用信息 As Collection         '将费用信息所有字段信息写入集合中
    Dim lstr格式文件名 As String           '记录打印格式的文件名
    Dim lrec格式文件名对象 As Object           '记录打印格式的文件名对象
    Dim lrec费用信息 As Object             '记录详的费用信息
    Dim lrec费用票据信息 As Object         ' 记录与票据有关的信息
    Dim i As Long                         '循环变量
    Dim j As Long                         '循环变量
    Dim k As Long                         '循环变量
    Dim lstr交费人 As String               '记录交费人姓名
    Dim lstr交费单位 As String             ' 记录交费单位姓名
    Dim lsge打折比率 As Single            '记录打折比率
    Dim lsge金额 As Single                '记录打金额
    Dim lbln退费成功标记 As Boolean        ' 记录退费成功状态
    
    Dim lobj汇总记录 As Object
    
    lbln退费成功标记 = False
    
    '****************退费信息处理*****************
    With cgrd费用信息
    If .Row = 0 Then
        MsgBox "请选择要退费的费用信息。", vbInformation, "退费"
        Exit Sub
    End If
    
    '更新查询数据接口
    '时间：2002/08/05
    '作者：徐冀川
    Dim lstr存储过程 As String      '定义变量记录执行存储过程语句
    mstrSQL = func查询条件
    lstr存储过程 = "exec 收费管理_返回收费信息 " + mstrSQL
    Set rs查找记录 = dafuncGetData(lstr存储过程)
    
    'Set rs查找记录 = func查找记录(mstrSQL)
    
    '*****以下为林涛修改-0902*****************
    Dim lstrtemp As String
    lstrtemp = .TextMatrix(.Row, 0)
    For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                 rs查找记录.MoveFirst
                 rs查找记录.Move i - 1
                If lstrtemp <> rs查找记录("收费批号") Then
                    MsgBox "你已选择了不同的收费批号！请重新选择。", vbOKOnly, "退费"
                    Call sub选择同批费用
                    Exit Sub
                End If
             End If
    Next i
    '*****以上为林涛修改-0902*****************
    
    '功能：增加对费用信息的验证
    '时间：2002/08/05
    '作者：徐冀川
    
    If rs查找记录.RecordCount > 0 Then
        For i = 1 To .Rows - 1
            If fun校验费信息(.TextMatrix(.RowSel, 0)) = False Then
                sffuncMsg "此费用信息已退费！"
                Exit Sub
            End If
            
        Next
    End If
    
    If MsgBox("交费单位     ：" & .TextMatrix(.Row, 6) & Chr(13) & Chr(10) & "交费人       ：" & .TextMatrix(.Row, 5) & Chr(13) & Chr(10) & "收费批号     ：" & .TextMatrix(.Row, 0) & Chr(13) & Chr(10) & "总金额       ：" & ctxt总金额 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "   您真要退费吗？", vbYesNo, "退费") = vbNo Then Exit Sub
    dasubBeginTran      '开始事务
    For i = 1 To .Rows - 1
        If .IsSelected(i) = True Then
             rs查找记录.MoveFirst
             rs查找记录.Move i - 1
             Set mcol费用信息 = New Collection
             '下面各个字段值原封不动，键名与数据库：收费管理_费用信息表字段相对应
             mcol费用信息.Add rs查找记录("收费批号"), "收费批号"
             mcol费用信息.Add rs查找记录("收费编号"), "收费编号"
             mcol费用信息.Add rs查找记录("收费项目编号"), "收费项目编号"
             mcol费用信息.Add rs查找记录("数量"), "数量"
             mcol费用信息.Add rs查找记录("单价"), "单价"
             mcol费用信息.Add rs查找记录("金额"), "金额"
             mcol费用信息.Add rs查找记录("交费人"), "交费人"
             mcol费用信息.Add IIf(IsNull(rs查找记录("交费单位编号")), "", rs查找记录("交费单位编号")), "交费单位编号"
             mcol费用信息.Add rs查找记录("交费日期"), "交费日期"
             mcol费用信息.Add rs查找记录("收费人"), "收费人"
             mcol费用信息.Add rs查找记录("主管科室经手人"), "主管科室经手人"
             mcol费用信息.Add rs查找记录("主管科室编号"), "主管科室编号"
             mcol费用信息.Add rs查找记录("打折比率"), "打折比率"
             mcol费用信息.Add rs查找记录("交费方式"), "交费方式"
             mcol费用信息.Add IIf(IsNull(rs查找记录("交费单位名称")), "", rs查找记录("交费单位名称")), "交费单位名称"
             '下面三个字段是要修改的内容
             mcol费用信息.Add "2", "收费状态"
             mcol费用信息.Add um用户编号, "退费人"
             mcol费用信息.Add func获取服务器日期, "退费日期"
             mcol费用信息.Add IIf(IsNull(rs查找记录("收据号")), "", rs查找记录("收据号")), "收据号"
             '通知业务层修改费用信息
             pobj收费管理.func修改费用信息 mcol费用信息
         End If
    Next i
    '若退费成功，则清除表格已作退费的数据
     dasubCommitTran
     i = 1
     Do While i <= .Rows - 1
        DoEvents
        If .IsSelected(i) Then
RemoveLine:            .RemoveItem i
            If i <= .Rows - 1 Then
                If .IsSelected(i) Then GoTo RemoveLine
            End If
        End If
        i = i + 1
     Loop
    MsgBox "退费已成功。", vbInformation, "退费"
    lbln退费成功标记 = True
    Call sub选择同批费用
    Call 动态调整TextBox
    Set mcol费用信息 = Nothing
    End With
   
    '功能：增加对退费信息的打印功能。
    '时间：2002/02/20
    '作者：徐冀川
    If lbln退费成功标记 = True Then
         If Cchk退费打印标识.Value = 0 Then Exit Sub
         
         '修改：2002-9-29（杨春）退费票据打印独立出来，以便可以补打。
         sub打印退费票据 rs查找记录("收费批号")
    End If
    
    Exit Sub
errhandle:
    dasubRollBack
    Call sfsub错误处理("收费界面对象", "frm退费", "sub退费", Err.Number, Err.Description, True)
    Exit Sub
    Resume
End Sub

'功能：增加对退费信息的打印功能。
'时间：2002/02/20
'作者：徐冀川
'修改：2002-9-29（杨春）把退费票据打印独立出来。
Private Sub sub打印退费票据(ByVal para收费批号 As String)
    Dim lcol费用打印信息集 As Collection    '存放费用信息的集合
    Dim lcol费用信息 As Collection         '将费用信息所有字段信息写入集合中
    Dim lstr格式文件名 As String           '记录打印格式的文件名
    Dim lrec格式文件名对象 As Object           '记录打印格式的文件名对象
    Dim lrec费用信息 As Object             '记录详的费用信息
    Dim lrec费用票据信息 As Object         ' 记录与票据有关的信息
    Dim i As Long                         '循环变量
    Dim j As Long                         '循环变量
    Dim k As Long                         '循环变量
    Dim lstr交费人 As String               '记录交费人姓名
    Dim lstr交费单位 As String             ' 记录交费单位姓名
    Dim lsge打折比率 As Single            '记录打折比率
    Dim lsge金额 As Single                '记录打金额
    
    Dim lobj汇总记录 As Object
    
    On Error GoTo errHanler
    
    '设置打印按钮不可用
     ctlb工具栏.Buttons(7).Enabled = False
    '获取与费用信息相关的票据信息
    Set lrec费用票据信息 = pobj收费管理.funcExecute("select b.票据类型编号 from 收费管理_收费项目字典表 b, 收费管理_费用信息表 c " & _
                                               "Where b.收费项目编号 = c.收费项目编号 and c.收费批号 ='" & _
                                               para收费批号 & "' group by b.票据类型编号", "cls费用信息")
    '校检与费用信息相关的票据信息
    If (lrec费用票据信息 Is Nothing) Or (lrec费用票据信息.BOF And lrec费用票据信息.EOF) Then
        sffuncMsg "未检索到收费项目的票据类型信息,无法进行打印！", sf警告
        Exit Sub
    Else
        lrec费用票据信息.MoveFirst
    End If
    
    '按票据类型取出费用信息
    For i = 0 To lrec费用票据信息.RecordCount - 1
        '获取打印费用信息
        Set lrec费用信息 = pobj收费管理.funcExecute("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & para收费批号 & "'", "cls费用信息")
        '校检费用信息
        If (lrec费用信息 Is Nothing) Or (lrec费用信息.BOF And lrec费用信息.EOF) Then
            sffuncMsg "无可打印信息！", sf警告
            Exit Sub
        End If
        '处理费用信息中交费人和交费单位为空值的情况
        If IIf(IsNull(lrec费用信息("交费单位名称").Value), "", lrec费用信息("交费单位名称")) <> "" Then
            lstr交费单位 = lrec费用信息("交费单位名称").Value
        Else
            lstr交费单位 = ""
        End If
        If IIf(IsNull(lrec费用信息("交费人").Value), "", lrec费用信息("交费人")) <> "" Then
            lstr交费人 = lrec费用信息("交费人").Value
        Else
            lstr交费人 = ""
        End If
        '初始化打折比率值
        lsge打折比率 = 1
        Set lcol费用打印信息集 = New Collection
        
        '修改：2002-9-29（杨春）合并打印。
        Set lobj汇总记录 = pobj收费管理.funcExecute("select 收费项目编号,单价=avg(单价),数量=sum(数量),金额=sum(金额) from 收费管理_打印费用信息 " _
                        & "where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & para收费批号 _
                        & "' group by 收费批号,收费项目编号", "cls费用信息")
        
        '将费用信息加入到合对象中
        For j = 0 To lobj汇总记录.RecordCount - 1
            '修改：2002-9-29（杨春）获取当前项目的详细信息。
            Set lrec费用信息 = pobj收费管理.funcExecute("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrec费用票据信息("票据类型编号") & " and 收费批号='" & para收费批号 & "' and 收费项目编号='" & lobj汇总记录("收费项目编号") & "'", "cls费用信息")
            
            Set lcol费用信息 = New Collection
            For k = 0 To lrec费用信息.Fields.Count - 1
                If lrec费用信息.Fields(k).Name = "交费单位名称" Or lrec费用信息.Fields(k).Name = "交费人" Or lrec费用信息.Fields(k).Name = "打折比率" Or lrec费用信息.Fields(k).Name = "金额" Then
                    If lrec费用信息.Fields(k).Name = "交费单位名称" Then lcol费用信息.Add lstr交费单位, "交费单位名称"
                    If lrec费用信息.Fields(k).Name = "交费人" Then lcol费用信息.Add lstr交费人, "交费人"
                    If lrec费用信息.Fields(k).Name = "打折比率" Then
                        lsge打折比率 = lrec费用信息(k).Value
                        lcol费用信息.Add lsge打折比率, "打折比率"
                    End If
'                        If lrec费用信息.Fields(k).Name = "金额" Then
'                            lsge金额 = 0 - lrec费用信息(k).Value
'                            lcol费用信息.Add lsge金额, "金额"
'                        End If
                ElseIf lrec费用信息.Fields(k).Name <> "单价" And lrec费用信息.Fields(k).Name <> "数量" And lrec费用信息.Fields(k).Name <> "金额" Then
                    '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
                    lcol费用信息.Add lrec费用信息(k).Value, lrec费用信息.Fields(k).Name
                End If
            Next k
            
            '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
            lcol费用信息.Add Format(lobj汇总记录("单价").Value, "0.00"), "单价"
            lcol费用信息.Add lobj汇总记录("数量").Value, "数量"
            lcol费用信息.Add Format(0 - lobj汇总记录("金额").Value, "0.00"), "金额"
            
            lcol费用信息.Add "年龄值", "年龄"
            lcol费用信息.Add "性别值", "性别"
            lcol费用信息.Add "住院号值", "住院号"
            lcol费用信息.Add "病种值", "病种"
            lcol费用信息.Add "2002", "入院日期"
            lcol费用信息.Add "2002", "出院日期"
            lcol费用信息.Add "入院操作员值", "入院操作员"
            lcol费用信息.Add "经治医生值", "经治医生"
            
            lcol费用打印信息集.Add lcol费用信息
            'If Not lrec费用信息.EOF Then lrec费用信息.MoveNext
            If Not lobj汇总记录.EOF Then lobj汇总记录.MoveNext
        Next j
        '获取格式文件名
        Set lrec格式文件名对象 = pobj收费管理.funcExecute("select * from 收费管理_票据设置信息表 where 票据类型编号='" & lrec费用票据信息("票据类型编号") & "' and 对应业务='一般'", "cls费用信息")
        If lrec格式文件名对象 Is Nothing Then
            sffuncMsg "未查找到票据格式文件！", sf警告
        End If
        If lrec格式文件名对象.BOF And lrec格式文件名对象.EOF Then
            sffuncMsg "未查找到票据格式文件！", sf警告
        Else
            lstr格式文件名 = lrec格式文件名对象("票据格式文件名称")
            
            '修改：2002-6-25（杨春）增加参数para退费。
            Call pobj收费管理.sub打印票据(lcol费用打印信息集, App.Path & "\" & lstr格式文件名, , lsge打折比率, lrec格式文件名对象("最大项数").Value, True)
        End If
        '判断记录集
        If Not lrec费用票据信息.EOF Then lrec费用票据信息.MoveNext
    Next i

    '设置打印按钮可用
    ctlb工具栏.Buttons(7).Enabled = True

    Exit Sub
errHanler:
    Call sfsub错误处理("收费界面对象", "frm退费", "sub打印退费票据", Err.Number, Err.Description, True)
    ctlb工具栏.Buttons(7).Enabled = True
    Exit Sub
    Resume
End Sub
'功能：获取服务器系统当前时间值
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cstu状态栏.Panels(1).Text = "退费"
End Sub

'功能：关闭窗体，设置pblninuse为False
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
'    Set pobj收费管理 = Nothing
'    Set pobj业务设置 = Nothing
'    Set pobj单位定位 = Nothing
    Set mobjGUI = Nothing
    Set rs查找记录 = Nothing
End Sub
'功能：根据查询的数据填充表格
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub sub填充表格1()
 
Dim i As Integer
    On Error GoTo errhandle
    cprg进度.Visible = True
    cprg进度.Value = 15
    
    
    
    'mstrSQL = func查询条件
    'Set rs查找记录 = func查找记录(mstrSQL)
    
    '更新查询数据接口
    '时间：2002/08/05
    '作者：徐冀川
    Dim lstr存储过程 As String      '定义变量记录执行存储过程语句
    mstrSQL = func查询条件
    lstr存储过程 = "exec 收费管理_返回收费信息 " + mstrSQL
    Set rs查找记录 = dafuncGetData(lstr存储过程)
    
    
    If rs查找记录.RecordCount <= 0 Then
        MsgBox "未找到匹配记录，可能该项目未交费" & Chr(13) & Chr(10) & "，请检查输入是否正确！", vbInformation, "退费"
        ctmr定时.Enabled = False
        ctxt提示.Visible = False
        cprg进度.Value = 0
        cprg进度.Visible = False
        Exit Sub
    End If
    cgrd费用信息.Clear       '清空表格
    cgrd费用信息.Cols = 7    '只显示7列
    'cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |打折比率"
    cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |收据号 |打折比率"
    'rs查找记录.MoveLast      '为了确保得到正确的.RecordCount结果
    cgrd费用信息.Rows = rs查找记录.RecordCount + 1
    rs查找记录.MoveFirst
    i = 1
    Do Until rs查找记录.EOF
        '填充费用信息表格：
        cgrd费用信息.TextMatrix(i, 0) = IIf(IsNull(rs查找记录.Fields("收费批号")), "", rs查找记录.Fields("收费批号"))
        cgrd费用信息.TextMatrix(i, 1) = IIf(IsNull(rs查找记录.Fields("收费编号")), "", rs查找记录.Fields("收费编号"))
        'cgrd费用信息.TextMatrix(i, 2) = rs查找记录.Fields("收费项目编号")
        cgrd费用信息.TextMatrix(i, 3) = IIf(IsNull(rs查找记录.Fields("数量")), "0", rs查找记录.Fields("数量"))
        cgrd费用信息.TextMatrix(i, 4) = IIf(IsNull(rs查找记录.Fields("金额")), "0.00", rs查找记录.Fields("金额"))
        cgrd费用信息.TextMatrix(i, 5) = IIf(IsNull(rs查找记录.Fields("交费人")), "", rs查找记录.Fields("交费人"))
        cgrd费用信息.TextMatrix(i, 7) = IIf(IsNull(rs查找记录.Fields("打折比率")), "1", rs查找记录.Fields("打折比率"))
        '转换单位编号为名称
        Dim lstrNum As String
        Dim lstrUnitName As String
        Dim lrdsTemp As ADODB.Recordset
        lstrNum = IIf(IsNull(rs查找记录.Fields("交费单位编号")), "", rs查找记录.Fields("交费单位编号"))
        lstrUnitName = IIf(IsNull(rs查找记录.Fields("交费单位名称")), "", rs查找记录.Fields("交费单位名称"))
        If lstrUnitName = vbNullString Then
            If lstrNum <> vbNullString Then
                Set lrdsTemp = pobj收费管理.funcExecute("select 单位名称 from 单位档案_单位基本信息表 where upper(申请编号)=upper('" & lstrNum & "')", "cls费用信息")
                If Not (lrdsTemp Is Nothing) Then
                    If lrdsTemp.RecordCount = 1 Then
                        cgrd费用信息.TextMatrix(i, 6) = lrdsTemp("单位名称")
                    Else
                        cgrd费用信息.TextMatrix(i, 6) = lstrNum
                    End If
                Else
                    cgrd费用信息.TextMatrix(i, 6) = lstrNum
                End If
                Set lrdsTemp = Nothing
            Else
                cgrd费用信息.TextMatrix(i, 6) = lstrNum
            End If
        Else
            cgrd费用信息.TextMatrix(i, 6) = lstrUnitName
        End If
        
        '转换收费项目编号为名称
        lstrNum = rs查找记录.Fields("收费项目编号")
        If lstrNum <> vbNullString Then
            Set lrdsTemp = pobj收费管理.func查询收费项目("upper(收费项目编号)=upper('" & lstrNum & "')")
            If Not (lrdsTemp Is Nothing) Then
                If lrdsTemp.RecordCount = 1 Then
                    cgrd费用信息.TextMatrix(i, 2) = lrdsTemp("收费项目名称")
                Else
                    cgrd费用信息.TextMatrix(i, 2) = lstrNum
                End If
            Else
                cgrd费用信息.TextMatrix(i, 2) = lstrNum
            End If
                Set lrdsTemp = Nothing
        Else
            cgrd费用信息.TextMatrix(i, 2) = lstrNum
        End If
        
        rs查找记录.MoveNext
        i = i + 1
    Loop
    Call sub选择同批费用
    Call 动态调整TextBox
    
    ctmr定时.Enabled = False
    cprg进度.Value = cprg进度.Max
    cprg进度.Visible = False
    ctxt提示.Visible = False
    
    Exit Sub
errhandle:
        ctmr定时.Enabled = False
        ctxt提示.Visible = False
        cprg进度.Value = 0
        cprg进度.Visible = False
        Call sfsub错误处理("收费界面对象", "frm退费", "sub退费", Err.Number, Err.Description, True)
End Sub

'功能：修改费用在界面上的显示
'时间：2002/08/01
'作者：徐冀川
Private Sub sub填充表格()
Dim i As Integer
    On Error GoTo errhandle
    Dim lint退费记录数 As Long
    '*********************
    '判断条件：
    If Trim(cinb退费输入(收费批号).Text) = "" And Trim(cinb退费输入(收据号).Text) = "" And Trim(cinb退费输入(交费人).Text) = "" And Trim(cinb退费输入(交费单位).Text) = "" Then
        If cchk按时间查询.Value = 0 Then
            MsgBox "至少输入一个条件或指定时间范围！", vbInformation, "退费"
            Exit Sub
        Else
            If cdtp日期(开始日期).Value > cdtp日期(结束日期).Value Then
                MsgBox "开始日期不能大于结束日期。", vbInformation, "退费"
                Exit Sub
            ElseIf DateDiff("d", cdtp日期(开始日期).Value, cdtp日期(结束日期).Value) > 90 Then
                MsgBox "日期范围不能大于90天。", vbInformation, "退费"
                Exit Sub
            End If
        End If
    Else
        If cdtp日期(开始日期).Value > cdtp日期(结束日期).Value Then
            MsgBox "开始日期不能大于结束日期。", vbInformation, "退费"
            Exit Sub
        End If
    End If
    '*********************
    ctxt提示.Visible = True
    ctmr定时.Enabled = True
    ctxt提示.Text = "请稍后..."
    cprg进度.Visible = True
    cprg进度.Value = 15
    Me.Refresh
    
    '功能：重新从数据库中获取费用信息
    '注意：利用存储过程获取，还有费用次数，退费费用的次数
    '时间：2002/08/02
    '作者：徐冀川
    Dim lstr存储过程 As String
    mstrSQL = func查询条件
    lstr存储过程 = "exec 收费管理_返回收费信息 " + mstrSQL
    Set rs查找记录 = dafuncGetData(lstr存储过程)
    
    Dim lobjRec As Object       '定义临时记录变量
    Dim lInt As Long            '定义循环变量
    
    lstr存储过程 = "exec 收费管理_返回费用次数 " + mstrSQL
    Set lobjRec = dafuncGetData(lstr存储过程)
    
    If lobjRec.RecordCount > 0 Then
        lobjRec.MoveFirst
        For lInt = 0 To lobjRec.RecordCount - 1
            If lobjRec("项目") = "总次数" Then
                clab记录数.Caption = IIf(IsNull(lobjRec("次数")), "0", lobjRec("次数"))
            End If
                        
            If lobjRec("项目") = "退费次数" Then
                clab退费记录数.Caption = IIf(IsNull(lobjRec("次数")), "0", lobjRec("次数"))
            End If
                        
            lobjRec.MoveNext
        Next
    End If
    
    If rs查找记录.RecordCount <= 0 Then
        ctxt提示.Visible = False
        ctmr定时.Enabled = False
        cprg进度.Value = 0
        cprg进度.Visible = False
        MsgBox "未找到匹配记录，可能该项目未交费" & Chr(13) & Chr(10) & "请检查输入是否正确！", vbInformation, "退费"
        
        Exit Sub
    End If
    
   
    cgrd费用信息.Clear       '清空表格
    cgrd费用信息.Cols = 10   '只显示9列
    cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |收据号 |打折比率|交费日期     "
    'rs查找记录.MoveLast      '为了确保得到正确的.RecordCount结果
    cgrd费用信息.Rows = rs查找记录.RecordCount + 1
    rs查找记录.MoveFirst
    i = 1
    Do Until rs查找记录.EOF
        '填充费用信息表格：
        
        If rs查找记录("标识") = 2 Then
            'cvfg接种计划.Cell(flexcpBackColor, Row, 2, Row, cvfg接种计划.Cols - 1) = M_CON_时间到
            cgrd费用信息.Cell(flexcpBackColor, i, 0, i, 9) = &HFFC0C0
        Else
            
        End If
        
        cgrd费用信息.TextMatrix(i, 0) = IIf(IsNull(rs查找记录.Fields("收费批号")), "", rs查找记录.Fields("收费批号"))
        cgrd费用信息.TextMatrix(i, 1) = IIf(IsNull(rs查找记录.Fields("收费编号")), "", rs查找记录.Fields("收费编号"))
        cgrd费用信息.TextMatrix(i, 2) = IIf(IsNull(rs查找记录.Fields("收费项目名称")), "", rs查找记录.Fields("收费项目名称"))
        cgrd费用信息.TextMatrix(i, 3) = IIf(IsNull(rs查找记录.Fields("数量")), "0", rs查找记录.Fields("数量"))
        cgrd费用信息.TextMatrix(i, 4) = IIf(IsNull(rs查找记录.Fields("金额")), "0.00", rs查找记录.Fields("金额"))
        cgrd费用信息.TextMatrix(i, 5) = IIf(IsNull(rs查找记录.Fields("交费人")), "", rs查找记录.Fields("交费人"))
        cgrd费用信息.TextMatrix(i, 6) = IIf(IsNull(rs查找记录.Fields("交费单位名称")), "", rs查找记录.Fields("交费单位名称"))
        cgrd费用信息.TextMatrix(i, 7) = IIf(IsNull(rs查找记录.Fields("收据号")), "", rs查找记录.Fields("收据号"))
        cgrd费用信息.TextMatrix(i, 8) = IIf(IsNull(rs查找记录.Fields("打折比率")), "1", rs查找记录.Fields("打折比率"))
        cgrd费用信息.TextMatrix(i, 9) = IIf(IsNull(rs查找记录.Fields("交费日期")), "", rs查找记录.Fields("交费日期"))
        
        rs查找记录.MoveNext
        i = i + 1
    Loop
    
    Call sub选择同批费用
    Call 动态调整TextBox
    ctmr定时.Enabled = False
    'cstu状态栏.Panels(1).Text = "完成"
    cprg进度.Value = cprg进度.Max
    cprg进度.Visible = False
    ctxt提示.Visible = False
    Exit Sub
errhandle:
        ctmr定时.Enabled = False
        ctxt提示.Visible = False
        cprg进度.Value = 0
        cprg进度.Visible = False
        Call sfsub错误处理("收费界面对象", "frm退费", "sub退费", Err.Number, Err.Description, True)
End Sub

'功能：获取服务器系统当前时间值
'输入：无
'输出：无
'返回：时间值，精确到秒
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Function func获取服务器日期() As Date
Dim lrsTemp As New ADODB.Recordset '临时存放执行存储过程获得的结果RecordSet
    On Error GoTo errhandle
    Set lrsTemp = dafuncGetData("Select getdate() as 日期")
    func获取服务器日期 = Format(lrsTemp("日期"), "yyyy-mm-dd hh:mm:ss")  '取数据
    Set lrsTemp = Nothing
    Exit Function
errhandle:
    Call sfsub错误处理("收费界面对象", "frm退费", "func获取服务器日期", Err.Number, Err.Description, True)
End Function
 


'功能：响应用户的工具栏操作
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29

Private Sub mobjGUI_Operate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandle
    Select Case Operate
        Case "查询"
            ctlb工具栏.Buttons(1).Enabled = False
            cgrd费用信息.Clear
            cgrd费用信息.Rows = 1
            'cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |打折比率"
            cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |收据号 |打折比率|收费日期  "
            sub填充表格
            ctlb工具栏.Buttons(1).Enabled = True
        Case "清空"
            '删除查询条件
            cinb退费输入(3).Text = ""
            'cinb退费输入(4).Text = ""
            cinb退费输入(1).Text = ""
            cinb退费输入(0).Text = ""
            
            cgrd费用信息.Clear
            cgrd费用信息.Rows = 1
            'cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |打折比率"
            cgrd费用信息.FormatString = "收费批号        |收费编号       |收费项目     |数量 |金额   |交费人  |交费单位          |收据号 |打折比率|收费日期  "
        Case "退费"
            Call sub退费
            
            '在退费后，界面上的数据通过查询获得
            '时间：2002/08/05 徐冀川
            mobjGUI_Operate "查询", False
        Case "打印"
            If cgrd费用信息.Row < 1 Then Exit Sub
            '修改：2002-9-29（杨春）若选中退费记录，则打印退费票据。
            If cgrd费用信息.Cell(flexcpBackColor, cgrd费用信息.Row, 0) = Label7.BackColor Then
                '退费。
                Call sub打印退费票据(cgrd费用信息.TextMatrix(cgrd费用信息.Row, 0))
            Else
                Call sub打印票据
            End If
    End Select
    Exit Sub
errhandle:
    sffuncMsg Operate & "不成功。" & Err.Description, sf警告
End Sub

'功能：可以多次打印退费信息
'时间：2002/08/02
'作者：徐冀川
Private Sub sub选择同批费用()
Dim i As Integer
Dim lcur金额 As Currency
Dim lbln As Boolean
Dim lInt As Long
    ctxt总金额.Text = "0.00 元"
    ctxt退费批号.Text = ""
    
    lInt = cgrd费用信息.RowSel
    If lInt > 0 Then
         If cgrd费用信息.TextMatrix(lInt, 4) >= 0 Then
            lbln = True
         Else
            lbln = False
         End If
    End If
    With cgrd费用信息
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
        
            If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) Then
                .IsSelected(i) = True
            End If
        Next i
        
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                .TopRow = i
                Exit For
            End If
        Next i
                
        lcur金额 = 0
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
               lcur金额 = lcur金额 + CCur(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 8))
            End If
        Next i
        ctxt总金额.Text = CStr(lcur金额) & " 元"
        ctxt退费批号.Text = .TextMatrix(.Row, 0)
    Else
        ctxt总金额.Text = "0.00 元"
        ctxt退费批号.Text = ""
    End If
    End With
    
    
    '功能:界面元素设置,如果费用没有值,或是没有权限,退费为不可用
    '作者:徐冀川
    '时间:2002/08/14
    '退费权限的控制
    If umfunc校验用户权限("收费管理_退费") Then
        If ctxt总金额.Text = "0 元" Then
            ctlb工具栏.Buttons(4).Enabled = False
        Else
            ctlb工具栏.Buttons(4).Enabled = True
        End If
    Else
        ctlb工具栏.Buttons(4).Enabled = False
    End If
    
    '打印权限的控制
    If umfunc校验用户权限("收费管理_票据打印") Then
        ctlb工具栏.Buttons(5).Enabled = True
    Else
        ctlb工具栏.Buttons(5).Enabled = False
    End If
    
    With cgrd费用信息
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
            If lbln = True Then
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, 4) >= 0 Then
                    .IsSelected(i) = True
                End If
            Else
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, 4) < 0 Then
                    .IsSelected(i) = True
                End If
            End If
        Next i
        
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                .TopRow = i
                Exit For
            End If
        Next i
                
        lcur金额 = 0
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
               lcur金额 = lcur金额 + CCur(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 8))
            End If
        Next i
        ctxt总金额.Text = CStr(lcur金额) & " 元"
        ctxt退费批号.Text = .TextMatrix(.Row, 0)
    Else
        ctxt总金额.Text = "0.00 元"
        ctxt退费批号.Text = ""
    End If
    End With
End Sub

Private Sub 动态调整TextBox()
   
End Sub


'功能：校验费用信息状态
'说明：在收费前，在数据库中校验费用信息的状态，只有是
'      没有退费的信息才能退费。
'返回值：fun校验费信息为true,表示可有退费,否则不可能退费。
'作者：徐冀川
'时间：2002/08/05

Private Function fun校验费信息(ByVal para收费批号 As String) As Boolean
On Error GoTo errhandler
    Dim lstrSql As String           '定义变量记录sql语句
    Dim lobjRec As Object           '定义对象记录临时结果集
    
    '初始化函数返回值
    fun校验费信息 = False
    
    If para收费批号 = "" Then
        Exit Function
    Else
        lstrSql = "select distinct(收费批号),收费状态 from 收费管理_费用信息表  where 收费批号='" & para收费批号 & "'"
        Set lobjRec = dafuncGetData(lstrSql)
        
        If lobjRec.RecordCount > 0 Then
            If lobjRec("收费状态") = 1 Then
                fun校验费信息 = True
            Else
                fun校验费信息 = False
            End If
        End If
    End If
Exit Function
errhandler:
    fun校验费信息 = False
End Function
