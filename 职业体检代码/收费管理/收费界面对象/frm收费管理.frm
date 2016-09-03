VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm收费管理 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "收费管理"
   ClientHeight    =   7620
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm收费管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Cchk退费打印标识 
      Caption         =   "退费时打印票据"
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   7080
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   6960
      Width           =   10995
      Begin VB.TextBox ctxt合计 
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
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   120
         Width           =   2415
      End
      Begin VB.Timer ctmr定时 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   5640
         Top             =   120
      End
      Begin VB.TextBox ctxt退费批号 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   120
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总金额"
         Height          =   180
         Index           =   1
         Left            =   7680
         TabIndex        =   17
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "收费批号"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         Height          =   180
         Index           =   0
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "费用信息"
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10980
      Begin VB.OptionButton coptType 
         Caption         =   "报废票据"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "未收费记录"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "退费记录"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "收费记录"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
         Height          =   5565
         Left            =   75
         TabIndex        =   1
         Top             =   540
         Width           =   7935
         _cx             =   25376428
         _cy             =   25372248
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
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
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
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
         DataMember      =   ""
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   5565
         Left            =   8040
         TabIndex        =   22
         Top             =   540
         Width           =   2895
         _cx             =   162862066
         _cy             =   162866776
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
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
         FormatString    =   ""
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
      End
      Begin VB.Label clblMark 
         BackColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   5640
         TabIndex        =   20
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已作废"
         Height          =   180
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Width           =   540
      End
      Begin VB.Label clab退费记录数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   9900
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退费次数："
         Height          =   180
         Left            =   8880
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label clab记录数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   8200
         TabIndex        =   11
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费次数："
         Height          =   180
         Left            =   7200
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchk预览 
         Caption         =   "打印前预览"
         Height          =   255
         Left            =   6960
         TabIndex        =   21
         Top             =   120
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComCtl2.DTPicker cdtp日期 
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   529
      _Version        =   393216
      Format          =   21299201
      CurrentDate     =   36951
   End
   Begin VB.Menu cmnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu cmnuItemView 
         Caption         =   "查询(&Q)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "刷新(&R)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "快速查找"
         Index           =   3
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "退出(&X)"
         Index           =   5
      End
   End
   Begin VB.Menu cmnuFee 
      Caption         =   "收费(&F)"
      Begin VB.Menu cmnuItemFee 
         Caption         =   "新增(&N)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "收费(&E)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "修改"
         Index           =   5
      End
   End
   Begin VB.Menu cmnuBackFee 
      Caption         =   "报废"
      Begin VB.Menu cmnuItemBackFee 
         Caption         =   "退费(&R)"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu cmnuItemBackFee 
         Caption         =   "报废(&B)"
         Index           =   2
      End
   End
   Begin VB.Menu cmnuPrint 
      Caption         =   "打印"
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "票据"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "打印银行对帐单"
         Index           =   2
      End
   End
   Begin VB.Menu cmenuLocate 
      Caption         =   "定位"
      Visible         =   0   'False
      Begin VB.Menu cmenuItemLocate 
         Caption         =   "快速定位"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm收费管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

Private mstr单位编号 As String

'查询条件
Private mstr收费批号 As String
Private mstr收据号 As String
Private mstr单位名称 As String
Private mstr交费人 As String
Private mstr开始日期 As String
Private mstr截止日期 As String
Private mstr业务分类 As String

Private mobjQueryResult As Object
Private mcolIndex As Collection

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
Private Sub cgrdMain_Click()
    On Error GoTo errHandler
    sub显示一批费用明细
    
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "cgrdMain_Click", Err.Number, Err.Description, False)
End Sub

 



'功能：在费用信息表中屏蔽部分按键
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cgrdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
    End If
End Sub


Private Sub cgrdMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbRightButton Then
        PopupMenu cmenuLocate, vbPopupMenuRightButton ', X, Y
    End If
End Sub

Private Sub cmenuItemLocate_Click(Index As Integer)
    Dim lstr收据号 As String
    Dim i As Long
    
    lstr收据号 = InputBox("收据号：", "快速定位", "")
    If lstr收据号 <> "" Then
        cgrdMain.Row = 0
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.TextMatrix(i, mcolIndex("收据号")) = lstr收据号 Then
                cgrdMain.Row = i
                
                Exit For
            End If
        Next
        cgrdMain_Click
    End If
End Sub


Private Sub cmnuItemBackFee_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '退费
        '****************退费信息处理*****************
        With cgrdMain
            If .Row = 0 Then
                MsgBox "请选择要退费的费用信息。", vbInformation, "系统提示"
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, 0, .Row, 0) = clblMark.BackColor Then
                MsgBox "该记录已退费！", vbInformation, "系统提示"
                Exit Sub
            End If
            If MsgBox("交费单位：" & .TextMatrix(.Row, mcolIndex("交费单位")) & Chr(13) & Chr(10) & "   交费人：" & .TextMatrix(.Row, mcolIndex("交费人")) & Chr(13) & Chr(10) & "收费编号 ：" & .TextMatrix(.Row, mcolIndex("收费编号")) & Chr(13) & Chr(10) & "  总金额 ：" & ctxt总金额 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "   您确定要退这笔费用吗？", vbYesNo, "系统提示") = vbYes Then
                pobj收费管理.sub退费 .TextMatrix(.Row, mcolIndex("收费编号")), um用户编号, Format(Date, "yyyy-mm-dd")
            
                If Cchk退费打印标识.Value <> 1 Then
                    MsgBox "退费已成功。", vbOKOnly + vbInformation, "系统提示"
                End If
          
                If Cchk退费打印标识.Value = 1 Then
                     pobj收费管理.sub打印退费票据 .TextMatrix(.Row, mcolIndex("收费编号")), IIf(cchk预览.Value = 1, True, False)
                End If
                
                '刷新界面。
                sub查询并显示记录
                
            End If
        End With
    
    Case 2 '报废
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要退费的费用信息。", vbInformation, "系统提示"
            Exit Sub
        End If
        If MsgBox("你确信要作废选项中的费用信息吗？作废操作不能恢复！", vbQuestion + vbYesNo + vbDefaultButton2, "系统提示") = vbYes Then
            pobj收费管理.sub报废费用信息 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号"))
            '刷新界面。
            sub查询并显示记录
        End If
    End Select
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "cmnuItemBackFee_Click", Err.Number, Err.Description, False)
End Sub

Private Sub cmnuItemFee_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '直接收费-新增
        frm直接收费.pstr收费编号 = ""
        frm直接收费.Show 1, Me
        
        sub查询并显示记录
    Case 2 '收费
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要收费的费用记录！", vbInformation, "系统提示"
            Exit Sub
        End If
        frm直接收费.pstr收费编号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号"))
        frm直接收费.Show 1, Me
        
        sub查询并显示记录
    Case 3 '删除
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要删除的费用记录！", vbInformation, "系统提示"
            Exit Sub
        End If
        With cgrdMain
            If MsgBox("交费单位：" & .TextMatrix(.Row, mcolIndex("交费单位")) & Chr(13) & Chr(10) & "   交费人：" & .TextMatrix(.Row, mcolIndex("交费人")) & Chr(13) & Chr(10) & "收费批号 ：" & .TextMatrix(.Row, mcolIndex("收费批号")) & Chr(13) & Chr(10) & "  总金额 ：" & ctxt总金额 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "   您确定要删除这笔费用吗？", vbYesNo, "系统提示") = vbYes Then
                pobj收费管理.sub删除 .TextMatrix(.Row, mcolIndex("收费编号"))
            
                '刷新界面。
                sub查询并显示记录
                
            End If
        End With
    Case 5 '修改交费方式
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要修改的费用记录！", vbInformation, "系统提示"
            Exit Sub
        End If
        frm修改收费.pstr收费批号 = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号"))
        frm修改收费.Show 1, Me
        If frm修改收费.pblnOk Then
            sub查询并显示记录
        End If
    End Select
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "cmnuItemFee_Click", Err.Number, Err.Description, False)
End Sub

Private Sub cmnuItemPrint_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '票据
        '判断是否选中记录
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要打印的费用信息。", vbInformation, "系统提示"
            Exit Sub
        End If
        If func录入票据号 <> "" Then
            If coptType(1).Value Then
                '退费票据。
                pobj收费管理.sub打印退费票据 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号")), IIf(cchk预览.Value = 1, True, False)
            Else
                
                '收费票据。
                pobj收费管理.sub打印票据 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号")), IIf(cchk预览.Value = 1, True, False)
                
                sub查询并显示记录
            End If
        End If
    Case 2 '银行对帐单
        frm打印银行对帐单.Show
    End Select
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "cmnuItemPrint_Click", Err.Number, Err.Description, False)
End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '查询
        '输入查询条件.
        frm查询.pstr收费批号 = mstr收费批号
        frm查询.pstr收据号 = mstr收据号
        frm查询.pstr单位名称 = mstr单位名称
        frm查询.pstr交费人 = mstr交费人
        frm查询.pstr开始日期 = mstr开始日期
        frm查询.pstr截止日期 = mstr截止日期
        frm查询.pstr业务分类 = mstr业务分类
        frm查询.Show 1, Me
        If frm查询.pblnOk Then
            
            mstr收费批号 = frm查询.pstr收费批号
            mstr收据号 = frm查询.pstr收据号
            mstr单位名称 = frm查询.pstr单位名称
            mstr交费人 = frm查询.pstr交费人
            mstr开始日期 = frm查询.pstr开始日期
            mstr截止日期 = frm查询.pstr截止日期
            mstr业务分类 = frm查询.pstr业务分类
            sub查询并显示记录
        End If
    Case 2 '刷新
        sub显示记录
    Case 3 '快速定位
        cmenuItemLocate_Click 1
    Case 5
        Unload Me
    End Select

    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "cmnuItemView_Click", Err.Number, Err.Description, False)
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    
    sub显示记录
    
    '只有未收费记录可以收费、删除
    ctlb工具栏.Buttons(4).Enabled = coptType(2).Value
    cmnuItemFee(2).Enabled = coptType(2).Value
    cmnuItemFee(3).Enabled = coptType(2).Value
    
    
    '只有收费记录可以退费,报废。
    cmnuItemBackFee(1).Enabled = coptType(0).Value
    cmnuItemBackFee(2).Enabled = coptType(0).Value
    
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "coptType_Click()", Err.Number, Err.Description, False)
    
End Sub

Private Sub Form_Load()
    If pblnInUse Then Exit Sub
    Dim lcol工具栏按钮 As Collection
    Dim lLen As Integer
    On Error GoTo errHandler
    pblnInUse = True                              '指示窗体已启动

    '初始化工具栏
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.c工具栏 = ctlb工具栏
    Set lcol工具栏按钮 = New Collection
    lcol工具栏按钮.Add "查询(&Q)105"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "新增(&N)102"
    lcol工具栏按钮.Add "收费(&F)103"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "打印"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "报表(&T)110"
    lcol工具栏按钮.Add "明细(&D)102"
    lcol工具栏按钮.Add "退出"
    mobjGUI.subInitialize lcol工具栏按钮, ""
    If Not umfunc校验用户权限("收费管理_内部收费信息删除") Then
        cmnuItemFee(3).Visible = False
    End If
    If Not umfunc校验用户权限("收费管理_直接收费") Then
        ctlb工具栏.Buttons(3).Visible = False
        ctlb工具栏.Buttons(4).Visible = False
        ctlb工具栏.Buttons(5).Visible = False
        cmnuFee.Visible = False
    End If
    If Not umfunc校验用户权限("收费管理_退费") Then
        cmnuItemBackFee(1).Visible = False
    End If
    If Not umfunc校验用户权限("收费管理_报废") Then
        If cmnuItemBackFee(1).Visible Then
            cmnuItemBackFee(2).Visible = False
        Else
            cmnuBackFee.Visible = False
        End If
    End If
    'If Not umfunc校验用户权限("收费管理_票据打印") Then
        ctlb工具栏.Buttons(6).Visible = False
        ctlb工具栏.Buttons(7).Visible = False
        cmnuPrint.Visible = False
    'End If
    
    '默认显示当天的所有收费记录。
    mstr开始日期 = Format(Date, "yyyy-mm-dd")
    mstr截止日期 = Format(Date, "yyyy-mm-dd")
    sub查询并显示记录
    
    cmnuItemBackFee(1).Enabled = coptType(0).Value
    cmnuItemBackFee(2).Enabled = coptType(0).Value
    coptType_Click 1
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "Form_Load", Err.Number, Err.Description, False)
End Sub


'功能：在其它管理中，提供对收费票据的打印功能.
'时间: 2002/02/20
'作者：徐冀川
Private Sub sub打印票据()
On Error GoTo errHandle
    
    '判断是否选中记录
    If cgrdMain.Row = 0 Then
        MsgBox "请选择要打印的费用信息。", vbInformation, "系统提示"
        Exit Sub
    End If
    
    pobj收费管理.sub打印票据 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号")), IIf(cchk预览.Value = 1, True, False)

Exit Sub
errHandle:
    sfsub错误处理 "收费界面部件", "frm退费", "sub打印票据", Err.Number, Err.Description, True
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Frame2.Width = Me.ScaleWidth - Frame2.Left - 60
    Frame4.Width = Me.ScaleWidth - Frame4.Left - 60
    Frame4.Top = Me.ScaleHeight - Frame4.Height - 120
    
    Frame2.Height = Frame4.Top - Frame2.Top - 60
    cgrdMain.Width = Frame2.Width * 0.7
    cgrdMain.Height = Frame2.Height - cgrdMain.Top - 60
    
    cgrdDetail.Left = cgrdMain.Left + cgrdMain.Width + 60
    cgrdDetail.Width = Frame2.Width - cgrdDetail.Left - 60
    cgrdDetail.Height = cgrdMain.Height
    
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
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub


Private Sub sub查询并显示记录()
    Dim i As Integer
    Dim lint退费记录数 As Long
    Dim lobjRec As Object       '定义临时记录变量
    Dim lInt As Long            '定义循环变量
    
    
    On Error GoTo errHandle
    cgrdDetail.Rows = 1
    
    '查询收费记录.
    Set mobjQueryResult = pobj收费管理.func收费管理界面查询(mstr收费批号, mstr收据号, mstr单位名称, mstr交费人, mstr开始日期, mstr截止日期, mstr业务分类, lobjRec)
    
    If lobjRec.RecordCount > 0 Then
        lobjRec.MoveFirst
        For lInt = 0 To lobjRec.RecordCount - 1
            If lobjRec("项目") = "总次数" Then
                clab记录数.Caption = IIf(IsNull(lobjRec("次数")), "0", lobjRec("次数"))
            End If
                        
'            If lobjRec("项目") = "退费次数" Then
'                clab退费记录数.Caption = IIf(IsNull(lobjRec("次数")), "0", lobjRec("次数"))
'            End If
                        
            lobjRec.MoveNext
        Next
    End If
    
    sub显示记录
    
    Exit Sub
errHandle:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "sub查询并显示记录()", Err.Number, Err.Description, True)
End Sub

Private Sub sub显示记录()
    Dim i As Long
    On Error GoTo errHandle
    
    cgrdDetail.Rows = 1
    
    If coptType(0).Value Then
        mobjQueryResult.Filter = "标识=1"
        ctlb工具栏.Buttons(4).Enabled = False
    ElseIf coptType(1).Value Then
        mobjQueryResult.Filter = "标识=2"
    ElseIf coptType(3).Value Then
        mobjQueryResult.Filter = "标识=3"
    Else
        mobjQueryResult.Filter = "标识=0"
    End If
'    mobjQueryResult.Sort = "收费批号,收费编号"
    mobjQueryResult.Sort = "收费编号"
    Set cgrdMain.DataSource = mobjQueryResult
    
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.Cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    
    cgrdMain.ColHidden(mcolIndex("收费状态")) = True
    cgrdMain.ColHidden(mcolIndex("标识")) = True
    
    
'    Call sub选择同批费用
    If cgrdMain.Row > 0 Then
        sub显示一批费用明细
    End If
    '显示合计。
    Dim dblTotal As Double
    For i = 1 To cgrdMain.Rows - 1
        dblTotal = Format(dblTotal + cgrdMain.ValueMatrix(i, mcolIndex("金额")), "0.00")
        
        '显示退费记录的颜色。
        If cgrdMain.TextMatrix(i, mcolIndex("收费状态")) = 2 Or cgrdMain.TextMatrix(i, mcolIndex("收费状态")) = 3 Then
            cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = clblMark.BackColor
        End If
            
    Next
    ctxt合计 = Format(dblTotal, "0.00")
    cgrdMain.AutoSize 0, cgrdMain.Cols - 1
    Exit Sub
errHandle:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "sub显示记录()", Err.Number, Err.Description, True)
    
End Sub

Private Sub mobjGUI_Operate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandle
    Select Case Operate
        Case "查询"
            cmnuItemView_Click 1
        Case "新增"
            cmnuItemFee_Click 1
        Case "收费"
            cmnuItemFee_Click 2
        
        Case "退费"
            cmnuItemBackFee_Click 1
        Case "报废"
            cmnuItemBackFee_Click 2
        Case "明细"
            If cgrdMain.Row < 1 Then Exit Sub
            frm明细.pNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号"))
            frm明细.Show
        Case "打印"
            cmnuItemPrint_Click 1
        Case "报表"
            frm报表.Show
    End Select
    Exit Sub
errHandle:
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
    
    lInt = cgrdMain.RowSel
    If lInt > 0 Then
         If cgrdMain.TextMatrix(lInt, mcolIndex("数量")) >= 0 Then
            lbln = True
         Else
            lbln = False
         End If
    End If
    With cgrdMain
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
               lcur金额 = lcur金额 + CCur(.TextMatrix(i, mcolIndex("数量"))) * Val(.TextMatrix(i, mcolIndex("单价")))
            End If
        Next i
        ctxt总金额.Text = CStr(lcur金额) & " 元"
        ctxt退费批号.Text = .TextMatrix(.Row, 0)
    Else
        ctxt总金额.Text = "0.00 元"
        ctxt退费批号.Text = ""
    End If
    End With
    
    
   
    With cgrdMain
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
            If lbln = True Then
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, mcolIndex("数量")) >= 0 Then
                    .IsSelected(i) = True
                End If
            Else
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, mcolIndex("数量")) < 0 Then
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
               lcur金额 = lcur金额 + CCur(.TextMatrix(i, mcolIndex("数量"))) * Val(.TextMatrix(i, mcolIndex("单价")))
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

Private Sub sub显示一批费用明细()
    Dim lstrNo As String
    Dim lobjRec As Object
    On Error GoTo errHandler
    If cgrdMain.Row < 1 Then Exit Sub
    
    lstrNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号"))
    Set lobjRec = pobj收费管理.func查询费用明细(lstrNo)
    cgrdDetail.FormatString = ""
    Set cgrdDetail.DataSource = lobjRec
    cgrdDetail.AutoResize = True
    cgrdDetail.MergeCol(0) = True
    cgrdDetail.MergeCells = flexMergeRestrictColumns
    ctxt退费批号.Text = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号"))
    ctxt总金额.Text = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("金额"))
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "sub显示一批费用明细()", Err.Number, Err.Description, True)
End Sub

