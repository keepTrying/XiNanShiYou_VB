VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm财务监管 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "财务监管"
   ClientHeight    =   7620
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm财务监管.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdQuery 
      Caption         =   "查  询"
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox clstName 
      Height          =   300
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   720
      Width           =   1815
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
         TabIndex        =   9
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
         TabIndex        =   6
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
         TabIndex        =   4
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
         TabIndex        =   10
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "收费批号"
         Height          =   240
         Left            =   240
         TabIndex        =   7
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
         TabIndex        =   5
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "费用明细信息"
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10980
      Begin VB.OptionButton coptType 
         Caption         =   "作废票据"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "收费记录"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   8
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "作废"
         Height          =   180
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   360
      End
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
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
   End
   Begin MSComCtl2.DTPicker cdtp截止日期 
      Height          =   300
      Left            =   3480
      TabIndex        =   15
      Top             =   720
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker cdtp开始日期 
      Height          =   300
      Left            =   1320
      TabIndex        =   16
      Top             =   720
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   36951
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收费员："
      Height          =   180
      Left            =   5520
      TabIndex        =   20
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期范围"
      Height          =   180
      Left            =   360
      TabIndex        =   18
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   3000
      TabIndex        =   17
      Top             =   720
      Width           =   180
   End
   Begin VB.Menu cmnuView 
      Caption         =   "系统(&S)"
      Begin VB.Menu cmnuItemView 
         Caption         =   "退出(&X)"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frm财务监管"
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
Private mstr收据号 As String

Private mobjQueryResult As Object
Private mcolIndex As Collection

Private mstrSQL As String  '条件字符串
  

Private Sub ccmdQuery_Click()
    sub查询并显示记录
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
    If coptType(1).Value Then
        ctlb工具栏.Buttons(5).Visible = True
    Else
        ctlb工具栏.Buttons(5).Visible = False
    End If
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
End Sub


Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '查询
        sub查询并显示记录
    Case 2 '刷新
        sub显示记录
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
    lcol工具栏按钮.Add "号段(&Q)105"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "报表(&T)110"
    lcol工具栏按钮.Add "交款设置(&J)111"
    lcol工具栏按钮.Add "取消作废(&H)109"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    Dim lobjRec As Object, i As Integer
    Dim lobjRec1 As Object
    
    clstName.Clear
    Set lobjRec = dafuncGetData("select 编号,姓名 from 系统管理_员工基本信息视图 order by 编号")
    For i = 1 To lobjRec.RecordCount
        Set lobjRec1 = dafuncGetData("select * from 系统管理_用户操作权限表 where 用户编号='" & lobjRec(0) & "' and 权限名='收费管理_直接收费'")
        If lobjRec1.RecordCount > 0 Then
            clstName.AddItem lobjRec(0) & " " & lobjRec(1)
        End If
        lobjRec.MoveNext
    Next
    If clstName.ListCount > 0 Then
        clstName.ListIndex = 0
    Else
        MsgBox "当前没有设置具有收费权限的人员！", vbInformation, "系统提示"
    End If
    
    '默认显示当天的所有收费记录。
    cdtp开始日期.Value = Format(Date, "yyyy-mm-dd")
    cdtp截止日期.Value = Format(Date, "yyyy-mm-dd")
    
    sub查询并显示记录
    
    coptType_Click 0
    Exit Sub
errHandler:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "Form_Load", Err.Number, Err.Description, False)
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
    Dim lInt As Long            '定义循环变量
    
    
    On Error GoTo errhandle
    cgrdDetail.Rows = 1
    
    '查询收费记录.
    Set mobjQueryResult = pobj收费管理.func财务监管界面查询(mstr收据号, Left(clstName.Text, InStr(clstName.Text, " ") - 1), cdtp开始日期.Value, cdtp截止日期.Value)
    
    sub显示记录
    
    Exit Sub
errhandle:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "sub查询并显示记录()", Err.Number, Err.Description, True)
End Sub

Private Sub sub显示记录()
    Dim i As Long
    On Error GoTo errhandle
    
    cgrdDetail.Rows = 1
    
    If coptType(0).Value Then
        mobjQueryResult.Filter = "标识=1"
    ElseIf coptType(1).Value Then
        mobjQueryResult.Filter = "标识=3"
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
        
        '显示作废记录的颜色。
        If cgrdMain.TextMatrix(i, mcolIndex("收费状态")) = 3 Then
            cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = clblMark.BackColor
        End If
            
    Next
    ctxt合计 = Format(dblTotal, "0.00")
    cgrdMain.AutoSize 0, cgrdMain.Cols - 1
    Exit Sub
errhandle:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "sub显示记录()", Err.Number, Err.Description, True)
    
End Sub


Private Sub mobjGUI_Operate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandle
    Select Case Operate
        Case "号段"
            frm号段设置.clstName.ListIndex = clstName.ListIndex
            frm号段设置.ccmdSave.Enabled = False
            frm号段设置.Show 1
        Case "报表"
            frm报表.clblName.Visible = True
            frm报表.clstName.Visible = True
            frm报表.Show 1
        Case "交款设置"
            frm交款设置.Show 1
        Case "取消作废"
            If cgrdMain.Row = 0 Then
                MsgBox "请选择要取消的票据信息。", vbInformation, "系统提示"
                Exit Sub
            End If
            If MsgBox("你确信要将选中的发票恢复为正常吗？", vbQuestion + vbYesNo + vbDefaultButton2, "系统提示") = vbYes Then
                pobj收费管理.sub取消报废费用信息 cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("收费编号"))
                '刷新界面。
                sub查询并显示记录
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

