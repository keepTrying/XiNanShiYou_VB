VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm数据导入导出 
   BorderStyle     =   0  'None
   Caption         =   "收费数据导入导出"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame5 
      Caption         =   "数据选择"
      Height          =   645
      Left            =   4785
      TabIndex        =   31
      Top             =   6225
      Width           =   2805
      Begin VB.OptionButton copt基础数据 
         Caption         =   "基础数据"
         Height          =   240
         Left            =   1740
         TabIndex        =   33
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton copt业务数据 
         Caption         =   "业务数据"
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   255
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   1440
         X2              =   1440
         Y1              =   120
         Y2              =   615
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   1425
         X2              =   1425
         Y1              =   105
         Y2              =   750
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "基础数据"
      Height          =   765
      Left            =   4755
      TabIndex        =   29
      Top             =   4710
      Width           =   5685
      Begin VB.CheckBox cchk系统信息 
         Caption         =   "系统信息"
         Enabled         =   0   'False
         Height          =   270
         Left            =   180
         TabIndex        =   30
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "注：导入导出该信息可能需要较长的时间"
         Height          =   240
         Left            =   2235
         TabIndex        =   34
         Top             =   375
         Width           =   3360
      End
   End
   Begin VB.Frame cfra操作进度 
      Caption         =   "操作进度:"
      Height          =   555
      Left            =   4755
      TabIndex        =   27
      Top             =   5655
      Width           =   5670
      Begin MSComctlLib.ProgressBar cprg操作进度 
         Height          =   285
         Left            =   75
         TabIndex        =   28
         Top             =   195
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VSFlex6Ctl.vsFlexGrid Cgrd记录显示 
      Height          =   5985
      Left            =   75
      TabIndex        =   22
      Top             =   945
      Width           =   4575
      _cx             =   4202374
      _cy             =   4204861
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorAlternate=   12640511
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
      Rows            =   27
      Cols            =   5
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
   Begin VB.Frame Frame1 
      Caption         =   "操作选择"
      Height          =   630
      Left            =   7710
      TabIndex        =   13
      Top             =   6225
      Width           =   2715
      Begin VB.OptionButton copt操作选择 
         Caption         =   "数据导出"
         Height          =   180
         Index           =   1
         Left            =   1500
         TabIndex        =   15
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton copt操作选择 
         Caption         =   "数据导入"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   285
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   1335
         X2              =   1335
         Y1              =   105
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1320
         X2              =   1320
         Y1              =   105
         Y2              =   600
      End
   End
   Begin VB.Frame cfra业务数据 
      Caption         =   "业务数据"
      Height          =   3750
      Left            =   4725
      TabIndex        =   1
      Top             =   855
      Width           =   5700
      Begin VB.CheckBox cchk项目选择 
         Caption         =   "所有信息"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   16
         Top             =   3345
         Width           =   1110
      End
      Begin VB.CheckBox cchk项目选择 
         Caption         =   "费用信息"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   1110
      End
      Begin VB.CheckBox cchk项目选择 
         Caption         =   "收费项目"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1845
         Width           =   1110
      End
      Begin VB.CheckBox cchk项目选择 
         Caption         =   "收费标准"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   2145
         Width           =   1110
      End
      Begin VB.CheckBox cchk项目选择 
         Caption         =   "票据格式"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   2445
         Width           =   1110
      End
      Begin VB.CheckBox cchk项目选择 
         Caption         =   "打折情况"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   2745
         Width           =   1110
      End
      Begin VB.CheckBox cchk项目选择 
         Caption         =   "系统设置"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   7
         Top             =   3045
         Width           =   1110
      End
      Begin VB.Frame Frame3 
         Caption         =   "条件"
         Height          =   1470
         Index           =   1
         Left            =   1245
         TabIndex        =   2
         Top             =   240
         Width           =   4320
         Begin VB.CheckBox cchk按时间查询 
            Height          =   210
            Left            =   1170
            TabIndex        =   26
            Top             =   270
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.TextBox ctxt时间 
            Height          =   300
            Index           =   1
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   19
            Text            =   "00:00:00"
            Top             =   1005
            Width           =   1455
         End
         Begin VB.TextBox ctxt时间 
            Height          =   300
            Index           =   0
            Left            =   945
            MaxLength       =   8
            TabIndex        =   18
            Text            =   "00:00:00"
            Top             =   1005
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker cdtp日期 
            Height          =   300
            Index           =   0
            Left            =   945
            TabIndex        =   3
            Top             =   570
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   36951
         End
         Begin MSComCtl2.DTPicker cdtp日期 
            Height          =   300
            Index           =   1
            Left            =   2715
            TabIndex        =   4
            Top             =   570
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   36951
         End
         Begin VB.Label Label4 
            Caption         =   "按时间查询"
            Height          =   225
            Left            =   135
            TabIndex        =   25
            Top             =   285
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "时间范围"
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   1065
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   2460
            TabIndex        =   20
            Top             =   1065
            Width           =   180
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   2460
            TabIndex        =   6
            Top             =   630
            Width           =   180
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "日期范围"
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   630
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar cstu状态栏 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   7410
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "数据导入导出"
            TextSave        =   "数据导入导出"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   3945
      Top             =   7335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin VB.Label fdfsfd 
      Caption         =   "记录数："
      Height          =   300
      Left            =   90
      TabIndex        =   24
      Top             =   6990
      Width           =   750
   End
   Begin VB.Label clbl记录数 
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   930
      TabIndex        =   23
      Top             =   7005
      Width           =   1485
   End
End
Attribute VB_Name = "frm数据导入导出"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************frm数据导入导出******************************************************************************
'创建时间：                 2001-3-30
'创建人：                   林涛
'修改时间：
'修改人：
'***************************BEGIN*****************************************************************************************
Option Explicit
Public pblnInUse As Boolean
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
'定义常量
Private Const 数据导入 = 0
Private Const 数据导出 = 1
Private Const 费用信息 = 1
Private Const 收费项目 = 2
Private Const 收费标准 = 3
Private Const 票据格式 = 4
Private Const 打折情况 = 5
Private Const 系统设置 = 6
Private Const 所有信息 = 7
Private Const 开始 = 0
Private Const 结束 = 1

Private mstr条件 As String   '查询条件字符串

'导入或导出的记录总数或各项目总数
Private mint流动的总记录数 As Integer
Private mint流动的费用信息记录数 As Integer
Private mint流动的收费项目记录数 As Integer
Private mint流动的收费标准记录数 As Integer
Private mint流动的票据格式记录数 As Integer
Private mint流动的打折情况记录数 As Integer
Private mint流动的系统设置记录数 As Integer

Private pobj收费管理 As Object
Private pobj业务设置 As Object
Private pobj单位定位 As Object  '单位档案接口
Private Const mstrMDBFile = "\收费管理2001.mdb"

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


'功能： 选择是否按时间范围来查询
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-4-13
Private Sub cchk按时间查询_Click()
Dim i As Integer
    For i = 1 To 7
        cchk项目选择(i).Value = 0
    Next i
    Cgrd记录显示.Clear
    Cgrd记录显示.Rows = 27
    Cgrd记录显示.Cols = 5
    clbl记录数.Caption = ""
End Sub

'功能： 选择要导入或导出的项目，刷新表格显示的数据
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub cchk项目选择_Click(Index As Integer)
Dim i As Integer
    If Not func验证时间(ctxt时间(开始)) Or Not func验证时间(ctxt时间(结束)) Then '用户输入的时间不正确，则退出本过程
        Exit Sub
    End If

    '条件字符串
    If cchk按时间查询 Then
        If copt操作选择(数据导出) Then
            mstr条件 = "交费日期 between '" & cdtp日期(开始) & " " & ctxt时间(开始) & "' and '" & cdtp日期(结束) & " " & ctxt时间(结束) & "'" & " and 收费状态='1'"
        Else
            mstr条件 = "交费日期 between #" & cdtp日期(开始) & " " & ctxt时间(开始) & "# and #" & cdtp日期(结束) & " " & ctxt时间(结束) & "#" & " and 收费状态='1'"
        End If
    Else
            mstr条件 = "收费状态='1'"
    End If
    If Index = 所有信息 And cchk项目选择(Index) Then   '选择了“所有信息”项
        For i = 费用信息 To 系统设置
            cchk项目选择(i).Value = 1
        Next i
        Call sub填充表格("费用信息", copt操作选择(数据导出), mstr条件) '如果选择了所有项目，则在表格中仅显示要导入或导出的费用信息
        Exit Sub
    End If
    
    If cchk项目选择(Index) = 0 Then       '如果用户原来已经选择了该项目，现在要取消该项，则刷新表格
        cchk项目选择(所有信息) = 0
        Cgrd记录显示.Clear
        Cgrd记录显示.Rows = 27
        Cgrd记录显示.Cols = 5
        clbl记录数.Caption = ""
        For i = 费用信息 To 系统设置
            If cchk项目选择(i) Then
                Call sub填充表格(cchk项目选择(i).Caption, copt操作选择(数据导出), mstr条件)
                Exit Sub
            End If
        Next i
    Else
        Call sub填充表格(cchk项目选择(Index).Caption, copt操作选择(数据导出), mstr条件)
    End If
End Sub




'功能： 查询日期条件改变，屏幕上的数据先清空
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-4-13

Private Sub cdtp日期_Change(Index As Integer)
Dim i As Integer
    
    If cchk按时间查询 = 0 Then Exit Sub
    Cgrd记录显示.Clear
    Cgrd记录显示.Rows = 27
    Cgrd记录显示.Cols = 5
    clbl记录数.Caption = ""
    For i = 1 To 7
        cchk项目选择(i).Value = 0
    Next i
End Sub

'功能： 点“操作选择”按钮
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub copt操作选择_Click(Index As Integer)
Dim i As Integer
    If Index = 数据导入 Then
          If ctlb工具栏.Buttons.Item(3).Caption = "导出(&E)" Then
             For i = 1 To 7
                 cchk项目选择(i).Value = 0
             Next i
             Cgrd记录显示.Clear
             Cgrd记录显示.Rows = 27
             Cgrd记录显示.Cols = 5
             clbl记录数.Caption = ""
          End If
          ctlb工具栏.Buttons.Item(3).Caption = "导入(&I)"
    Else
          If ctlb工具栏.Buttons.Item(3).Caption = "导入(&I)" Then
             For i = 1 To 7
                 cchk项目选择(i).Value = 0
             Next i
             Cgrd记录显示.Clear
             Cgrd记录显示.Rows = 27
             Cgrd记录显示.Cols = 5
             clbl记录数.Caption = ""
          End If
          ctlb工具栏.Buttons.Item(3).Caption = "导出(&E)"
    End If
End Sub

Private Sub copt基础数据_Click()
Call subEnabled(False)
End Sub

Private Sub copt业务数据_Click()
    Call subEnabled(True)
End Sub

'功能：点工具栏按钮
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub ctlb工具栏_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim i As Integer
On Error GoTo errhandle
    Select Case Button.Caption
           Case "清空(&C)1"
                Cgrd记录显示.Clear
                Cgrd记录显示.Rows = 27
                Cgrd记录显示.Cols = 5
                clbl记录数.Caption = ""
                For i = 1 To 7
                   cchk项目选择(i).Value = 0
                Next i
           Case "导入(&I)"
                If copt业务数据 Then
                    Call subBegin
                Else
                If cchk系统信息 Then
                    If MsgBox("基础数据，费用信息和票据设置将清空，要继续吗？！", vbInformation + vbYesNo, Mid(ctlb工具栏.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                    Me.Enabled = False
                    'If copt操作选择(数据导入) Then
                    
                    cprg操作进度.Max = 1
                    dafuncGetData "exec 系统管理_清空所有数据"
                    dasubBeginTran
                    umsub数据导入 App.Path & mstrMDBFile, False, cprg操作进度
                    dasubCommitTran
                    'Else
                    '    umsub数据导出 App.Path & mstrMDBFile, cprg操作进度
                    'End If
                    
                    cprg操作进度.Value = 0
                    MsgBox "完成！", vbInformation, Left(ctlb工具栏.Buttons.Item(3).Caption, 2)
                    cprg操作进度.Max = 100
                    Me.Enabled = True
                    
                End If
             End If
                
           Case "导出(&E)"
                If copt业务数据 Then
                    Call subBegin
                Else
                If cchk系统信息 Then
                    If MsgBox("你真的要继续吗？！", vbInformation + vbYesNo, Mid(ctlb工具栏.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                    Me.Enabled = False
                    'If copt操作选择(数据导入) Then
                    '    umsub数据导入 App.Path & mstrMDBFile, True, cprg操作进度
                    'Else
                    cprg操作进度.Max = 1
                    umsub数据导出 App.Path & mstrMDBFile, cprg操作进度
                    'End If
                    
                    cprg操作进度.Value = 0
                    MsgBox "完成！", vbInformation, Left(ctlb工具栏.Buttons.Item(3).Caption, 2)
                    cprg操作进度.Max = 100
                    Me.Enabled = True
                End If
             End If
           
    End Select
    Exit Sub
errhandle:
    'MsgBox Err.Number & " " & Err.Description
    MsgBox "操作失败！", vbInformation, Left(ctlb工具栏.Buttons.Item(3).Caption, 2)
    cprg操作进度.Value = 0
    cprg操作进度.Max = 100
    cfra操作进度.Caption = "操作进度:"
    Me.Enabled = True

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
    lsngW = ctlb工具栏.ButtonWidth
    lsngH = ctlb工具栏.ButtonHeight
    lsngSepW = ctlb工具栏.Buttons(2).Width
    
    With cstu状态栏
    If X <= lsngW And Y <= lsngH Then
       .Panels(1).Text = " 清除表格上的数据"
    Else
       .Panels(1).Text = ""
    End If
    
    If X <= 2 * lsngW + lsngSepW And X > lsngW + lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "导入：将本地数据库数据添加到服务器，导出：将服务器中的数据导出到本地库"
       If copt操作选择(数据导入).Value = True Then
       
             ctlb工具栏.Buttons(3).ToolTipText = "导入(&I)"
       Else
             ctlb工具栏.Buttons(3).ToolTipText = "导出(&E)"
       End If
    End If
    
    If X <= 3 * lsngW + 2 * lsngSepW And X > 2 * lsngW + 2 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "关闭导入导出窗口"
    End If
    
    End With
End Sub

'功能： 查询时间条件改变，先清空窗口上的数据
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-4-13
Private Sub ctxt时间_Change(Index As Integer)
Dim i As Integer
    
    If cchk按时间查询 = 0 Then Exit Sub
    Cgrd记录显示.Clear
    Cgrd记录显示.Rows = 27
    Cgrd记录显示.Cols = 5
    clbl记录数.Caption = ""
    For i = 1 To 7
        cchk项目选择(i).Value = 0
    Next i
End Sub




'功能： 在进行时间条件输入时，禁止通过鼠标右键进行删除等操作
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-4-13

Private Sub ctxt时间_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ctxt时间(Index).Locked = True
End Sub

'功能： 在进行时间条件输入时，禁止通过鼠标右键进行删除等操作
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-4-13

Private Sub ctxt时间_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctxt时间(Index).Locked = False
End Sub

'功能： 屏蔽Delete、和Del键
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-4-13
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    If KeyCode = 46 Or KeyCode = 110 Then
       KeyCode = 0
    End If
    'If Ctlb工具栏.Buttons.Item(3).Caption = "导出(&E)" Then
        If Shift = 4 And KeyCode = vbKeyI Then
            'Call subBegin
            If copt业务数据 Then
                Call subBegin
             Else
                If cchk系统信息 Then
                    If copt操作选择(数据导入) Then
                        If MsgBox("基础数据将清空，要继续吗？！", vbInformation + vbYesNo, Mid(ctlb工具栏.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                        Me.Enabled = False
                        cprg操作进度.Max = 1
                        umsub数据导入 App.Path & mstrMDBFile, True, cprg操作进度
                        
                        cprg操作进度.Value = 0
                        MsgBox "完成！", vbInformation, Left(ctlb工具栏.Buttons.Item(3).Caption, 2)
                        cprg操作进度.Max = 100
                        Me.Enabled = True
                    End If
                End If
             End If
        ElseIf Shift = 4 And KeyCode = vbKeyE Then
             If copt业务数据 Then
                Call subBegin
             Else
                If cchk系统信息 Then
                    If copt操作选择(数据导入).Value = False Then
                        If MsgBox("你真的要继续吗？！", vbInformation + vbYesNo, Mid(ctlb工具栏.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                        Me.Enabled = False
                        cprg操作进度.Max = 1
                        umsub数据导出 App.Path & mstrMDBFile, cprg操作进度
                        
                        cprg操作进度.Value = 0
                        MsgBox "完成！", vbInformation, Left(ctlb工具栏.Buttons.Item(3).Caption, 2)
                        cprg操作进度.Max = 100
                        Me.Enabled = True
                    End If
                End If
             End If
            
            
            
        End If
    'End If
    Exit Sub
errhandle:
    MsgBox "操作失败！", vbInformation, Left(ctlb工具栏.Buttons.Item(3).Caption, 2)
    cprg操作进度.Value = 0
    cprg操作进度.Max = 100
    cfra操作进度.Caption = "操作进度:"
    Me.Enabled = True
End Sub

'功能：在状态栏上显示"数据导入导出"
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cstu状态栏.Panels(1).Text = "数据导入导出"
End Sub


'功能：格式化及校正用户输入的时间值
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub ctxt时间_KeyPress(Index As Integer, KeyAscii As Integer)
        With ctxt时间(Index)
    If (KeyAscii >= 48 And KeyAscii <= 58) Then
        
        Select Case .SelStart
               Case 0
                    If Val(Mid(.Text, 2, 1)) <= 4 Then
                        If KeyAscii > 50 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 49 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 0
                        .SelLength = 1
                        .SetFocus
               Case 1
                    If Val(Mid(.Text, 1, 1)) = 2 Then
                        If KeyAscii > 52 Then
                            KeyAscii = 0
                        End If
                    End If
                    .SelStart = 1
                    .SelLength = 1
                    .SetFocus
               Case 2
                    If Val(Mid(.Text, 5, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 3
                        .SelLength = 1
                        .SetFocus
               Case 3
                    If Val(Mid(.Text, 5, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                    End If
                    .SelStart = 3
                    .SelLength = 1
                    .SetFocus
                   
               Case 4
                    If Val(Mid(.Text, 4, 1)) = 6 Then
                        If KeyAscii > 48 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 4
                        .SelLength = 1
                        .SetFocus
             
               Case 5
                    If Val(Mid(.Text, 8, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                     End If
                        .SelStart = 6
                        .SelLength = 1
                        .SetFocus
                     
               Case 6
                    If Val(Mid(.Text, 8, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 6
                        .SelLength = 1
                        .SetFocus
                   
               Case 7
                    If Val(Mid(.Text, 7, 1)) = 6 Then
                        If KeyAscii > 48 Then
                            KeyAscii = 0
                        End If
                    Else
                    End If
                        .SelStart = 7
                        .SelLength = 1
                        .SetFocus
                    
        End Select
        
    ElseIf KeyAscii = 8 And .SelStart > 0 Then
             KeyAscii = 0
            .SelStart = .SelStart - 1
            .SelLength = 0
    Else
        KeyAscii = 0
    End If
       End With
End Sub

'功能：装载窗体，初始化界面
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
'修改：2001-4-26林涛
Private Sub Form_Load()
    'Dim lobj业务配置 As Object
Dim i As Integer
    If pblnInUse Then Exit Sub

        pblnInUse = True
        Set pobj收费管理 = CreateObject("收费业务对象.cls收费管理")
        Set pobj业务设置 = CreateObject("收费业务对象.cls业务设置")
        Set pobj单位定位 = CreateObject("单位档案业务.ClsUnitInterface")
   
       Dim lcol工具栏按钮 As Collection

     '初始化工具栏
       Set mobjGUI = New cls界面通用对象
       Set mobjGUI.Form = Me
       Set mobjGUI.c工具栏 = ctlb工具栏
       Set lcol工具栏按钮 = New Collection
       lcol工具栏按钮.Add "清空"
       lcol工具栏按钮.Add "|"
       lcol工具栏按钮.Add "导入(&I)112"
       lcol工具栏按钮.Add "|"
       lcol工具栏按钮.Add "退出"
       mobjGUI.subInitialize lcol工具栏按钮, ""
       Set lcol工具栏按钮 = Nothing
       
       cdtp日期(开始) = Date
       cdtp日期(结束) = Date
       
      'subCopyFile
      
      Dim llngattri As Long   '文件的属性值
      llngattri = GetFileAttributes(App.Path & "\收费管理2001.mdb") '读文件的属性
      If Dir(App.Path & "\收费管理2001.mdb") <> "收费管理2001.mdb" Then
            MsgBox "未找到要导入的“收费管理2001.mdb”文件，仅作导出数据或请退出重试。", vbInformation, "数据导入导出"
            'For i = 1 To 7
            '    cchk项目选择(i).Enabled = False
            'Next i
            copt操作选择(0).Enabled = False
            copt操作选择(1).Value = True
            ctlb工具栏.Buttons(3).Caption = "导出(&E)"
            'copt操作选择(1).Enabled = False
            'ctlb工具栏.Buttons(3).Enabled = False
      ElseIf llngattri = 33 Or llngattri = 35 Or llngattri = 3 Or llngattri = 1 Then
            MsgBox "“" & App.Path & "\收费管理2001.mdb”文件为只读属性，请退出修改重试。", vbInformation, "数据导入导出"
            'For i = 1 To 7
            '    cchk项目选择(i).Enabled = False
            'Next i
            copt操作选择(0).Enabled = False
            copt操作选择(1).Value = True
            ctlb工具栏.Buttons(3).Caption = "导出(&E)"
            'copt操作选择(1).Enabled = False
            'Ctlb工具栏.Buttons(3).Enabled = False
      Else
          'For i = 1 To 7
          '      cchk项目选择(i).Enabled = True
          'Next i
          copt操作选择(0).Enabled = True
          copt操作选择(1).Enabled = True
          ctlb工具栏.Buttons(3).Enabled = True
     End If
End Sub

'功能：根据查询条件和查询项目，查询数据库或本地数据库
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Function Func采集数据(ByVal Str采集项目 As String, ByVal bln方向标志 As Boolean, Optional ByVal Str条件 As String) As ADODB.Recordset
Dim lobjTemp As Object
    On Error GoTo errhandle
    Select Case bln方向标志
            Case True  '导出
                Select Case Str采集项目
                       Case "费用信息"
                            Set lobjTemp = pobj收费管理.func查询费用信息(Str条件)  '肯定有时间范围的条件
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的费用信息记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的费用信息记录数 = 0
                                End If
                            Else
                                mint流动的费用信息记录数 = 0
                            End If
                            
                       Case "收费项目"
                            Set lobjTemp = pobj收费管理.func查询收费项目("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的收费项目记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的收费项目记录数 = 0
                                End If
                            Else
                                mint流动的收费项目记录数 = 0
                            End If
                
                            
                       Case "收费标准"
                            Set lobjTemp = pobj业务设置.func查询收费标准("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的收费标准记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的收费标准记录数 = 0
                                End If
                            Else
                                mint流动的收费标准记录数 = 0
                            End If
                
                            
                       Case "票据格式"
                           Set lobjTemp = pobj业务设置.func查询票据信息("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的票据格式记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的票据格式记录数 = 0
                                End If
                            Else
                                mint流动的票据格式记录数 = 0
                            End If
                       
                    
                       Case "打折情况"
                            Set lobjTemp = pobj业务设置.func查询打折信息("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的打折情况记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的打折情况记录数 = 0
                                End If
                            Else
                                mint流动的打折情况记录数 = 0
                            End If
                
                       
                       Case "系统设置"
                            Set lobjTemp = pobj业务设置.func查询业务配置信息("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的系统设置记录数 = 1
                                Else
                                    mint流动的系统设置记录数 = 0
                                End If
                            Else
                                mint流动的系统设置记录数 = 0
                            End If
                                        
                End Select
                
            Case False  '导入
                Select Case Str采集项目
                       Case "费用信息"
                            
                            Set lobjTemp = pobj收费管理.func获取外部数据("收费管理_费用信息表", Str条件)   '肯定有时间范围的条件
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的费用信息记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的费用信息记录数 = 0
                                End If
                            Else
                                mint流动的费用信息记录数 = 0
                            End If
                            
                       Case "收费项目"
                            Set lobjTemp = pobj收费管理.func获取外部数据("收费管理_收费项目字典表")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的收费项目记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的收费项目记录数 = 0
                                End If
                            Else
                                mint流动的收费项目记录数 = 0
                            End If
                            
                       Case "收费标准"
                            Set lobjTemp = pobj收费管理.func获取外部数据("收费管理_收费标准信息表")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的收费标准记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的收费标准记录数 = 0
                                End If
                            Else
                                mint流动的收费标准记录数 = 0
                            End If
                            
                       Case "票据格式"
                            Set lobjTemp = pobj收费管理.func获取外部数据("收费管理_票据设置信息表")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的票据格式记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的票据格式记录数 = 0
                                End If
                            Else
                                mint流动的票据格式记录数 = 0
                            End If
                                                        
                       Case "打折情况"
                            Set lobjTemp = pobj收费管理.func获取外部数据("收费管理_打折信息表")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的打折情况记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的打折情况记录数 = 0
                                End If
                            Else
                                mint流动的打折情况记录数 = 0
                            End If
                       
                       Case "系统设置"
                            Set lobjTemp = pobj收费管理.func获取外部数据("收费管理_业务配置表")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func采集数据 = lobjTemp
                                If Func采集数据.RecordCount > 0 Then
                                    mint流动的系统设置记录数 = Func采集数据.RecordCount
                                Else
                                    mint流动的系统设置记录数 = 0
                                End If
                            Else
                                mint流动的系统设置记录数 = 0
                            End If
                          
               End Select
                 
    End Select
Exit Function
errhandle:
sfsub错误处理 "收费界面对象", "frm数据导入导出", "Func采集数据", Err.Number, Err.Description, True
End Function

'功能：根据导入或导出的项目，调用收费管理对象相应的导入导出方法
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub Sub数据导入导出(ByVal Str项目 As String, ByVal bln方向标志 As Boolean, Optional ByVal Str条件 As String)
On Error GoTo errhandle
    Select Case bln方向标志
           Case False '导出
                Select Case Str项目
                       Case "费用信息"
                            pobj收费管理.func导出费用信息 Str条件
                         
                       Case "收费项目"
                            pobj收费管理.func导出收费项目

                       Case "收费标准"
                            pobj业务设置.func导出收费标准
                            
                       Case "票据格式"
                            pobj业务设置.func导出票据信息
                            
                       Case "打折情况"
                            pobj业务设置.func导出打折信息
                            
                       Case "系统设置"
                            pobj业务设置.func导出业务配置信息
                            
                End Select
           Case True '导入
                Select Case Str项目
                       Case "费用信息"
                            pobj收费管理.func导入费用信息 Str条件
                         
                       Case "收费项目"
                            pobj收费管理.func导入收费项目
                            
                       Case "收费标准"
                            pobj业务设置.func导入收费标准
                            
                       Case "票据格式"
                            pobj业务设置.func导入票据信息
                            
                       Case "打折情况"
                            pobj业务设置.func导入打折信息
                            
                       Case "系统设置"
                            pobj业务设置.func导入业务配置信息
                            
                End Select
           
    End Select
Exit Sub
errhandle:
    sfsub错误处理 "收费界面对象", "frm数据导入导出", "Sub数据导入导出", Err.Number, Err.Description, True
End Sub

'功能：开始导入导出工作，进度条显示操作进度
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub subBegin()
    'On Error GoTo errhandle
    If Not func验证时间(ctxt时间(开始)) Or Not func验证时间(ctxt时间(结束)) Then
        Exit Sub
    End If
    sub统计记录数
    If mint流动的总记录数 <= 0 Then Exit Sub
    If MsgBox("确定要继续吗？", vbYesNo + vbQuestion, Mid(ctlb工具栏.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
    Me.Enabled = False    '开始导入或导出数据时,禁止窗体接收数据。
    With cprg操作进度
        .Max = mint流动的总记录数
        .Min = 0
        .Value = 0
    On Error Resume Next
    Select Case copt操作选择(数据导入)
           Case False '导出时考贝文件
               subCopyFile
               
               If cchk按时间查询 Then
                    mstr条件 = "交费日期 between '" & cdtp日期(开始) & " " & ctxt时间(开始) & "' and '" & cdtp日期(结束) & " " & ctxt时间(结束) & "'" & " and 收费状态='1'"
               Else
                    mstr条件 = "收费状态='1'"
               End If
               If cchk项目选择(费用信息).Value = 1 Then
                    Sub数据导入导出 "费用信息", False, mstr条件
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的费用信息记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的费用信息记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk项目选择(收费项目).Value = 1 Then

                    Sub数据导入导出 "收费项目", False
                
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的收费项目记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的收费项目记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk项目选择(收费标准).Value = 1 Then
                    Sub数据导入导出 "收费标准", False
                   
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的收费标准记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的收费标准记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk项目选择(票据格式).Value = 1 Then
                    Sub数据导入导出 "票据格式", False
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的票据格式记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的票据格式记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk项目选择(打折情况).Value = 1 Then
                    Sub数据导入导出 "打折情况", False
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的打折情况记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的打折情况记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk项目选择(系统设置).Value = 1 Then
                    Sub数据导入导出 "系统设置", False
                    .Value = .Value + mint流动的系统设置记录数
                   
                    cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                    Me.Refresh
                    'Err.Number = 0
               End If
           Case True
               If cchk按时间查询 Then
                    mstr条件 = "交费日期 between #" & cdtp日期(开始) & " " & ctxt时间(开始) & "# and #" & cdtp日期(结束) & " " & ctxt时间(结束) & "#" & " and 收费状态='1'"
               Else
                    mstr条件 = "收费状态='1'"
               End If
               If cchk项目选择(费用信息).Value = 1 Then
                    Sub数据导入导出 "费用信息", True, mstr条件
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的费用信息记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的费用信息记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk项目选择(收费项目) Then
                    Sub数据导入导出 "收费项目", True
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的收费项目记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的收费项目记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk项目选择(收费标准).Value = 1 Then
                    Sub数据导入导出 "收费标准", True
                   
                   If Err.Number = 0 Then
                        .Value = .Value + mint流动的收费标准记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的收费标准记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk项目选择(票据格式).Value = 1 Then
                    Sub数据导入导出 "票据格式", True
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的票据格式记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的票据格式记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk项目选择(打折情况).Value = 1 Then
                    Sub数据导入导出 "打折情况", True
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint流动的打折情况记录数
                
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                    Else
                        If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - mint流动的打折情况记录数
                        cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk项目选择(系统设置).Value = 1 Then
               Dim lobjTempIN  As Object
               Dim lobjTempOUT As Object
                    Set lobjTempIN = pobj业务设置.func查询业务配置信息("")
                    Set lobjTempOUT = pobj收费管理.func获取外部数据("收费管理_业务配置表")
                    If Not (lobjTempIN Is Nothing) And Not (lobjTempOUT Is Nothing) Then
                        If lobjTempIN.RecordCount > 0 And lobjTempOUT.RecordCount > 0 Then
                            If lobjTempIN("科目级数") <> lobjTempOUT("科目级数") Then
                                If MsgBox("将要导入的科目级数与服务器数据不符，是否继续？", vbInformation + vbYesNo, "导入业务设置信息") = vbNo Then
                                    mint流动的系统设置记录数 = 0
                                    
                                    If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - 1
                                    cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                                    Me.Refresh
                                    GoTo WayOut
                                End If
                            Else
                                Sub数据导入导出 "系统设置", True
                                .Value = .Value + mint流动的系统设置记录数
                                    
                                cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                                Me.Refresh
                            End If
                        
                       End If
                   Else
                       mint流动的系统设置记录数 = 0
                       
                       If mint流动的总记录数 > 0 Then mint流动的总记录数 = mint流动的总记录数 - 1
                       cfra操作进度.Caption = "操作进度:已" & .Value & "条记录/总共" & mint流动的总记录数 & "条记录"
                       Me.Refresh
                   End If
               
               
               End If
    End Select
    End With
WayOut:
    cstu状态栏.Panels(1).Text = "已完成！"
    '恢复窗口
    Me.Enabled = True
    'cprg操作进度.Value = 0
    'cfra操作进度.Caption = "操作进度:"
    Set lobjTempIN = Nothing
    Set lobjTempOUT = Nothing
    MsgBox "完成！", vbInformation, Left(ctlb工具栏.Buttons.Item(3).Caption, 2)
    cprg操作进度.Value = 0
    cfra操作进度.Caption = "操作进度:"
    Exit Sub
errhandle:
    'sffuncMsg Err.Description
    'MsgBox "操作失败！", vbInformation, Left(Ctlb工具栏.Buttons.Item(3).Caption, 2)
    cprg操作进度.Value = 0
    cfra操作进度.Caption = "操作进度:"
    Me.Enabled = True
End Sub


'功能：检查输入的时间值是否正确
'输入：无
'输出：无
'返回：返回一个布尔值
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Function func验证时间(strTime As String) As Boolean
    strTime = Trim(strTime)
    If Not IsDate(strTime) Then
        func验证时间 = False
        Exit Function
    End If
    If InStr(1, strTime, ":") = 0 Then
        func验证时间 = False
        Exit Function
    Else
        func验证时间 = True
    End If
    
End Function


'功能： 根据查询的结果填充表格
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub sub填充表格(ByVal Str采集项目 As String, ByVal bln方向标志 As Boolean, Optional ByVal Str条件 As String)
Dim lrsTemp As ADODB.Recordset
Dim i As Integer
Dim j As Integer

    On Error GoTo errhandle
    Set lrsTemp = Func采集数据(Str采集项目, bln方向标志, Str条件)
    If Not (lrsTemp Is Nothing) Then
        If lrsTemp Is Nothing Then Exit Sub
        If lrsTemp.RecordCount <= 0 Then Exit Sub
        Cgrd记录显示.Clear   '清空表格
        Cgrd记录显示.Rows = lrsTemp.RecordCount + 1
        Cgrd记录显示.Cols = lrsTemp.Fields.Count
        If Cgrd记录显示.Rows < 27 Then Cgrd记录显示.Rows = 27
        If Cgrd记录显示.Cols < 5 Then Cgrd记录显示.Cols = 5
        '填充表格
        For i = 0 To lrsTemp.Fields.Count - 1
            Cgrd记录显示.TextMatrix(0, i) = lrsTemp.Fields(i).Name
        Next i
        lrsTemp.MoveFirst
        For i = 1 To lrsTemp.RecordCount
            For j = 0 To lrsTemp.Fields.Count - 1
                Cgrd记录显示.TextMatrix(i, j) = IIf(IsNull(lrsTemp(j).Value), "", lrsTemp(j).Value)
            Next j
            lrsTemp.MoveNext
        Next i
        lrsTemp.MoveFirst
        'Set Cgrd记录显示.DataSource = lrsTemp '填充表格
         
        clbl记录数.Caption = lrsTemp.RecordCount
        Set lrsTemp = Nothing
    End If
Exit Sub
errhandle:

End Sub

'功能： 统计要操作的总记录数和相应的记录数
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub sub统计记录数()
If cchk项目选择(费用信息).Value = 0 Then mint流动的费用信息记录数 = 0
If cchk项目选择(收费项目).Value = 0 Then mint流动的收费项目记录数 = 0
If cchk项目选择(收费标准).Value = 0 Then mint流动的收费标准记录数 = 0
If cchk项目选择(票据格式).Value = 0 Then mint流动的票据格式记录数 = 0
If cchk项目选择(打折情况).Value = 0 Then mint流动的打折情况记录数 = 0
If cchk项目选择(系统设置).Value = 0 Then mint流动的系统设置记录数 = 0
mint流动的总记录数 = mint流动的费用信息记录数 + mint流动的收费项目记录数 + mint流动的收费标准记录数 + mint流动的票据格式记录数 + _
                       mint流动的打折情况记录数 + mint流动的系统设置记录数
End Sub



'功能： 关闭窗口。
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub Form_Unload(Cancel As Integer)
    Set mobjGUI = Nothing
    pblnInUse = False
    Set pobj收费管理 = Nothing
    Set pobj业务设置 = Nothing
    Set pobj单位定位 = Nothing
End Sub

'功能：响应用户点击工具栏操作
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：林涛
'创建时间：2001-3-29
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
Dim i As Integer
Select Case Operate
        Case "清空"
        
            Cgrd记录显示.Clear
            Cgrd记录显示.Rows = 27
            Cgrd记录显示.Cols = 5
            clbl记录数.Caption = ""
            
            For i = 1 To 7
               cchk项目选择(i).Value = 0
            Next i
            
End Select
End Sub



Private Sub subEnabled(ByVal lbln是否业务数据 As Boolean)
Dim i As Integer
    If lbln是否业务数据 Then
        For i = 费用信息 To 所有信息
            cchk项目选择(i).Enabled = True
            cchk系统信息.Enabled = False
            cchk按时间查询.Enabled = True
            cdtp日期(开始).Enabled = True
            cdtp日期(结束).Enabled = True
            ctxt时间(开始).Enabled = True
            ctxt时间(结束).Enabled = True
        Next i
    Else
        For i = 费用信息 To 所有信息
            cchk项目选择(i).Enabled = False
            cchk系统信息.Enabled = True
            cchk按时间查询.Enabled = False
            cdtp日期(开始).Enabled = False
            cdtp日期(结束).Enabled = False
            ctxt时间(开始).Enabled = False
            ctxt时间(结束).Enabled = False
        Next i
        
    End If
End Sub

Private Sub subCopyFile()
On Error GoTo errhandle
    Dim lstrFile As String
    lstrFile = Replace(App.Path, "收费管理", "公用组件") & mstrMDBFile
    CopyFile lstrFile, App.Path & mstrMDBFile, 0
    Exit Sub
errhandle:
    
End Sub
