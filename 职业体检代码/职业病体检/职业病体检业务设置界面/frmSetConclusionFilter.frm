VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Begin VB.Form frmSetConclusionFilter 
   Caption         =   "体检结论判断条件设置"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11070
   Icon            =   "frmSetConclusionFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11070
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退出(&X)"
      Height          =   400
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton ccmdUpdate 
      Caption         =   "修改(&M)"
      Enabled         =   0   'False
      Height          =   426
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton ccmdAdd 
      Caption         =   "新增(&A)"
      Height          =   426
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1275
   End
   Begin VB.CommandButton ccmdDelete 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   426
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "在下面录入一个体检结论的判断条件"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   10815
      Begin VB.CommandButton ccmdRemoveRow 
         Caption         =   "删除行(&R)"
         Enabled         =   0   'False
         Height          =   427
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3000
         Width           =   1275
      End
      Begin VB.ComboBox ccmbOperator 
         Height          =   300
         Left            =   5500
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton ccmdOk 
         Caption         =   "保存(&S)"
         Enabled         =   0   'False
         Height          =   427
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton ccmdCancel 
         Caption         =   "取消(&C)"
         Enabled         =   0   'False
         Height          =   427
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2400
         Width           =   1275
      End
      Begin 录入控件.ctlInputGrid cgrdInput 
         Height          =   1935
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3413
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
         Cols            =   3
         Rows            =   1
         Count           =   0
         Rows            =   1
         Cols            =   3
      End
      Begin 录入控件.ctlInputFrame cifFilter 
         Height          =   1725
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3043
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Rows            =   2
         Cols            =   17
         DistanceofRow   =   0
         BorderStyle     =   0
         FormatString    =   "编号,1,0,2,体检结论,1,3,8,描述,1,12,8,序号,2,0,1,体检项目,2,2,8,判断条件,2,11,2,判断值,2,14,4"
         Count           =   7
         titleInputBox0001=   "编号"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   2
         orderInputBox0001=   1
         valueInputBox0001=   ""
         datatypeInputBox0001=   3
         colInputBox0001 =   0
         rowInputBox0001 =   1
         PassWordCharInputBox0001=   0   'False
         主键InputBox0001=   0   'False
         允许等于最大值InputBox0001=   0   'False
         允许等于最小值InputBox0001=   0   'False
         字典名称InputBox0001=   ""
         显示字典字段InputBox0001=   ""
         保存字典字段InputBox0001=   ""
         名称InputBox0001=   "编号"
         缺省值InputBox0001=   ""
         保存缺省值InputBox0001=   ""
         长度InputBox0001=   4
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         EnableInputBox0001=   0   'False
         允许多选InputBox0001=   0   'False
         titleInputBox0002=   "体检结论"
         statusinfoInputBox0002=   ""
         lengthInputBox0002=   8
         orderInputBox0002=   2
         valueInputBox0002=   ""
         datatypeInputBox0002=   3
         colInputBox0002 =   3
         rowInputBox0002 =   1
         PassWordCharInputBox0002=   0   'False
         主键InputBox0002=   0   'False
         允许等于最大值InputBox0002=   0   'False
         允许等于最小值InputBox0002=   0   'False
         字典名称InputBox0002=   "体检结论字典"
         显示字典字段InputBox0002=   "名称"
         保存字典字段InputBox0002=   "InnerID"
         名称InputBox0002=   "体检结论"
         缺省值InputBox0002=   ""
         保存缺省值InputBox0002=   ""
         长度InputBox0002=   50
         MaxInputBox0002 =   ""
         MinInputBox0002 =   ""
         VisibleInputBox0002=   -1  'True
         PermitNullInputBox0002=   0   'False
         TriggerstrInputBox0002=   ""
         CheckInDictInputBox0002=   -1  'True
         允许多选InputBox0002=   0   'False
         titleInputBox0003=   "描述"
         statusinfoInputBox0003=   ""
         lengthInputBox0003=   8
         orderInputBox0003=   3
         valueInputBox0003=   ""
         datatypeInputBox0003=   3
         colInputBox0003 =   12
         rowInputBox0003 =   1
         PassWordCharInputBox0003=   0   'False
         主键InputBox0003=   0   'False
         允许等于最大值InputBox0003=   0   'False
         允许等于最小值InputBox0003=   0   'False
         字典名称InputBox0003=   ""
         显示字典字段InputBox0003=   ""
         保存字典字段InputBox0003=   ""
         名称InputBox0003=   "描述"
         缺省值InputBox0003=   ""
         保存缺省值InputBox0003=   ""
         长度InputBox0003=   100
         MaxInputBox0003 =   ""
         MinInputBox0003 =   ""
         VisibleInputBox0003=   -1  'True
         PermitNullInputBox0003=   -1  'True
         TriggerstrInputBox0003=   ""
         允许多选InputBox0003=   0   'False
         titleInputBox0004=   "序号"
         statusinfoInputBox0004=   ""
         lengthInputBox0004=   1
         orderInputBox0004=   4
         valueInputBox0004=   ""
         datatypeInputBox0004=   2
         colInputBox0004 =   0
         rowInputBox0004 =   2
         PassWordCharInputBox0004=   0   'False
         主键InputBox0004=   0   'False
         允许等于最大值InputBox0004=   0   'False
         允许等于最小值InputBox0004=   -1  'True
         字典名称InputBox0004=   ""
         显示字典字段InputBox0004=   ""
         保存字典字段InputBox0004=   ""
         名称InputBox0004=   "序号"
         缺省值InputBox0004=   ""
         保存缺省值InputBox0004=   ""
         长度InputBox0004=   4
         MaxInputBox0004 =   ""
         MinInputBox0004 =   "1"
         VisibleInputBox0004=   -1  'True
         PermitNullInputBox0004=   0   'False
         TriggerstrInputBox0004=   ""
         允许多选InputBox0004=   0   'False
         titleInputBox0005=   "体检项目"
         statusinfoInputBox0005=   ""
         lengthInputBox0005=   8
         orderInputBox0005=   5
         valueInputBox0005=   ""
         datatypeInputBox0005=   3
         colInputBox0005 =   2
         rowInputBox0005 =   2
         PassWordCharInputBox0005=   0   'False
         主键InputBox0005=   0   'False
         允许等于最大值InputBox0005=   0   'False
         允许等于最小值InputBox0005=   0   'False
         字典名称InputBox0005=   "体检项目字典"
         显示字典字段InputBox0005=   "名称"
         保存字典字段InputBox0005=   "编码"
         名称InputBox0005=   "体检项目"
         缺省值InputBox0005=   ""
         保存缺省值InputBox0005=   ""
         长度InputBox0005=   50
         MaxInputBox0005 =   ""
         MinInputBox0005 =   ""
         VisibleInputBox0005=   -1  'True
         PermitNullInputBox0005=   0   'False
         TriggerstrInputBox0005=   ""
         CheckInDictInputBox0005=   -1  'True
         允许多选InputBox0005=   0   'False
         titleInputBox0006=   "判断条件"
         statusinfoInputBox0006=   ""
         lengthInputBox0006=   2
         orderInputBox0006=   6
         valueInputBox0006=   ""
         datatypeInputBox0006=   0
         colInputBox0006 =   11
         rowInputBox0006 =   2
         PassWordCharInputBox0006=   0   'False
         主键InputBox0006=   0   'False
         允许等于最大值InputBox0006=   0   'False
         允许等于最小值InputBox0006=   0   'False
         字典名称InputBox0006=   ""
         显示字典字段InputBox0006=   ""
         保存字典字段InputBox0006=   ""
         名称InputBox0006=   "判断条件"
         缺省值InputBox0006=   "="
         保存缺省值InputBox0006=   "="
         长度InputBox0006=   10
         MaxInputBox0006 =   ""
         MinInputBox0006 =   ""
         VisibleInputBox0006=   -1  'True
         PermitNullInputBox0006=   0   'False
         TriggerstrInputBox0006=   ""
         允许多选InputBox0006=   0   'False
         titleInputBox0007=   "判断值"
         statusinfoInputBox0007=   ""
         lengthInputBox0007=   4
         orderInputBox0007=   7
         valueInputBox0007=   ""
         datatypeInputBox0007=   3
         colInputBox0007 =   14
         rowInputBox0007 =   2
         PassWordCharInputBox0007=   0   'False
         主键InputBox0007=   0   'False
         允许等于最大值InputBox0007=   0   'False
         允许等于最小值InputBox0007=   0   'False
         字典名称InputBox0007=   ""
         显示字典字段InputBox0007=   ""
         保存字典字段InputBox0007=   ""
         名称InputBox0007=   "判断值"
         缺省值InputBox0007=   ""
         保存缺省值InputBox0007=   ""
         长度InputBox0007=   50
         MaxInputBox0007 =   ""
         MinInputBox0007 =   ""
         VisibleInputBox0007=   -1  'True
         PermitNullInputBox0007=   0   'False
         TriggerstrInputBox0007=   ""
         允许多选InputBox0007=   0   'False
         ErrColor        =   12648447
      End
      Begin 录入控件.ctlInputDictGrid cidgMain 
         Height          =   3495
         Left            =   6120
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   6165
         Cols            =   10
         Count           =   0
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
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      _cx             =   23740357
      _cy             =   23728503
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "^编号 |^体检结论                     |^描述                                 "
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "      (3) 一个体检结论在上面网格中出现两次，表示该结论有两种独立的判断方法。"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   6840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "      (2) 一个体检结论的一种判断条件可以由多个子条件组成（下面网格中一行代表一个子条件），子条件之间是并且关系。"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   10080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明：(1) 设置体检结论的判断条件后，医师在填写完毕体检人员所有体检项目的结果后，系统会据此条件自动得出体检结论。"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   6960
      Width           =   10080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已有的体检结论判断分组："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2160
   End
End
Attribute VB_Name = "frmSetConclusionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春

Private WithEvents mobj界面通用对象 As cls界面通用对象
Attribute mobj界面通用对象.VB_VarHelpID = -1

Private mobj体检结论条件 As Object 'clsConclusionFilter

Private Sub cgrdMain_AfterSort(ByVal Col As Long, Order As Integer)
    On Error Resume Next
    cgrdMain.Row = 0
    ccmdDelete.Enabled = False
    ccmdUpdate.Enabled = False
    
End Sub

'编号,1,0,2,体检结论,1,3,8,描述,1,12,8,序号,2,0,1,体检项目,2,2,8,判断条件,2,11,3,判断值,2,15,3
Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lcolInfo As New Collection
    Dim lcolItem As Collection
    Dim i As Long
    
    On Error GoTo errHandler
    
    '获取并显示所有体检结论判断条件分组情况。
    Set lobjRec = pobj业务对象.所有体检结论条件
    gfsubLoadGridFromRec cgrdMain, lobjRec
    If cgrdMain.Rows > 1 Then cgrdMain.Rows = cgrdMain.Rows - 1
    
    '创建全局变量mobj界面通用对象。
    Set mobj界面通用对象 = New cls界面通用对象
    
    With mobj界面通用对象
        Set .Form = Me
        Set .c录入板 = cifFilter
        Set .c记录表 = cgrdInput
        Set .c字典表 = cidgMain
        .pint详细信息开始编号 = 4
        
        .subInitialize lcolInfo, ""
    End With
    
    '创建全局变量“mobj体检结论条件”。
    Set mobj体检结论条件 = CreateObject("职业病对象.clsConclusionFilter")
        
    '获取所有可选的判断条件：[符号，说明]。
    Set lcolInfo = mobj体检结论条件.判断条件枚举
    
    '获取体检结论字典。
    Set lobjRec = pobjDict.Fetch("体检结论字典视图")
    
    '设置录入板“体检结论”录入框的字典内容。
    Set cifFilter.InfoCollection(2).DictRecordSet = lobjRec
    
    '初始化体检录入板的判断条件字典。
    ccmbOperator.Clear
    For i = 1 To lcolInfo.Count
        ccmbOperator.AddItem lcolInfo(i)("符号")
    Next
    
    '获取所有体检项目。
    Dim lobj体检项目集 As Object
    Set lobj体检项目集 = CreateObject("职业病对象.clsTestItemSet")
    Set lobjRec = lobj体检项目集.体检项目
    
    '设置录入板“体检项目”录入框的字典内容。
    Set cifFilter.InfoCollection(5).DictRecordSet = lobjRec
    
    cifFilter.Enabled = False
    cgrdInput.Enabled = False
    ccmdOk.Enabled = False
    ccmdCancel.Enabled = False
    ccmdRemoveRow.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "Form_Load", 6666, lstrError, False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Or Chr(KeyAscii) = "，" Then
        '不允许输入“'”和“，”。
        KeyAscii = 0
    End If

End Sub
Private Sub ccmbOperator_Click()
    On Error Resume Next
    cifFilter.ItemText(5) = ccmbOperator.Text
End Sub

Private Sub ccmbOperator_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        cifFilter.SetFocus
        cifFilter.ItemSetfocus 6
    ElseIf KeyCode = vbKeyTab Then
        cifFilter.SetFocus
        cifFilter.ItemSetfocus 6
    ElseIf KeyCode = vbKeyTab And (Shift And vbShiftMask = vbShiftMask) Then
        cifFilter.ItemSetfocus 5
    End If
End Sub

Private Sub ccmbOperator_LostFocus()
    On Error Resume Next
    ccmbOperator.Visible = False
End Sub

Private Sub ccmdRemoveRow_Click()
    Dim lblnSuc As Boolean '删除是否成功。
    On Error GoTo errHandler
    
    '删除对象的。
    If cgrdInput.Row > 0 And Val(cifFilter.ItemText(3)) <> 0 Then
        mobj体检结论条件.subRemoveFilter cifFilter.ItemText(3)
    
        '删除界面网格上的。
        mobj界面通用对象.subOperate optDELETE, lblnSuc
    Else
        sffuncMsg "请先选择要删除的行！"
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "ccmdRemoveRow_Click", 6666, lstrError, False
End Sub

Private Sub cifFilter_ItemGetFocus(Index As Integer)
    On Error GoTo errHandler
    If Index = 5 Then
        ccmbOperator.Visible = True
        ccmbOperator.SetFocus
    End If
    Exit Sub
errHandler:
End Sub
Private Sub cifFilter_ItemLostFocus(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        mobj界面通用对象_ItemLostFocus Index, "体检结论", cifFilter.ItemText(Index), cifFilter.ItemTrueText(Index), False
    End If
End Sub


Private Sub ccmdCancel_Click()
    On Error GoTo errHandler
    
    '若是修改，重新在录入区显示cgrdMain当前行内容；若是新建，清空录入区。
    If cgrdMain.Row > 0 Then
        cgrdMain_Click
    Else
        subClear
    End If
    
    '调用"subReset"恢复界面。
    subReset
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "ccmdCancel_Click", 6666, lstrError, False
End Sub

Private Sub ccmdDelete_Click()
    On Error GoTo errHandler
    '询问。
    If sffuncMsg("你确认要删除“" & cgrdMain.TextMatrix(cgrdMain.Row, 2) & "”的判断条件吗？", sf询问) Then
    
        '删除当前选号编号分组的所有判断条件。
        mobj体检结论条件.subDelete
        
        '清空界面。
        subClear
        cgrdMain.RemoveItem cgrdMain.Row
        
        ccmdDelete.Enabled = False
        ccmdUpdate.Enabled = False
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "ccmdDelete_Click", 6666, lstrError, False
End Sub

Private Sub ccmdOk_Click()
    Dim lblnIsExist As Boolean '标志当前分组是否是新建。
    Dim llngRow As Long        'cgrdMain网格当前选中的行号。
    
    On Error GoTo errHandler
    
    If mobj体检结论条件.ID <> 0 And mobj体检结论条件.判定条件.Count > 0 Then
        lblnIsExist = mobj体检结论条件.是否已存在
        '保存当前正在修改或新增的条件。
        mobj体检结论条件.描述 = cifFilter.ItemText(2)
        mobj体检结论条件.subSave
    
        '若是新建，在cgrdMain中显示新建的体检结论条件分组情况；
        '否则，按新录入信息修改cgrdMain当前行；
        If lblnIsExist Then
            '修改行。
            llngRow = cgrdMain.Row
            cgrdMain.TextMatrix(llngRow, 1) = cifFilter.ItemTrueText(1)
            cgrdMain.TextMatrix(llngRow, 2) = cifFilter.ItemText(1)
            cgrdMain.TextMatrix(llngRow, 3) = cifFilter.ItemText(2)
        Else
            cifFilter.ItemText(0) = mobj体检结论条件.编号
            '添加行。
            cgrdMain.AddItem cifFilter.ItemText(0) & vbTab & cifFilter.ItemTrueText(1) & vbTab & cifFilter.ItemText(1) & vbTab & cifFilter.ItemText(2)
            If cgrdMain.Row > 0 Then
                cgrdMain_Click
            End If
                
        End If
    Else
        Err.Raise 6666, , "系统无法保存！因为：" & Chr(13) & Chr(10) & "必须选择体检结论、录入判断条件！并且在录入完毕一个子条件后，一定要在最后一项按回车键，保证录入的子条件出现在下面的网格中。"
    End If
    
    '恢复界面。
    subReset
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "ccmdOk_Click", 6666, lstrError, False
End Sub

Private Sub ccmdUpdate_Click()
    On Error GoTo errHandler
    '所有录入区可用，界面上只有"确定"、取消、退出按钮可用，焦点自动到"体检结论"录入框。
    subBeginEdit
    
    '体检结论不可修改。
    cifFilter.ItemEnable(1) = False
    cifFilter.ItemEnable(2) = True
    
    '焦点自动变到"描述"录入框。
    cifFilter.ItemSetfocus 3
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "ccmdUpdate_Click", 6666, lstrError, False
End Sub

Private Sub cgrdMain_Click()

    On Error GoTo errHandler
    If cgrdMain.Row > 0 Then
        ccmdDelete.Enabled = True
    
        subShowHistory
    
        ccmdUpdate.Enabled = True
    Else
        ccmdDelete.Enabled = False
        ccmdUpdate.Enabled = False
    End If
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "cgrdMain_Click", 6666, lstrError, False
End Sub
Private Sub subShowHistory()
    Dim llng编号 As Long
    Dim lcol条件 As Collection
    Dim lcolFilterItem As Variant
    
    Dim lcol基本信息 As Collection
    
    Dim lcol详细内容 As Collection
    Dim lcolRow As Collection
    Dim lcolItem As Collection
    
    On Error GoTo errHandler
    
    If cgrdMain.Row < 1 Then Exit Sub
    cgrdInput.Rows = 1
    cifFilter.ClearContent
    
    '设置"mobj体检结论条件"的属性"ID"、"编号"。
    llng编号 = cgrdMain.TextMatrix(cgrdMain.Row, 0)
    mobj体检结论条件.subClear
    mobj体检结论条件.编号 = llng编号
    mobj体检结论条件.ID = cgrdMain.TextMatrix(cgrdMain.Row, 1)
    
    '获取当前分组的所有条件。
    Set lcol条件 = mobj体检结论条件.判定条件
    
    '根据对象属性在录入区显示该体检结论的判断条件。
    Set lcol基本信息 = New Collection
    Set lcolItem = New Collection
    lcolItem.Add llng编号, "显示内容"
    lcolItem.Add llng编号, "保存内容"
    lcol基本信息.Add lcolItem, "编号"
    Set lcolItem = New Collection
    lcolItem.Add mobj体检结论条件.体检结论名称, "显示内容"
    lcolItem.Add mobj体检结论条件.ID, "保存内容"
    lcol基本信息.Add lcolItem, "体检结论"
    Set lcolItem = New Collection
    lcolItem.Add mobj体检结论条件.描述, "显示内容"
    lcolItem.Add mobj体检结论条件.描述, "保存内容"
    lcol基本信息.Add lcolItem, "描述"
    
    '把条件按通用录入控件的要求加入集合中。
    Set lcol详细内容 = New Collection
    For Each lcolFilterItem In lcol条件
        Set lcolRow = New Collection
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("序号"), "显示内容"
        lcolItem.Add lcolFilterItem("序号"), "保存内容"
        lcolRow.Add lcolItem, "序号"
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("体检项目名称"), "显示内容"
        lcolItem.Add lcolFilterItem("体检项目"), "保存内容"
        lcolRow.Add lcolItem, "体检项目"
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("判断条件"), "显示内容"
        lcolItem.Add lcolFilterItem("判断条件"), "保存内容"
        lcolRow.Add lcolItem, "判断条件"
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("标准值"), "显示内容"
        lcolItem.Add lcolFilterItem("标准值"), "保存内容"
        lcolRow.Add lcolItem, "判断值"
        
        lcol详细内容.Add lcolRow
    Next
    
    '在cgrdInput中显示所有已存在的条件。
    Set mobj界面通用对象.pcol基本信息 = lcol基本信息
    Set mobj界面通用对象.pcol详细信息 = lcol详细内容
    mobj界面通用对象.sub把集合内容填入录入板
    
    '修改、删除按钮可用。
    ccmdUpdate.Enabled = True
    ccmdDelete.Enabled = True
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "cgrdMain_Click", Err.Number, Err.Description, True
End Sub

Private Sub ccmdAdd_Click()
    On Error GoTo errHandler
    
    '设置对项“mobj体检结论条件.编号”为特殊值0，表示开始新增。
    cgrdMain.Row = 0
    mobj体检结论条件.编号 = 0
    
    '清空录入区。
    subClear
    
    '设置界面上只有"确定"、"取消"、"退出"按钮可用，所有录入区可用。
    subBeginEdit
    
    '体检结论可以录入。
    cifFilter.ItemEnable(1) = True
    cifFilter.ItemEnable(2) = True
    
    '焦点到“体检结论”。
    cifFilter.ItemSetfocus 1
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "ccmdAdd_Click", 6666, lstrError, False
End Sub

Private Sub ccmdExit_Click()
    On Error Resume Next
    Set mobj界面通用对象.Form = Nothing
    Unload Me
End Sub

'功能：清空录入区。
Private Sub subClear()
    On Error Resume Next
    cifFilter.ClearContent
    cifFilter.InfoCollection.ClearHistory
    Set cgrdInput.InfoCollection = cifFilter.InfoCollection
End Sub

'功能：设置界面上控件的状态，准备录入。
Private Sub subBeginEdit()
    On Error GoTo errHandler

    '界面上只有录入区、确定、取消、退出可用。
    cgrdMain.Enabled = False
    ccmdAdd.Enabled = False
    ccmdDelete.Enabled = False
    ccmdUpdate.Enabled = False
    ccmdExit.Enabled = True
    
    '录入区可用。
    Frame1.Enabled = True
    cifFilter.Enabled = True
    cidgMain.Enabled = True
    cgrdInput.Enabled = True
    ccmdOk.Enabled = True
    ccmdCancel.Enabled = True
    ccmdRemoveRow.Enabled = True
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "subBeginEdit", Err.Number, Err.Description, True
End Sub

'功能：恢复界面上控件状态，不能录入。
Private Sub subReset()
    On Error Resume Next
    cgrdMain.Row = 0
    
    '界面上只有主网格、新增、退出可用。
    cgrdMain.Enabled = True
    ccmdAdd.Enabled = True
    If cgrdMain.Row > 0 Then
        ccmdDelete.Enabled = True
        ccmdUpdate.Enabled = True
    Else
        ccmdDelete.Enabled = False
        ccmdUpdate.Enabled = False
    End If
    ccmdExit.Enabled = True
        
    '录入区不可用。
    Frame1.Enabled = False
    cifFilter.Enabled = False
    cidgMain.Enabled = False
    cgrdInput.Enabled = False
    ccmdOk.Enabled = False
    ccmdCancel.Enabled = False
    ccmdRemoveRow.Enabled = False
    
    mobj体检结论条件.编号 = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj界面通用对象 = Nothing
    Set mobj体检结论条件 = Nothing
    
End Sub

'功能：录入子条件完毕，加入“mobj体检结论条件”的属性中。
Private Sub mobj界面通用对象_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lstrError  As String
    Dim i As Long
    
    On Error GoTo errHandler
    '录入一子条件完毕，加入对象中。
    Select Case Operate
    Case "添加"
        If cifFilter.ItemsError.Count > 0 Then
            Err.Raise 6666, , "录入还有错误，请更正黄色框内容。"
        End If
        '查找条件网格中是否已存在该序号。
        For i = 1 To cgrdInput.Rows - 1
            If cgrdInput.TextMatrix(i, 3) = cifFilter.ItemTrueText(3) Then
                lstrError = "序号不允许重复。" & Chr(13) & Chr(10)
                Exit For
            End If
        Next
        '修改：2001-8-24（判断值是否合法）。
        Select Case cifFilter.Box1("判断条件").Text
        Case "<", ">", "<=", ">="
            '判断值必须是数字。
            If Not IsNumeric(cifFilter.Box1("判断值").Text) Then
                lstrError = lstrError & "判断条件为< (或>, <=, >=) 时，判断值必须输入数值型。"
            End If
        End Select
        If lstrError <> "" Then
            Err.Raise 6666, , "录入数据不符合系统规定，无法添加：" & Chr(13) & Chr(10) & lstrError
        End If
        
        '输入：序号，体检项目，体检项目名称，判断条件，判断值。
        mobj体检结论条件.subAddFilter cifFilter.ItemTrueText(3), cifFilter.ItemTrueText(4), cifFilter.ItemText(4), cifFilter.ItemText(5), cifFilter.ItemText(6)
        
    Case "修改"
        '修改：2001-8-24（判断值是否合法）。
        Select Case cifFilter.Box1("判断条件").Text
        Case "<", ">", "<=", ">="
            '判断值必须是数字。
            If Not IsNumeric(cifFilter.Box1("判断值").Text) Then
                Err.Raise 6666, , "录入数据不符合系统规定，无法添加：" & Chr(13) & Chr(10) & "判断条件为< (或>, <=, >=) 时，判断值必须输入数值型。"
            End If
        End Select
        '先删除旧的。
        If cgrdInput.Row > 0 Then
            mobj体检结论条件.subRemoveFilter cifFilter.ItemHistory(cgrdInput.Row)(4)
        End If
        '查找条件网格中是否已存在该序号。
        For i = 1 To cgrdInput.Rows - 1
            If cgrdInput.TextMatrix(i, 3) = cifFilter.ItemTrueText(3) And i <> cgrdInput.Row Then
                Err.Raise 6666, , "序号不允许重复。"
            End If
        Next
        '再添加。
        mobj体检结论条件.subAddFilter cifFilter.ItemTrueText(3), cifFilter.ItemTrueText(4), cifFilter.ItemText(4), cifFilter.ItemText(5), cifFilter.ItemText(6)

    Case "退出"
        Set mobj界面通用对象.Form = Nothing
        Unload Me
    End Select
    
    Exit Sub
errHandler:
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "mobj界面通用对象_BeforeOperate", 6666, lstrError, False
    cifFilter.ItemSetfocus 3
    Cancel = True
End Sub

Private Sub mobj界面通用对象_ItemLostFocus(ByVal Index As Integer, ByVal 名称 As String, ByVal 内容 As String, ByVal 保存内容 As String, ByVal IsError As Boolean)
    Dim i As Long
    Dim Row As Integer
    
    On Error GoTo errHandler
    If 内容 = "" Or (ActiveControl.Name <> "cifFilter") Then Exit Sub
    
    Select Case 名称
    Case "体检结论"
        If mobj体检结论条件.ID <> 保存内容 Then
            mobj体检结论条件.ID = 保存内容
        End If
    Case "描述"
        mobj体检结论条件.描述 = 保存内容
    Case "序号"
    End Select

    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetConclusionFilter", "mobj界面通用对象_ItemLostFocus", 6666, lstrError, False
    cifFilter.ItemSetfocus Index
End Sub
