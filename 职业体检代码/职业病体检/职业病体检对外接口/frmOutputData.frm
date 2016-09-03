VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutputData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "站外体检数据导出"
   ClientHeight    =   7800
   ClientLeft      =   1395
   ClientTop       =   1215
   ClientWidth     =   11055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1111
      ButtonWidth     =   900
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   8040
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "预览数据"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   3735
      Left            =   120
      TabIndex        =   19
      Top             =   3280
      Width           =   10695
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdPreviewData 
         Height          =   3315
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5847
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
         BackColorAlternate=   14737632
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
         Cols            =   12
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
      End
   End
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "浏览文件(&B)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox ctxtOutputDestination 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin VB.Frame cfraSelectInputData 
      Caption         =   "选择要导出的内容"
      ForeColor       =   &H00800000&
      Height          =   2025
      Left            =   7440
      TabIndex        =   15
      Top             =   1200
      Width           =   3375
      Begin VB.ListBox clstDataType 
         Height          =   1740
         ItemData        =   "frmOutputData.frx":0000
         Left            =   120
         List            =   "frmOutputData.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame cfrafiltrateCondition 
      Caption         =   "数据导出条件"
      ForeColor       =   &H00800000&
      Height          =   2025
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   7005
      Begin VB.CheckBox cchkOver 
         Caption         =   "已体检完毕"
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox cchkTemplate 
         Caption         =   "体检对象"
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1035
      End
      Begin VB.TextBox ctxtEndCode 
         Height          =   315
         Left            =   4320
         TabIndex        =   10
         Top             =   1080
         Width           =   2475
      End
      Begin VB.TextBox ctxtBeginCode 
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox ctxtUnit 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Top             =   690
         Width           =   4155
      End
      Begin VB.CheckBox cchkSystemCode 
         Caption         =   "系统编号"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1140
         Width           =   1035
      End
      Begin VB.CheckBox cchkUnitName 
         Caption         =   "单位名称"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox cchkMedicalDate 
         Caption         =   "体检日期"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "单位定位(&L)"
         Height          =   375
         Left            =   5580
         TabIndex        =   7
         Top             =   660
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker cdtpBeginDate 
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
         Left            =   1380
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         Format          =   23658496
         CurrentDate     =   36951
      End
      Begin MSComCtl2.DTPicker cdtpEndDate 
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
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         Format          =   23658496
         CurrentDate     =   36951
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "到"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3240
         TabIndex        =   21
         Top             =   270
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "到"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4020
         TabIndex        =   14
         Top             =   1200
         Width           =   180
      End
   End
   Begin MSComctlLib.ProgressBar cprgDatatranform 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7425
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19447
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   6960
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "本操作适用于体检与办健康证业务系统分家的情况，以便分家后的健康证系统导入要办证的从业人员体检情况。"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   7080
      Width           =   10095
   End
   Begin VB.Label clabOutputDestintion 
      AutoSize        =   -1  'True
      Caption         =   "导出文件："
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "frmOutputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjGUI  As cls界面通用对象 '用于初始化工具栏。
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj对外接口 As ClsManageTransmission '体检对外接口对象

Private mstr系统编号固定部分 As String

Public pblnInUse As Boolean

Private Sub ctxtBeginCode_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtEndCode.SetFocus
    End If

End Sub

Private Sub Form_Load()
    Dim lcolInfo As New Collection
    Dim lobjRec As Object
    Dim i As Integer
    On Error GoTo errHandler
    pblnInUse = True
    
    '创建界面通用对象，初始化工具栏。
    Set mobjGUI = New cls界面通用对象
    With lcolInfo
        .Add "预览(&R)108"
        .Add "导出(&E)113"
        .Add "|"
        .Add "退出"
    End With
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
    End With
    mobjGUI.subInitialize lcolInfo, ""
    
    '创建体检对外接口对象。
    Set mobj对外接口 = CreateObject("体检对外接口部件.ClsManageTransmission")
    
    '获取保存在工作站配置文件中记录的上次导出的文件。
    ctxtOutputDestination.Text = mobj对外接口.工作站配置.内部导出文件
    
    '获取所有可能导入的数据分类。
    Set lobjRec = mobj对外接口.所有数据分类清单
    While Not lobjRec.EOF
        clstDataType.AddItem lobjRec.Fields("数据分类名")
        lobjRec.MoveNext
    Wend
    If clstDataType.ListCount > 0 Then
        clstDataType.Selected(0) = True
    End If
    '获取所有体检表名称。
    Dim lobj体检表模板集 As Object
    Set lobj体检表模板集 = CreateObject("体检对象部件.clsMedicalExamTemplateSet")
    Set lcolInfo = lobj体检表模板集.元素集
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
    Next
    Set lobj体检表模板集 = Nothing
    
    '缺省选择第一个。
    If ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.ListIndex = 0
    End If
    
       
    '在其它条件的check框未被选中时,输入框不可。.
    ctxtBeginCode.Enabled = False
    ctxtEndCode.Enabled = False
    ctxtUnit.Enabled = False
    ccmbTemplate.Enabled = False
    ccmdLocateUnit.Enabled = False
        
    '设置日期的初始值。
    cdtpBeginDate.Value = Format(Date, "yyyy-mm-dd")
    cdtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
    '导入,预览按钮在选定文件后才变为可用.
    If Len(ctxtOutputDestination.Text) = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    
    '初始时，按钮不可用。
    If ctxtOutputDestination.Text = "" Then
        ctbMain.Buttons(1).Enabled = False
    End If
    
    '获取系统编号固定部分。
    Dim lobj体检 As Object '体检对象，获取系统编号的固定部分。
    Set lobj体检 = CreateObject("体检对象部件.clsMedicalExam")
    mstr系统编号固定部分 = lobj体检.系统编号固定部分
    Set lobj体检 = Nothing
    
    ctbMain.Buttons(2).Enabled = False
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmOutputData", "form_load", Err.Number, Err.Description, False
End Sub

Private Sub ctxtBeginCode_GotFocus()
    On Error Resume Next
    If Trim(ctxtBeginCode) = "" Then
        ctxtBeginCode.Text = mstr系统编号固定部分
        ctxtBeginCode.SelStart = Len(ctxtBeginCode)
        ctxtBeginCode.SelLength = 0
    End If
End Sub

Private Sub ctxtEndCode_GotFocus()
    On Error Resume Next
    If Trim(ctxtEndCode) = "" Then
        ctxtEndCode.Text = mstr系统编号固定部分
        ctxtEndCode.SelStart = Len(ctxtEndCode)
        ctxtEndCode.SelLength = 0
    End If

End Sub

Private Sub ctxtOutputDestination_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If cchkMedicalDate.Value = 1 Then
            cdtpBeginDate.SetFocus
        ElseIf cchkSystemCode.Value = 1 Then
            ctxtBeginCode.SetFocus
        ElseIf cchkUnitName.Value = 1 Then
            ctxtUnit.SetFocus
        Else
            clstDataType.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub
Private Sub Form_Activate()
    On Error Resume Next
    ctxtOutputDestination.SetFocus
End Sub

'功能： 点击体检日期check框时，设置日期输入控件的状态。
'作者： 刘浩
Private Sub cchkMedicalDate_Click()
    On Err GoTo errHandler
    If cchkMedicalDate.Value = 1 Then
        cdtpBeginDate.Enabled = True
        cdtpEndDate.Enabled = True
        cdtpBeginDate.SetFocus
    Else
        cdtpBeginDate.Enabled = False
        cdtpEndDate.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmOutputData", "cchkMedicalDate_Click", Err.Number, Err.Description, False
End Sub

'功能： 点击系统编号check框时，设置系统编号输入控件的状态。
'作者： 刘浩
Private Sub cchkSystemCode_Click()
    On Err GoTo errHandler
    If cchkSystemCode.Value = 1 Then
        ctxtBeginCode.Enabled = True
        ctxtEndCode.Enabled = True
        ctxtBeginCode.SetFocus
        ctxtBeginCode.SelStart = Len(ctxtBeginCode)
        ctxtBeginCode.SelLength = 0
        
    Else
        ctxtBeginCode.Enabled = False
        ctxtEndCode.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmOutputData", "cchkSystemCode_Click", Err.Number, Err.Description, False
End Sub

'功能： 点击单位名称check框时，设置单位名称输入控件的状态。
'作者： 刘浩
Private Sub cchkUnitName_Click()
    On Err GoTo errHandler
    If cchkUnitName.Value = 1 Then
        ctxtUnit.Enabled = True
        ccmdLocateUnit.Enabled = True
        ctxtUnit.SetFocus
    Else
        ctxtUnit.Enabled = False
        ccmdLocateUnit.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmOutputData", "cchkUnitName_Click", Err.Number, Err.Description, False
End Sub

'功能： 点击体检对象check框时，设置体检对象输入控件的状态。
'作者： 刘浩
Private Sub cchkTemplate_click()
    On Err GoTo errHandler
    If cchkTemplate.Value = 1 Then
        ccmbTemplate.Enabled = True
    Else
        ccmbTemplate.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmOutputData", "ccmdTemplate_click", Err.Number, Err.Description, False
End Sub

'功能: 弹出文件查找窗口.
Private Sub ccmdBrowse_Click() ' 设置“CancelError”为 True
    On Error GoTo errHandler
    ccdgBrowse.CancelError = True
    ' 设置标志
    ccdgBrowse.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    ccdgBrowse.Filter = "All Files (*.*)|*.*|Access file" & _
        "(*.mdb)|*.mdb|Batch Files (*.bat)|*.bat"
    ccdgBrowse.FilterIndex = 2
    ccdgBrowse.ShowOpen
    
    ctxtOutputDestination.Text = ccdgBrowse.FileName
    
    '判断输入的合法性。
    ctxtOutputDestination_LostFocus
    
    Exit Sub
errHandler:
    Exit Sub
End Sub

Private Sub ctxtOutputDestination_LostFocus()
    On Error GoTo errHandler
    
    '判断输入的目的文件是否是mdb文件。
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(1).Enabled = False
    If ctxtOutputDestination.Text <> "" Then
        If UCase(Right(Trim(ctxtOutputDestination.Text), 3)) <> "MDB" Then
            sffuncMsg "请输入合法的数据导出目地文件必须是mdb后缀！", sf警告
        
            ctxtOutputDestination.Text = ""
        Else
            ctbMain.Buttons(1).Enabled = True
            ctbMain.Buttons(2).Enabled = True
        End If
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmOutputData", "ctxtOutputDestination_Validate", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mobjGUI = Nothing
    Set mobj对外接口 = Nothing
End Sub

'功能:  导入,预览报盘文件中的数据.
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lobjRange As Collection
    Dim lobjRec As Object '用于浏览的数据
    Dim lcolType As Collection
    Dim i As Integer
    
    Select Case Operate
        Case "预览"
            '获取要预览的数据范围。
            ctxtOutputDestination.SetFocus
            Set lobjRange = funcCalCon
            
            '将取出的数据显示在预览数据框内
            cgrdPreviewData.Rows = 1
            Set lobjRec = mobj对外接口.Func查看数据(lobjRange, 1) '(0导入/ 1导出)
            gfsubLoadGridFromRec cgrdPreviewData, lobjRec, , "健康档案编号,系统编号,公民身份号码,姓名,性别,出生日期,单位名称,体检日期,体检表名称,体检结论,诊断和处理意见,体检医师"
            cgrdPreviewData.Rows = cgrdPreviewData.Rows - 1
            
            ctbMain.Buttons(2).Enabled = True
            Cancel = True
        Case "导出"
            cprgDatatranform.Value = 0
            cprgDatatranform.Visible = True
            MousePointer = 11
            csbMain.Panels(1) = "正在准备导出，请稍候..."
            
            ctbMain.Enabled = False
            cfrafiltrateCondition.Enabled = False
            cfraSelectInputData.Enabled = False
            DoEvents
            '从列表框取出需要导出的数据分类名
            Set lcolType = New Collection
            For i = 0 To clstDataType.ListCount - 1
                If clstDataType.Selected(i) Then
                    lcolType.Add clstDataType.List(i), clstDataType.List(i)
                End If
            Next i
            If lcolType.Count = 0 Then
                Err.Raise 6666, , "请选择要导出的数据分类！"
            End If
            
            '获取要预览的数据范围。
            Set lobjRange = funcCalCon
            
            '导出前，并拷贝文件。
            csbMain.Panels(1) = "正在拷贝文件，请稍候..."
            mobj对外接口.sub导出准备 ctxtOutputDestination.Text
            DoEvents
            
            '开始导出，并显示进度。
            csbMain.Panels(1) = "正在导出，请稍候..."
            mobj对外接口.Sub数据导出 lobjRange, lcolType, cprgDatatranform
            DoEvents
            
            cprgDatatranform.Visible = False
            ctbMain.Enabled = True
            cfrafiltrateCondition.Enabled = True
            cfraSelectInputData.Enabled = True
            csbMain.Panels(1) = "导出成功。"
            MousePointer = 0
            Cancel = True
    End Select
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmOutputData", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    If Operate = "导出" Then
        ctbMain.Enabled = True
        cfrafiltrateCondition.Enabled = True
        cfraSelectInputData.Enabled = True
        csbMain.Panels(1) = "导出失败。"
    End If
    cprgDatatranform.Visible = False
    MousePointer = 0
End Sub

'功能：返回通过单位定位得到的单位名称，显示在单位名称文本框中，各单位名称之单用英文逗号分隔。
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobj体检管理 As Object
    Dim lobjRec As Object
    Dim lstrUnit As String
    
    '创建体检管理业务对象。
    Set lobj体检管理 = CreateObject("体检对象部件.clsManageMedicalExam")
    
    '单位定位。
    Set lobjRec = lobj体检管理.func单位定位
    
    '把定位出的单位名称显示在单位名称录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            lstrUnit = lobjRec.Fields("单位名称").Value
        End If
    End If
    If Trim(lstrUnit) <> "" Then
        If Trim(ctxtUnit.Text) <> "" Then
            ctxtUnit.Text = Trim(ctxtUnit.Text) & "," & lstrUnit
        Else
            ctxtUnit.Text = lstrUnit
        End If
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub

'功能：查找单位名称集字串中的中文逗号，并用英文逗号替换掉。
Private Sub ctxtUnit_LostFocus()
    On Error GoTo errHandler
    Dim i As Integer
    Dim lstrUnit As String
    
    lstrUnit = ctxtUnit.Text
    For i = 1 To Len(lstrUnit)
        If Mid(lstrUnit, i, 1) = "，" Then
            Mid(lstrUnit, i, 1) = ","
        End If
    Next i
    ctxtUnit.Text = lstrUnit
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub

'功能：先从"数据导入条件"frame中指定的范围值推算导出数据过滤条件
Private Function funcCalCon() As Object
    Dim lobjRange As Collection '[数据范围名，数据范围值] "数据导入条件"
    Dim lcolItem As Collection
    On Error GoTo errHandler
    
    Set lobjRange = New Collection
    If cchkMedicalDate.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add Format(cdtpBeginDate.Value, "yyyy-mm-dd"), "数据范围值"
        lcolItem.Add "开始日期", "数据范围名"
        lobjRange.Add lcolItem, "开始日期"
                
        Set lcolItem = New Collection
        lcolItem.Add Format(cdtpEndDate.Value, "yyyy-mm-dd"), "数据范围值"
        lcolItem.Add "结束日期", "数据范围名"
        lobjRange.Add lcolItem, "结束日期"
    End If
            
    If cchkSystemCode.Value = 1 Then
        Set lcolItem = New Collection
        If Len(ctxtBeginCode.Text) > 0 Then
            lcolItem.Add ctxtBeginCode.Text, "数据范围值"
            lcolItem.Add "从系统编", "数据范围名"
            lobjRange.Add lcolItem, "从系统编号"
        End If
                
        If Len(ctxtEndCode.Text) > 0 Then
            Set lcolItem = New Collection
            lcolItem.Add ctxtEndCode.Text, "数据范围值"
            lcolItem.Add "到系统编", "数据范围名"
            lobjRange.Add lcolItem, "到系统编号"
        End If
    End If
            
    If cchkUnitName.Value = 1 And Len(ctxtUnit.Text) > 0 Then
        Set lcolItem = New Collection
        lcolItem.Add ctxtUnit.Text, "数据范围值"
        lcolItem.Add "单位名称集", "数据范围名"
        lobjRange.Add lcolItem, "单位名称集"
    End If
    
    If cchkTemplate.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add ccmbTemplate.List(ccmbTemplate.ListIndex), "数据范围值"
        lcolItem.Add "体检对象", "数据范围名"
        lobjRange.Add lcolItem, "体检对象"
    End If
    
    If cchkOver.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add True, "数据范围值"
        lcolItem.Add "已体检完毕", "数据范围名"
        lobjRange.Add lcolItem, "已体检完毕"
    End If
    Set funcCalCon = lobjRange
    Exit Function

errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "funcCalCon", Err.Number, Err.Description, True
End Function

