VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmInputDataFromExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "外单位数据报盘"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10470
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   9600
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar cstbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   7080
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18415
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1111
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar cprgDataTransform 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "浏览文件(&B)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   840
      Width           =   1275
   End
   Begin VB.Frame cfraPreview 
      Caption         =   "预览数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4155
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   10335
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdPreview 
         Height          =   3855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6800
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "单位名称    |姓名   |性别 |年龄 "
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
   Begin VB.Frame cfraSelectMedicalTemplate 
      Caption         =   "选择体检表"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1395
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4935
      Begin VB.ListBox clstTemplate 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4620
      End
   End
   Begin VB.TextBox ctxtDataSource 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   4035
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   8520
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label clabDataSource 
      AutoSize        =   -1  'True
      Caption         =   "数据来源："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmInputDataFromExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mobj对外接口 As ClsManageTransmission
Private WithEvents mobjGUI  As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Public pblnInUse As Boolean

'功能: 检查用户修改的数据是否合法.
Private Sub cgrdPreview_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    If Not IsNumeric(cgrdPreview.TextMatrix(Row, Col)) And Col = 3 Then
        sffuncMsg "年龄字段只能为数字", sf警告
    End If
    
    If cgrdPreview.TextMatrix(Row, Col) <> "男" And cgrdPreview.TextMatrix(Row, Col) <> "女" _
        And Len(cgrdPreview.TextMatrix(Row, Col)) <> 0 And Col = 2 Then
        sffuncMsg "性别字段输入有误", sf警告
    End If

End Sub

Private Sub ctxtDataSource_LostFocus()
    On Err GoTo errHandler
    If Len(ctxtDataSource.Text) <> 0 And Right(ctxtDataSource.Text, 3) = "xls" Then
        ctbMain.Buttons(1).Enabled = True
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputFromExcel", "ctxtDataSource_LostFocus", Err.Number, Err.Description, False
End Sub

'初始化界面,
Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection
    Dim lcolTemplateSet As Object
    Dim lcolInfo As Collection
    
    Dim i As Integer
    
    On Error GoTo errHandler
    pblnInUse = True
    
    Set mobjGUI = New cls界面通用对象
    With lcol工具栏按钮
        .Add "预览(&R)108"
        .Add "导入(&I)112"
        .Add "|"
        .Add "退出"
    End With
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
    End With
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    Set mobj对外接口 = CreateObject("体检对外接口部件.ClsManageTransmission")
    
    '初始化"ctxtDataSource"文件框.(excel文件路径)
    ctxtDataSource.Text = mobj对外接口.工作站配置.Excel文件
    
    '初始化选择体检表列表框
    Set lcolTemplateSet = CreateObject("体检对象部件.ClsMedicalExamTemplateSet")
    lcolTemplateSet.体检表类型 = 3        '类型为初检的体检表
    Set lcolInfo = lcolTemplateSet.元素集
    For i = 1 To lcolInfo.Count
        clstTemplate.AddItem lcolInfo(i)
    Next i
    
    '导入,预览按钮在选定文件后才变为可用.
    If Len(ctxtDataSource.Text) = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    ctbMain.Buttons(2).Enabled = False
    If clstTemplate.ListCount > 0 Then
        clstTemplate.ListIndex = 0
    End If
    cprgDataTransform.ZOrder 0
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputFromExcel", "Form_Load", Err.Number, Err.Description, False
End Sub

'功能: 弹出文件查找窗口.
Private Sub ccmdBrowse_Click() ' 设置“CancelError”为 True
    ccdgBrowse.CancelError = True
    On Error GoTo errHandler
    ' 设置标志
    ccdgBrowse.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    ccdgBrowse.Filter = "All Files (*.*)|*.*|Excel file" & _
        "(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
    ccdgBrowse.FilterIndex = 2
    ccdgBrowse.ShowOpen
    ctxtDataSource.Text = ccdgBrowse.FileName
    If Len(ctxtDataSource.Text) <> 0 Then
        ctbMain.Buttons(1).Enabled = True
    End If
    
    Exit Sub
errHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjGUI = Nothing
    Set mobj对外接口 = Nothing
    pblnInUse = False
End Sub

'功能:  导入,预览报盘文件中的数据.
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lcolItem As Collection  '保存一条需要被导入的数据
    Dim lcolInfo As Collection  '存放所有需要导入的数据
    Dim i As Integer
    Dim j As Integer            '记录从报盘文件中读出的空行的行数
    On Error GoTo errHandler
    MousePointer = 11
    
    Select Case Operate
        Case "预览"
            '判断用户输入的.xls文件是否存在。
            If Dir(ctxtDataSource.Text) = "" Then
                Err.Raise 6666, , "输入的EXCEL 文件不在，请重新输入！"
            End If
            cstbMain.Panels(1) = "正在预览外单位报盘文件中内容，请稍候..."
            MousePointer = 11
            '预览文件内容。
            Set lcolItem = mobj对外接口.Func读报盘文件中体检人员信息(ctxtDataSource.Text)
            With cgrdPreview
                .Rows = lcolItem.Count + 1
                .Cols = 4
                .TextMatrix(0, 0) = "单位名称"
                .TextMatrix(0, 1) = "姓名"
                .TextMatrix(0, 2) = "性别"
                .TextMatrix(0, 3) = "年龄"
            End With
            j = 1
            For i = 1 To lcolItem.Count
                If lcolItem(i)("姓名") = "" Then
                    '如果报盘文件中有空行，则跳过此行。
                Else
                    cgrdPreview.TextMatrix(j, 0) = lcolItem(i)("单位名称")
                    cgrdPreview.TextMatrix(j, 1) = lcolItem(i)("姓名")
                    cgrdPreview.TextMatrix(j, 2) = lcolItem(i)("性别")
                    cgrdPreview.TextMatrix(j, 3) = lcolItem(i)("年龄")
                End If
                j = j + 1
            Next
            cgrdPreview.Rows = j
            
            ctbMain.Buttons(2).Enabled = True
            cgrdPreview.Editable = True
            cstbMain.Panels(1) = ""
            MousePointer = 0
            Cancel = True
        Case "导入"
            '判断用户输入的.xls文件是否存在。
            If Dir(ctxtDataSource.Text) = "" Then
                Err.Raise 6666, , "输入的EXCEL 文件不在，请重新输入！"
            End If
            If cgrdPreview.Rows = 1 Then
                Err.Raise 6666, , "输入的EXCEL 文件中无内容可以导入！"
            End If
            cstbMain.Panels(1) = "正在导入网格中体检人员名单，请稍候..."
            cprgDataTransform.Visible = True
            '导入预览数据框内的数据.
            Set lcolInfo = New Collection
            For i = 1 To cgrdPreview.Rows - 1
                Set lcolItem = New Collection
                lcolItem.Add cgrdPreview.TextMatrix(i, 0), "单位名称"
                lcolItem.Add cgrdPreview.TextMatrix(i, 1), "姓名"
                lcolItem.Add cgrdPreview.TextMatrix(i, 2), "性别"
                lcolItem.Add cgrdPreview.TextMatrix(i, 3), "年龄"
                
                lcolInfo.Add lcolItem, CStr(i)
            Next i
            mobj对外接口.Sub导入体检人员登记 clstTemplate.List(clstTemplate.ListIndex), cprgDataTransform, lcolInfo
            
            On Error Resume Next
            mobj对外接口.sub更改文件名 ctxtDataSource.Text
            
            cstbMain.Panels(1) = "导入成功。"
            cgrdPreview.Rows = 1
            ctbMain.Buttons(1).Enabled = False
            ctbMain.Buttons(2).Enabled = False
            
            cprgDataTransform.Visible = False
            MousePointer = 0
            Cancel = True
    End Select
    
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputFromExcel", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    cprgDataTransform.Visible = False
    cstbMain.Panels(1) = ""
    MousePointer = 0
End Sub


