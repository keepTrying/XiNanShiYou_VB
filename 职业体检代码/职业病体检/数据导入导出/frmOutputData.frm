VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutputData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "体检人员名单导出"
   ClientHeight    =   8355
   ClientLeft      =   1395
   ClientTop       =   1215
   ClientWidth     =   11925
   ClipControls    =   0   'False
   Icon            =   "frmOutputData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11925
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ccmdExit 
      Caption         =   "返回(&X)"
      Height          =   375
      Left            =   10440
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton ccmdExport 
      Caption         =   "导出(&P)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10440
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton ccmdView 
      Caption         =   "预览(&V)"
      Height          =   375
      Left            =   10440
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   8040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "预览数据"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   5895
      Left            =   120
      TabIndex        =   18
      Top             =   1965
      Width           =   11655
      Begin VB.CheckBox cchkAll 
         Caption         =   "全选"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdPreviewData 
         Height          =   5115
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   11415
         _cx             =   20135
         _cy             =   9022
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
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
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
         MultiTotals     =   0   'False
         SubtotalPosition=   1
         OutlineBar      =   1
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   -1  'True
         ShowComboButton =   0   'False
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   2160
            TabIndex        =   22
            Top             =   -480
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "浏览文件(&B)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox ctxtOutputDestination 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Frame cfrafiltrateCondition 
      Caption         =   "数据导出条件"
      ForeColor       =   &H00800000&
      Height          =   1185
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   10245
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox cchkTemplate 
         Caption         =   "体检对象"
         Height          =   285
         Left            =   6240
         TabIndex        =   10
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox ctxtEndCode 
         Height          =   315
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   2115
      End
      Begin VB.TextBox ctxtBeginCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox ctxtUnit 
         Height          =   315
         Left            =   7320
         TabIndex        =   6
         Top             =   240
         Width           =   2835
      End
      Begin VB.CheckBox cchkSystemCode 
         Caption         =   "体检单号"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   1035
      End
      Begin VB.CheckBox cchkUnitName 
         Caption         =   "单位名称"
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   240
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
         Left            =   1320
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
         Format          =   25296896
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
         Left            =   3960
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
         Format          =   25296896
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
         Left            =   3480
         TabIndex        =   19
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
         Height          =   300
         Left            =   3600
         TabIndex        =   15
         Top             =   840
         Width           =   180
      End
   End
   Begin MSComctlLib.ProgressBar cprgDatatranform 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   7920
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   6960
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
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
      Height          =   300
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Width           =   900
   End
End
Attribute VB_Name = "frmOutputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj对外接口 As ClsManageTransmission '体检对外接口对象

Private mstrID As String
Private mcolIndex As New Collection

Private Sub cchkAll_Click()
    Dim i As Long
    
        For i = 1 To cgrdPreviewData.Rows - 1
            cgrdPreviewData.Cell(flexcpChecked, i, 0, i, 0) = IIf(cchkAll.Value = 1, flexChecked, flexUnchecked)
        Next
    
End Sub

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdExport_Click()
    On Error GoTo errHandler
    Dim i As Long
    
    If ctxtOutputDestination = "" Then
        MsgBox "必须输入导出文件！", vbOKOnly + vbExclamation, "系统提示"
        ctxtOutputDestination.SetFocus
        Exit Sub
    End If
    '先把所有未选中数据删除。
    For i = 1 To cgrdPreviewData.Rows - 1
        If cgrdPreviewData.Cell(flexcpChecked, i, 0) = flexUnchecked Then
            dafuncGetData "delete temp_体检基本信息  where ID='" & mstrID & "' and 系统编号='" & cgrdPreviewData.TextMatrix(i, mcolIndex("系统编号")) & "'"
        End If
    Next
    
    mobj对外接口.sub导出体检人员名单 ctxtOutputDestination, mstrID
    MsgBox "导出成功！", vbOKOnly + vbExclamation, "系统提示"
    Exit Sub
errHandler:
    sfsub错误处理 "体检数据导入导出", "frmOutputData", "ccmdExport_Click", Err.Number, Err.Description, False
End Sub

Private Sub ccmdView_Click()
    Dim lobjRec As Object
    Dim i As Long
    
    '获取要预览的数据范围。
    ctxtOutputDestination.SetFocus
    ccmdExport.Enabled = False
    cgrdPreviewData.Rows = 1
    
    If mstrID <> "" Then
        dafuncGetData "delete temp_体检基本信息 where ID='" & mstrID & "'"
    End If
    
    '将取出的数据显示在预览数据框内
    Set lobjRec = mobj对外接口.func获取体检人员名单(IIf(cchkMedicalDate.Value = 1, Format(cdtpBeginDate.Value, "yyyy-mm-dd"), ""), IIf(cchkMedicalDate.Value = 1, Format(cdtpEndDate.Value, "yyyy-mm-dd"), ""), _
                                            IIf(cchkUnitName.Value = 1, ctxtUnit.Text, ""), IIf(cchkTemplate.Value = 1, ccmbTemplate.Text, ""), _
                                            IIf(cchkSystemCode.Value = 1, ctxtBeginCode.Text, ""), IIf(cchkSystemCode.Value = 1, ctxtEndCode.Text, ""))
    Set cgrdPreviewData.DataSource = lobjRec
    cgrdPreviewData.AutoSize 0, cgrdPreviewData.Cols - 1
    If lobjRec.recordcount > 0 Then
        lobjRec.movefirst
        mstrID = lobjRec("ID")
        ccmdExport.Enabled = True
        
        For i = 1 To cgrdPreviewData.Rows - 1
            cgrdPreviewData.Cell(flexcpChecked, i, 0, i, 0) = IIf(cchkAll.Value = 1, flexChecked, flexUnchecked)
        Next
    End If
    cgrdPreviewData.ColHidden(cgrdPreviewData.Cols - 1) = True
    
    Set mcolIndex = New Collection
    For i = 0 To cgrdPreviewData.Cols - 1
        mcolIndex.Add i, cgrdPreviewData.TextMatrix(0, i)
    Next
End Sub

Private Sub cgrdPreviewData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
    
End Sub

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
    
    '创建体检对外接口对象。
    Set mobj对外接口 = New ClsManageTransmission
    
    
    '获取所有体检表名称。
    Dim lobj体检表模板集 As Object
    Set lobj体检表模板集 = CreateObject("体检对象.clsMedicalExamTemplateSet")
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

        
    '设置日期的初始值。
    cdtpBeginDate.Value = Format(DateAdd("d", -7, Date), "yyyy-mm-dd")
    cdtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检数据导入导出", "frmOutputData", "form_load", Err.Number, Err.Description, False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        SendKeys vbKeyTab
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
    sfsub错误处理 "体检数据导入导出", "frmOutputData", "cchkMedicalDate_Click", Err.Number, Err.Description, False
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
    sfsub错误处理 "体检数据导入导出", "frmOutputData", "cchkSystemCode_Click", Err.Number, Err.Description, False
End Sub

'功能： 点击单位名称check框时，设置单位名称输入控件的状态。
'作者： 刘浩
Private Sub cchkUnitName_Click()
    On Err GoTo errHandler
    If cchkUnitName.Value = 1 Then
        ctxtUnit.Enabled = True
        ctxtUnit.SetFocus
    Else
        ctxtUnit.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检数据导入导出", "frmOutputData", "cchkUnitName_Click", Err.Number, Err.Description, False
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
    sfsub错误处理 "体检数据导入导出", "frmOutputData", "ccmdTemplate_click", Err.Number, Err.Description, False
End Sub

'功能: 弹出文件查找窗口.
Private Sub ccmdBrowse_Click() ' 设置“CancelError”为 True
    On Error GoTo errHandler
    ccdgBrowse.CancelError = True
    ' 设置标志
    ccdgBrowse.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    ccdgBrowse.Filter = "All Files (*.*)|*.*|文本文件" & _
        "(*.txt)|*.txt"
    ccdgBrowse.FilterIndex = 2
    ccdgBrowse.ShowOpen
    
    ctxtOutputDestination.Text = ccdgBrowse.FileName
    
    Exit Sub
errHandler:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If mstrID <> "" Then
        dafuncGetData "delete temp_体检基本信息 where ID='" & mstrID & "'"
    End If
    
    Set mobj对外接口 = Nothing
End Sub

