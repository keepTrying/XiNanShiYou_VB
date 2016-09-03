VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputDataFromMdb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "本单位数据导入"
   ClientHeight    =   7455
   ClientLeft      =   1170
   ClientTop       =   1215
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10470
   Begin VB.Frame Frame2 
      Caption         =   "数据服务器类型"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   120
      TabIndex        =   23
      Top             =   1260
      Width           =   5235
      Begin VB.OptionButton coptInUnit 
         Caption         =   "站内数据服务器"
         Height          =   315
         Left            =   2880
         TabIndex        =   25
         Top             =   180
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton coptOutUnit 
         Caption         =   "站外数据服务器"
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   1995
      End
   End
   Begin VB.CommandButton ccmdPrefech 
      Caption         =   "预读取数据(&W)"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   840
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "预览数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3195
      Left            =   120
      TabIndex        =   20
      Top             =   3780
      Width           =   10335
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdPreviewData 
         Height          =   2835
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5001
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
   Begin MSComctlLib.ProgressBar cprgDatatranform 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame cfraFiltrateCondition 
      Caption         =   "数据导入条件"
      ForeColor       =   &H00800000&
      Height          =   1785
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   6645
      Begin MSComCtl2.DTPicker cdtpBeginDate 
         Height          =   375
         Left            =   1320
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
         Format          =   23527425
         CurrentDate     =   36951
      End
      Begin VB.CheckBox cchkSystemCode 
         Caption         =   "系统编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox cchkUnitName 
         Caption         =   "单位名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox cchkMedicalDate 
         Caption         =   "体检日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "单位定位(&L)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox ctxtUnit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox ctxtBeginCode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox ctxtEndCode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker cdtpEndDate 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
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
         Format          =   23527425
         CurrentDate     =   36951
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "到"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3120
         TabIndex        =   22
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "到"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3840
         TabIndex        =   17
         Top             =   1320
         Width           =   180
      End
   End
   Begin VB.Frame cfraSelectInputData 
      Caption         =   "选择要导入的内容"
      ForeColor       =   &H00800000&
      Height          =   1905
      Left            =   6840
      TabIndex        =   15
      Top             =   1800
      Width           =   3555
      Begin VB.ListBox clstDataType 
         Height          =   1530
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.TextBox ctxtDataSource 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   0
      Top             =   840
      Width           =   4155
   End
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "浏览文件(&B)"
      Height          =   375
      Left            =   5460
      TabIndex        =   1
      Top             =   840
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8340
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
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
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   13
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
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   9300
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label ccmdDataSource 
      AutoSize        =   -1  'True
      Caption         =   "数据来源："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmInputDataFromMdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjGUI  As cls界面通用对象 '用于初始化工具览。
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj对外接口 As ClsManageTransmission  '业务对象。

Private mstr系统编号固定部分  As String

Public pblnInUse As Boolean                    '表明窗体是否已加载。主程序要使用。

Private Sub coptInUnit_Click()
    On Error GoTo errHandler
    
    MousePointer = 11
    '清空数据分类。
    clstDataType.Clear
    
    '界面暂时不能操作。
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
            
    sub预读取数据
        
    '恢复界面。
    ctbMain.Buttons(1).Enabled = True
    cfrafiltrateCondition.Enabled = True
    cfraSelectInputData.Enabled = True
    
    MousePointer = 0
    Exit Sub
errHandler:
    MousePointer = 0
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "coptInUnit_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub coptOutUnit_Click()
    On Error GoTo errHandler
    
    MousePointer = 11
    '清空数据范围。
    clstDataType.Clear
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
            
    sub预读取数据
        
    ctbMain.Buttons(1).Enabled = True
    
    cfrafiltrateCondition.Enabled = True
    cfraSelectInputData.Enabled = True
    
    MousePointer = 0
    Exit Sub
errHandler:
    MousePointer = 0
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "coptOutUnit_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub



Private Sub ctxtBeginCode_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtEndCode.SetFocus
    End If
End Sub

Private Sub ctxtDataSource_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxtDataSource.Text <> "" Then
            ccmdPrefech.Enabled = True
            ccmdPrefech.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtDataSource.SetFocus
    ctxtDataSource.SelStart = Len(ctxtDataSource)
    ctxtDataSource.SelLength = 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub
'功能：初始化界面。
'作者：刘浩。
Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lcol工具栏按钮 As New Collection
    
    On Error GoTo errHandler
    pblnInUse = True
    
    '创建界面通用对象，初始化工具栏。
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
    
    '创建体检对外接口对象。
    Set mobj对外接口 = CreateObject("体检对外接口部件.ClsManageTransmission")
    
    '从工作站配置文件中获取记录的上次导入文件。
    ctxtDataSource.Text = mobj对外接口.工作站配置.内部导入文件
    
    '在各条件的check框未被选中时,各条件输入框不可用.
    ctxtBeginCode.Enabled = False
    ctxtEndCode.Enabled = False
    ctxtUnit.Enabled = False
    ccmdLocateUnit.Enabled = False
    cdtpBeginDate.Value = Format(Date, "yyyy-mm-dd")
    cdtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
    '导入,预览按钮在选定文件后才变为可用.
    If Len(ctxtDataSource.Text) = 0 Then
        ccmdPrefech.Enabled = False
    End If
    If ctxtDataSource.Text = "" Then
        ccmdPrefech.Enabled = False
    End If
    ctbMain.Buttons(2).Enabled = False
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
    ctbMain.Buttons(1).Enabled = False
    
    cdtpBeginDate.Value = Date
    cdtpEndDate.Value = Date
    
    '获取系统编号固定部分。
    Dim lobj体检 As Object '体检对象，获取系统编号的固定部分。
    Set lobj体检 = CreateObject("体检对象部件.clsMedicalExam")
    mstr系统编号固定部分 = lobj体检.系统编号固定部分
    Set lobj体检 = Nothing
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

'功能： 用户点击体检日期check框时，时间输入控件在可用和不可用的状态之间变化。
'作者： 刘浩。
Private Sub cchkMedicalDate_Click()
    On Error GoTo errHandler
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
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "cchkMedicalDate_Click", Err.Number, Err.Description, False
End Sub

'功能： 在用户点击系统编号check框时，调整系统编号录入框的可用状态。
'作者： 刘浩。
Private Sub cchkSystemCode_Click()
    On Error GoTo errHandler
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
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "cchkSystemCode_Click", Err.Number, Err.Description, False
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

'功能： 在用户击单位名称check框时，调整单位名称录入框的可用状态。
'作者： 刘浩。
Private Sub cchkUnitName_Click()
    On Error GoTo errHandler
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
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "cchkUnitName_Click", Err.Number, Err.Description, False
End Sub

'功能：返回通过单位定位得到的单位名称，显示在单位名称文本框中，各单位名称之单用英文逗号分隔。
'作者： 刘浩。
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobj体检管理 As Object
    Dim lobjRec As Object
    Dim lstrUnit As String
    
    Set lobj体检管理 = CreateObject("体检对象部件.clsManageMedicalExam")
    Set lobjRec = lobj体检管理.func单位定位
    
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            lstrUnit = lobjRec.Fields("单位名称").Value
        End If
    End If
    If Len(lstrUnit) >= 1 Then
        If Len(ctxtUnit.Text) >= 1 Then
            ctxtUnit.Text = ctxtUnit.Text & "," & lstrUnit
        Else
            ctxtUnit.Text = lstrUnit
        End If
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub

'功能：查找单位名称集字串中的中文逗号，并用英文逗号替换掉。
'作者： 刘浩。
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

'功能:  预读出MDB文件中的导入条件和导入内容,并显示在界面上。
'作者： 刘浩。
Private Sub ccmdPrefech_Click()
    On Error GoTo errHandler
    
    MousePointer = 11
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
    Frame2.Enabled = False
    
    '把指定的文件更名后copy到程序执行的路文件夹下。
    If Len(ctxtDataSource.Text) = 0 Then
        Err.Number = 6666
        Err.Description = "预读取数据前，请输入正确的导入文件及路径！"
        GoTo errHandler
    Else
        mobj对外接口.sub导入准备 ctxtDataSource.Text
    End If
        
    sub预读取数据
        
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(2).Enabled = True
    cfrafiltrateCondition.Enabled = True
    cfraSelectInputData.Enabled = True
    Frame2.Enabled = True
    MousePointer = 0
    Exit Sub
errHandler:
    MousePointer = 0
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "ccmdPrefech_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
Private Sub sub预读取数据()
    Dim lobjRec As Object
    Dim i As Integer
    
    On Error GoTo errHandler
    '初始化"数据导入条件"frame内的各项。
    cdtpBeginDate.Value = Date
    cdtpEndDate.Value = Date
    ctxtBeginCode = ""
    ctxtEndCode = ""
    ctxtUnit = ""
    Set lobjRec = mobj对外接口.Func获取mdb数据交换文件中的数据范围
    While Not lobjRec.EOF
        Select Case lobjRec.Fields("范围名").Value
            Case "开始日期"
                cdtpBeginDate.Year = DatePart("yyyy", lobjRec.Fields("范围值").Value)
                cdtpBeginDate.Month = DatePart("m", lobjRec.Fields("范围值").Value)
                cdtpBeginDate.Day = DatePart("d", lobjRec.Fields("范围值").Value)
            Case "结束日期"
                cdtpEndDate.Year = DatePart("yyyy", lobjRec.Fields("范围值").Value)
                cdtpEndDate.Month = DatePart("m", lobjRec.Fields("范围值").Value)
                cdtpEndDate.Day = DatePart("d", lobjRec.Fields("范围值").Value)
            Case "单位名称集"
                ctxtUnit.Text = lobjRec.Fields("范围值").Value
            Case "从系统编号"
                If Len(lobjRec.Fields("范围值").Value) <> 0 Then
                    ctxtBeginCode.Text = lobjRec.Fields("范围值").Value
                End If
            Case "到系统编号"
                If Len(lobjRec.Fields("范围值").Value) <> 0 Then
                    ctxtEndCode.Text = lobjRec.Fields("范围值").Value
                End If
        End Select
        lobjRec.MoveNext
    Wend
        
    '把MDB数据库中包含的数据分类名（导出数据时记录在表传输数据类型表中）列在数据分类列表框中。
    Set lobjRec = mobj对外接口.Func获取mdb文件中的数据分类清单(IIf(coptInUnit, "站内数据服务器", "站外数据服务器"))
    clstDataType.Clear
    While Not lobjRec.EOF
        clstDataType.AddItem lobjRec.Fields("数据分类名").Value
        clstDataType.Selected(clstDataType.NewIndex) = True
        lobjRec.MoveNext
    Wend
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "sub预读取数据", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

'功能： 弹出文件查找窗口。
'作者： 刘浩。
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
    ctxtDataSource.Text = ccdgBrowse.FileName
    
    If ctxtDataSource.Text = "" Then
        ccmdPrefech.Enabled = False
    Else
        ccmdPrefech.Enabled = True
    End If
    Exit Sub
errHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set mobj对外接口 = Nothing
    Set mobjGUI = Nothing
    pblnInUse = False
    
End Sub

'功能： 导入,预览报盘文件中的数据.
'作者： 刘浩。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lcolRange As Collection '存放导入或预览数据的过滤条件。
    Dim lobjRec As Object '用于浏览的数据
    Dim lcolType As Collection
    Dim i As Integer
    
    Set lcolRange = funcCalCon
    Select Case Operate
        Case "预览"
            '将取出的数据显示在预览数据框内
            cgrdPreviewData.Rows = 1
            Set lobjRec = mobj对外接口.Func查看数据(lcolRange, 0) '(0导入/ 1导出)
            gfsubLoadGridFromRec cgrdPreviewData, lobjRec, , "健康档案编号,系统编号,公民身份号码,姓名,性别,出生日期,单位名称,体检日期,体检表名称,体检结论,诊断和处理意见,体检医师"
            cgrdPreviewData.Rows = cgrdPreviewData.Rows - 1
            
            cprgDatatranform.Value = 0
            Cancel = True
            
        Case "导入"
            cprgDatatranform.Value = 0
            cprgDatatranform.Visible = True
            MousePointer = 11
            csbMain.Panels(1) = "正在导入，请稍候..."
            '取出需要导入的数据分类名。
            Set lcolType = New Collection
            For i = 0 To clstDataType.ListCount - 1
                If clstDataType.Selected(i) Then
                    lcolType.Add clstDataType.List(i), clstDataType.List(i)
                End If
            Next i
            '准备导入（拷贝文件成临时文件）。
            'mobj对外接口.sub导入准备 ctxtDataSource.Text
            
            '开始导入。
            mobj对外接口.Sub数据导入 lcolRange, lcolType, cprgDatatranform
            
            '导入成功。
            csbMain.Panels(1) = "导入成功。"
            MousePointer = 0
            cprgDatatranform.Visible = False
            Cancel = True
    End Select
    
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    cprgDatatranform.Visible = False
    csbMain.Panels(1) = ""
    MousePointer = 0
    Exit Sub
    Resume
End Sub

'功能：  用户输入了MDB文件名后,"预读取数据" 按钮变为可用
'作者： 刘浩。
Private Sub ctxtDataSource_LostFocus()
    Dim lstrErrDes As String
    
    On Error GoTo errHandler
    
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(2).Enabled = False
    
    If Len(ctxtDataSource.Text) <> 0 Then
        If UCase(Right(ctxtDataSource.Text, 3)) = "MDB" And Dir(ctxtDataSource.Text) <> "" Then
            ccmdPrefech.Enabled = True
        Else
            ccmdPrefech.Enabled = False
            Err.Raise 6666, , "输入的文件名不合法，请重新输入！"
        End If
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "ctxtDataSource_LostFocus", Err.Number, Err.Description, False
End Sub

'功能： 通过"数据导入条件"frame框中的设置，在数据导入或预览前获取数据过滤条件。
'作者： 刘浩。
Private Function funcCalCon() As Object
    Dim lobjRange As Collection '[数据范围名，数据范围值] "数据导入条件"
    Dim lcolItem As Collection
    On Error GoTo errHandler
    
    Set lobjRange = New Collection
    If cchkMedicalDate.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add cdtpBeginDate.Value, "数据范围值"
        lcolItem.Add "开始日期", "数据范围名"
        lobjRange.Add lcolItem, "开始日期"
                
        Set lcolItem = New Collection
        lcolItem.Add cdtpEndDate.Value, "数据范围值"
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
            
    If cchkUnitName.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add ctxtUnit.Text, "数据范围值"
        lcolItem.Add "单位名称集", "数据范围名"
        lobjRange.Add lcolItem, "单位名称集"
    End If
    
    Set funcCalCon = lobjRange
    Exit Function
errHandler:
    sfsub错误处理 "体检对外接口部件", "frmInputDataFromMdb", "funcCalCon", Err.Number, Err.Description, True
End Function












