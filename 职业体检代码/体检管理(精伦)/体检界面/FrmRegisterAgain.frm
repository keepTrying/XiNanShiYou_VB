VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Begin VB.Form FrmRegisterAgain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "复查登记"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10335
   ClipControls    =   0   'False
   Icon            =   "FrmRegisterAgain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   2400
      Top             =   600
   End
   Begin VB.TextBox ctxtTemplate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   2175
   End
   Begin VB.PictureBox cpicPhoto 
      Height          =   1785
      Left            =   360
      ScaleHeight     =   1725
      ScaleWidth      =   1365
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1425
   End
   Begin MSComctlLib.Toolbar ctblTool 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchkPrint 
         Caption         =   "打印体检单"
         Height          =   345
         Left            =   6720
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CheckBox cchkClear 
         Caption         =   "保存后清空"
         Height          =   345
         Left            =   4920
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1425
      End
   End
   Begin VB.Frame cframBase 
      Caption         =   "登记基本信息"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   10095
      Begin VB.TextBox ctxtTubeNo 
         Height          =   315
         Left            =   5280
         TabIndex        =   26
         Top             =   480
         Width           =   2415
      End
      Begin VB.ListBox clstItem 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   3300
         ItemData        =   "FrmRegisterAgain.frx":0442
         Left            =   120
         List            =   "FrmRegisterAgain.frx":0444
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2190
      End
      Begin MSComCtl2.DTPicker cdtpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   0
         Top             =   1080
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
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
         Format          =   20119552
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.VScrollBar cvscLetter 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5775
         TabIndex        =   2
         Top             =   450
         Width           =   435
      End
      Begin 录入控件.ctlInputFrame ciptBase 
         Height          =   3780
         Left            =   2520
         TabIndex        =   14
         Top             =   2280
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   6668
         BackColor       =   15791081
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
         BackStyle       =   1
         BorderStyle     =   0
         Caption         =   "Frame1"
         Rows            =   5
         Cols            =   6
         DistanceofRow   =   0
         BorderStyle     =   0
         FormatString    =   "身份证号,1,0,3"
         Count           =   1
         titleInputBox0001=   "身份证号"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   3
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
         名称InputBox0001=   "身份证号"
         缺省值InputBox0001=   ""
         保存缺省值InputBox0001=   ""
         长度InputBox0001=   0
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         EnableInputBox0001=   0   'False
         允许多选InputBox0001=   0   'False
         ErrColor        =   15791081
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复查项目："
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label clblUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         TabIndex        =   22
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label clblAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3720
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label clblSex 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label clblName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         TabIndex        =   19
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   8
         Left            =   5280
         TabIndex        =   17
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   9
         Left            =   3720
         TabIndex        =   18
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Index           =   7
         Left            =   5280
         TabIndex        =   16
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Index           =   6
         Left            =   2520
         TabIndex        =   15
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复查日期："
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管编号："
         Height          =   180
         Index           =   1
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   900
      End
      Begin VB.Label clblSysNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         TabIndex        =   5
         Top             =   450
         Width           =   495
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "保存后请看状态栏"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6210
         TabIndex        =   4
         Top             =   450
         Width           =   1545
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7860
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15610
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   1680
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRegisterAgain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春

Private WithEvents mobjGUI As cls界面通用对象    '界面通用对象，用于初始化工具栏，录入板控件控制。
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj体检 As Object                       '体检对象

Public pstr旧系统编号 As String                 '记录原体检记录的系统编号。
Private mstr系统编号 As String                   '最近保存过的复查体检记录的系统编号。
Private mcolTubeNo As New Collection             '当前复查体检表可选择的试管字母。

Private mstr上个体检表名 As String
Private mstr系统编号固定部分 As String

'业务设置。
Private mbln快速录入 As Boolean

Private mblnInUse As Boolean                     '对应属性pblnInUse。

Private mcol收费项目 As Collection               'item:编号,key：编号。

'功能：记载当前窗体是否已加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
Public Property Get pblnInUse() As Boolean
    On Error GoTo errHandler
    pblnInUse = mblnInUse
    Exit Property
errHandler:
    'sfsub错误处理 "体检界面部件", "FrmRegisterAgain", "Property Get pblnInUse", Err.Number, Err.Description, True
End Property

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

'功能：窗体初始化：初始化工具栏（余下的初始化工作在定时器Timer1_timer中完成）。
Private Sub Form_Load()
    On Error GoTo errHandler

    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
        
    '显示等待状态。
    MousePointer = 11
    csbMain.Panels(1) = "窗体正在初始化，请稍侯..."
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    mstr上个体检表名 = ""
    
    '初始化界面通用对象。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    Dim lcol工具栏按钮 As New Collection           '工具栏上的按钮初始化集合。
    With lcol工具栏按钮
        .Add "选择项目(&I)111"
        .Add "保存"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctblTool
        Set .c状态栏 = csbMain
        Set .c录入板 = ciptBase
        
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcol工具栏按钮, ""
    End With
    
    '初始时，保存按钮不可用。
    ctblTool.Buttons(1).Enabled = False
    ctblTool.Buttons(2).Enabled = False
    
    cdtpDate.Value = Format(Date, "yyyy-mm-dd")
    
    If pobj业务对象.业务设置("试管编号自动生成") = "否" Then
        ctxtTubeNo.Visible = True
        ctxtTubeNo.TabIndex = 1
        clblLetter.Visible = False
        cvscLetter.Visible = False
    Else
        ctxtTubeNo.Visible = False
        clblLetter.Visible = True
        cvscLetter.Visible = True
    End If
        
    '为了加快窗体加载速度，余下初始化工作放在定时器中完成。
    Timer1.Enabled = True
    Exit Sub
    
errHandler:
    '错误处理。
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAgain", "Form_Load", 6666, lstrError, False
    
End Sub

'功能：完成form_load余下的窗体初始化工作。
Private Sub Timer1_Timer()
    Dim lcolInfo As Collection   '当日工作记忆簿中保存的单位名称集。
    Dim lobjRec As Object        '由业务对象获取的待复查名单。
    Dim i As Integer
       
    On Error GoTo errHandler
    
    '定时器不再起作用。
    Timer1.Enabled = False
    
    '界面暂时不可操作。
    cframBase.Enabled = False
    ciptBase.Enabled = False
    ctblTool.Enabled = False
    
    '创建体检对象。
    Set mobj体检 = CreateObject("体检对象.clsMedicalExam")
    mstr系统编号固定部分 = mobj体检.系统编号固定部分
    
    '判断业务设置是否打印体检单。
    If pobj业务对象.业务设置("是否打印体检单") = "是" Then
        cchkPrint.Visible = True
    End If
    
    '显示复查人员体检信息。
    SubGetPersonInfo
    
        
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmRegisterAgain", "Timer1_Timer", 6666, lstrError, False
        '界面不可操作。
        cframBase.Enabled = False
    End If
    
    '恢复界面可操作。
    cframBase.Enabled = True
    ciptBase.Enabled = True
    ctblTool.Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub

'清空
Private Sub subClear()
    On Error Resume Next
    clblSysNo.Caption = ""
    clblLetter.Caption = ""
    ctxtTubeNo.Text = ""
    cdtpDate.Value = Date
    ciptBase.ClearContent
    clstItem.Clear
    clblName.Caption = ""
    clblSex.Caption = ""
    clblAge.Caption = ""
    clblUnit.Caption = ""
    cframBase.Caption = "登记基本信息"
    '选择项目、保存按钮不可用。
    ctblTool.Buttons(1).Enabled = False
    ctblTool.Buttons(2).Enabled = False

    Set mcol收费项目 = New Collection
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '若新增体检记录没有保存，退回系统编号。
    If Not mobj体检 Is Nothing Then
        If mobj体检.系统编号 <> "" And Not mobj体检.是否已存在 Then
            '退回系统编号。
            mobj体检.sub退回复查系统编号 mobj体检.系统编号
        End If
    End If
    
    '释放对象。
    Set mobj体检 = Nothing
    
    '设置窗体没有启动的标志。
    mblnInUse = False

End Sub


Private Sub cvscLetter_Change()
    On Error Resume Next
    '点击滚动条，获得相应的字母。
    If mcolTubeNo.Count > 0 Then
        clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
    End If
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    Dim lstr流水号 As String
    
    On Error GoTo errHandler
    
    Select Case Operate
    Case "选择项目"
        Dim lcol复查项目 As Collection
        Dim lobj体检模板 As Object
        
        '获取体检表上已有的体检项目。
        Set lcol复查项目 = mobj体检.体检表.体检项目集("")
        
        '设置选择项目界面的属性。
        frmSelectItem.pstr体检表名称 = ctxtTemplate.Text
        Set frmSelectItem.pcol复查项目 = lcol复查项目
        Set frmSelectItem.pcol收费项目 = mcol收费项目
        '启动选择项目界面。
        frmSelectItem.Show 1
        If frmSelectItem.pblnOk Then
            '获取选中的复查项目。
            Set lcol复查项目 = frmSelectItem.pcol复查项目
            '修改体检表，并重新显示在列表中。
            clstItem.Clear
            mobj体检.体检表.Sub删除所有体检项目
            For i = 1 To lcol复查项目.Count
                mobj体检.体检表.Sub添加体检项目 lcol复查项目(i)("编码")
                clstItem.AddItem lcol复查项目(i)("编码") & " " & lcol复查项目(i)("名称")
            Next
            
            '获取设置的收费项目。
            Set mcol收费项目 = frmSelectItem.pcol收费项目
            
            '修改：2002-10-14（杨春）嘉定定制：显示收费金额。
            Dim ldblTotal As Double
            For i = 1 To mcol收费项目.Count
                ldblTotal = Format(ldblTotal + mcol收费项目(i)("单价"), "0.00")
            Next
            On Error Resume Next
            If sffunc判断集合键值是否存在(mobj体检.体检表.附加信息, "体检金额") Then
                ciptBase.Box1("体检金额").Text = ldblTotal
                mobj体检.体检表.Sub填附加信息值 "体检金额", ldblTotal
            End If
            
        End If
    Case "修改"
        FrmEditRegister.系统编号 = mstr系统编号
        FrmEditRegister.Move Me.Left, Me.Top
        FrmEditRegister.Show 1
        Cancel = True
    Case "保存"
        MousePointer = 11
        csbMain.Panels(1) = "正在保存体检登记信息，请稍侯..."
        With mobj体检
            If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
                If .体检表.试管编号字母 <> clblLetter.Caption Then
                    .体检表.试管编号字母 = clblLetter.Caption
                End If
            Else
                .试管编号 = ctxtTubeNo.Text
            End If
            .体检日期 = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            '设置体检类别为复查体检登记。
            .体检类别 = P_EXAM_AGAIN
        End With
        
        '保存复查登记信息。
        pobj业务对象.Sub体检登记 mobj体检, pstr旧系统编号, IIf(cchkPrint.Value = 1, True, False), mcol收费项目
                
        '记录当前保存过的系统编号。
        mstr系统编号 = clblSysNo.Caption
        
        '在状态栏上显示保存后的系统编号、试管编号。
        csbMain.Panels(2) = "上次保存的体检系统编号：" & mstr系统编号 & "，试管编号：" & mobj体检.试管编号 & "。"
        
        If cchkClear = 1 Then
            subClear
        End If
        
        '试管字母不能再选择。
        cvscLetter.Enabled = False
        
        csbMain.Panels(1) = ""
        MousePointer = 0
        Cancel = True
    Case "退出"
'        '若新增体检记录没有保存，退回系统编号。
'        If mobj体检.系统编号 <> "" And Not mobj体检.是否已存在 Then
'            '退回系统编号。
'            mobj体检.sub退回复查系统编号 mobj体检.系统编号
'        End If
        
    End Select
    Exit Sub
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAgain", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    csbMain.Panels(1) = ""
    MousePointer = 0
    Cancel = True
    Exit Sub
    Resume
End Sub

'功能：显示指定系统编号的体检人员的信息在界面上。
Private Sub SubGetPersonInfo()
    Dim lobj体检 As Object     'clsMedicalExam旧体检记录。
    Dim lobj体检模板 As Object 'clsMedicalExamTemplate
    Dim lcolInfo As New Collection
    Dim i As Integer
    Dim j As Integer
    
    Dim lstrTemp As String
    Dim lstrTubeNo As String
    Dim lstrSysNo As String
    
    On Error GoTo errHandler
    MousePointer = 11
    csbMain.Panels(1) = "正在显示当前复查人员的信息，请稍侯..."
    
    '界面暂时不可操作。
    ctblTool.Enabled = False
    
    '先退回旧系统编号。
    If Not mobj体检.是否已存在 And mobj体检.系统编号 <> "" Then
        mobj体检.sub退回复查系统编号 mobj体检.系统编号
    End If
        
    '创建旧体检对象。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    lobj体检.系统编号 = pstr旧系统编号
    
    ctxtTemplate.Text = lobj体检.复查体检表名称
    
    '按旧体检记录的体检表名初始化录入板。
    If mstr上个体检表名 <> lobj体检.体检表.体检表名 Then
        mstr上个体检表名 = lobj体检.体检表.体检表名
        '重新初始化录入板。
        On Error Resume Next
        mobjGUI.sub初始化录入板 mstr上个体检表名
        On Error GoTo errHandler
    End If
    
    '创建体检表模板对象。
    Set lobj体检模板 = CreateObject("体检对象.clsMedicalExamTemplate")
    lobj体检模板.体检表名 = ctxtTemplate.Text
    
    '获取复查体检表的体检项目。
    Set lcolInfo = lobj体检模板.体检项目集
    
    '显示复查项目。
    clstItem.Clear
    For i = 1 To lcolInfo.Count
        clstItem.AddItem lcolInfo(i).编码 & " " & lcolInfo(i).名称
    Next i
    
    '获取旧体检记录的附加信息。
    Set lcolInfo = lobj体检.体检表.附加信息
    
    '填写附加信息结果。
    sub填录入板值 ciptBase, mobjGUI, lcolInfo
    
    DoEvents
    
    '显示基本信息。
    With lobj体检.体检人员
        clblName.Caption = .姓名
        clblSex.Caption = .性别
        clblAge.Caption = .年龄
        clblUnit.Caption = .单位名称
        '相片
        cpicPhoto.Picture = .像片
    End With
    cframBase.Caption = "登记基本信息（" & clblName.Caption & "）"
    '分配新的系统编号
    lstrSysNo = mobj体检.Func分配复查系统编号(pstr旧系统编号)
    mobj体检.系统编号 = lstrSysNo
    clblSysNo.Caption = lstrSysNo
    
    '健康档案不变。
    mobj体检.体检人员.健康档案编号 = lobj体检.体检人员.健康档案编号
    
    
    '设置复查体检的体检表明，从而获取新试管编号。
    mobj体检.体检表.体检表名 = ctxtTemplate.Text
    
    If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
        '试管编号字母为空时cvscLetter可用
        If mobj体检.体检表.试管编号字母 = "" Then
            '将字母按逗号分开，加入mcoltubeNo中
            lstrTubeNo = lobj体检模板.试管字母编号
            If Right(lstrTubeNo, 1) <> "," Then lstrTubeNo = lstrTubeNo & ","
            lstrTemp = ""
            Set mcolTubeNo = New Collection
            For i = 1 To Len(lstrTubeNo)
                lstrTemp = lstrTemp & Mid(lstrTubeNo, i, 1)
                If Mid(lstrTubeNo, i, 1) = "," Then
                    If Left(lstrTemp, Len(lstrTemp) - 1) <> "" Then
                        mcolTubeNo.Add Left(lstrTemp, Len(lstrTemp) - 1)
                    End If
                    lstrTemp = ""
                End If
            Next i
            If mcolTubeNo.Count > 0 Then
                '试管字母改变了，给出提示。
                If clblLetter.Caption <> "" And clblLetter.Caption <> mcolTubeNo(1) Then
                    sffuncMsg "请注意，你现在选择的体检表使用的试管字母与前一个（" & clblLetter.Caption & "）不同了。"
                End If
            
                '赋值给clblLetter
                clblLetter.Caption = mcolTubeNo(1)
                cvscLetter.Enabled = True
                cvscLetter.Min = 1
                cvscLetter.Max = mcolTubeNo.Count
                cvscLetter.Value = 1
            Else
                ctblTool.Buttons(1).Enabled = False
                ctblTool.Buttons(2).Enabled = False
                '提示该体检表无可用的字母。
                Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
            End If
        Else
            '有字母，不能选择字母。
            clblLetter.Caption = mobj体检.体检表.试管编号字母
            cvscLetter.Enabled = False
        End If
    Else
        ctxtTubeNo = mobj体检.试管编号
    End If
    
    '保存附加信息
    For i = 1 To ciptBase.ItemCount
        If ciptBase.InfoCollection(i).字典名称 <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
            mobj体检.体检表.Sub填附加信息值 ciptBase.InfoCollection(i).名称, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
        Else
            mobj体检.体检表.Sub填附加信息值 ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
        End If
        
    Next i
            
    '保存按钮可用。
    ctblTool.Buttons(1).Enabled = True
    ctblTool.Buttons(2).Enabled = True
    
    '所有附加项目不可修改。
    For i = 0 To ciptBase.ItemCount - 1
        ciptBase.ItemEnable(i) = False
    Next
    
    '修改：2002-10-10（杨春）嘉定定制：显示体检金额。
    On Error Resume Next
    If sffunc判断集合键值是否存在(mobj体检.体检表.附加信息, "体检金额") Then
        ciptBase.Box1("体检金额").Text = lobj体检模板.收费标准金额
        mobj体检.体检表.Sub填附加信息值 "体检金额", lobj体检模板.收费标准金额
    End If
    
    Set mcol收费项目 = New Collection
    clstItem.Refresh
'    DoEvents
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        sfsub错误处理 "体检界面部件", "FrmRegisterAgain", "SubGetPersonInfo", Err.Number, Err.Description, False
    End If
    Set lobj体检模板 = Nothing
    Set lobj体检 = Nothing
    
    '恢复界面可操作。
    ctblTool.Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
