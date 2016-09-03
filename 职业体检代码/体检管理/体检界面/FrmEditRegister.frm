VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmEditRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "修改体检登记信息"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10335
   ClipControls    =   0   'False
   Icon            =   "FrmEditRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin 录入控件.ctlInputDictGrid c字典表 
      Height          =   3135
      Left            =   2640
      TabIndex        =   27
      Top             =   4200
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   2280
      Top             =   600
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "保存后清空"
      Height          =   270
      Left            =   8040
      TabIndex        =   12
      Top             =   240
      Width           =   1410
   End
   Begin VB.Frame frmPhoto 
      Caption         =   "照像："
      ForeColor       =   &H00800000&
      Height          =   4305
      Left            =   5400
      TabIndex        =   25
      Top             =   2760
      Width           =   4695
      Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
         Height          =   3570
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4485
         _ExtentX        =   8017
         _ExtentY        =   6297
         BackColor       =   0
         FontSize        =   9.75
         OriginalSize    =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7470
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18177
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin VB.Frame cframJBXX 
      Caption         =   "登记基本信息"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1725
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   10215
      Begin VB.TextBox ctxtTubeNo 
         Height          =   315
         Left            =   5640
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox ctxt体检单号 
         Height          =   315
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox ccmb体检类型 
         Height          =   300
         ItemData        =   "FrmEditRegister.frx":0442
         Left            =   8640
         List            =   "FrmEditRegister.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox ccmbTemplate 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         ItemData        =   "FrmEditRegister.frx":045C
         Left            =   3360
         List            =   "FrmEditRegister.frx":045E
         TabIndex        =   7
         Top             =   1200
         Width           =   3525
      End
      Begin VB.TextBox ctxtSysNo 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2265
      End
      Begin VB.TextBox ctxtName 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1200
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "FrmEditRegister.frx":0460
         Left            =   1560
         List            =   "FrmEditRegister.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   960
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "定位(&T)"
         Height          =   375
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox ctxtAge 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1200
         Width           =   495
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
         Height          =   315
         Left            =   8400
         TabIndex        =   30
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
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
         Format          =   132775936
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检单号："
         Height          =   180
         Index           =   7
         Left            =   6840
         TabIndex        =   29
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检类型："
         Height          =   180
         Index           =   8
         Left            =   8640
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表："
         Height          =   180
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Index           =   4
         Left            =   1560
         TabIndex        =   23
         Top             =   960
         Width           =   540
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6120
         TabIndex        =   22
         Top             =   480
         Width           =   675
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5880
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管编号："
         Height          =   180
         Index           =   1
         Left            =   5640
         TabIndex        =   19
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
         Height          =   180
         Index           =   2
         Left            =   8400
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   5
         Left            =   3360
         TabIndex        =   17
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   2640
         TabIndex        =   16
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin 录入控件.ctlInputFrame ciptBase 
      Height          =   4455
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   7858
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
      Caption         =   "Frame1"
      Rows            =   6
      Cols            =   27
      DistanceofRow   =   0
      AutoSize        =   0   'False
      FormatString    =   "身份证号,1,0,12"
      Count           =   1
      titleInputBox0001=   "身份证号"
      statusinfoInputBox0001=   ""
      lengthInputBox0001=   12
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
      名称InputBox0001=   "输入框 1"
      缺省值InputBox0001=   ""
      保存缺省值InputBox0001=   ""
      长度InputBox0001=   0
      MaxInputBox0001 =   ""
      MinInputBox0001 =   ""
      VisibleInputBox0001=   -1  'True
      PermitNullInputBox0001=   -1  'True
      TriggerstrInputBox0001=   ""
      允许多选InputBox0001=   0   'False
      ErrColor        =   12648447
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   2760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmEditRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：邓恒
'最后修改：杨春

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

Private mobj体检 As Object                   '体检对象

Private mblnTakePhoto  As Boolean            '业务设置“是否照相”。
Private mbln快速录入 As Boolean

Private mstr系统编号 As String               '初始要修改的体检记录的系统编号（窗体启动参数）。

Private mstr系统编号固定部分 As String

Private mstr单位申请编号 As String
Private mblnInUse As Boolean                 '对应属性pblnInUse。

Public pstr系统编号名称 As String '修改：2002-10-10（杨春）为嘉定定制增加该属性。

Private mcol收费项目 As Collection
Private mcol体检项目 As Collection

'功能：获取本窗体是否已启动的标志。
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

'功能：设置窗体启动参数：当前要修改的体检记录的系统编号。
Public Property Let 系统编号(ByVal vNewValue As String)
    mstr系统编号 = vNewValue
    
End Property


Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If

End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtAge.SetFocus
    End If
End Sub

Private Sub ccmbTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxtTubeNo.Visible Then
            ctxtTubeNo.SetFocus
        Else
            ciptBase.SetFocus
        End If
    End If
End Sub

Private Sub ccmbUnit_Click()
    Dim i As Integer
    
    On Error GoTo errHandler
    If ccmbUnit.Text = "" Then Exit Sub
    
    '判断录入的单位是否在列表中存在，不存在则加入列表
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '加入到列表框中
        ccmbUnit.AddItem ccmbUnit.Text
        
        '加载到工作记忆簿文件中
        pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
    Else
        '修改：2001-8-23（显示单位属性）。
        On Error Resume Next
        mstr单位申请编号 = pobj业务对象.当日工作记忆簿.单位编号(ccmbUnit.Text)
        sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
    End If

    Exit Sub

errHandler:
End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '若录入板没有录入项目，则直接保存。
        If mobj体检.体检表.附加信息.Count > 0 Then
            ciptBase.SetFocus
        Else
            ciptBase_LastLostFocus
        End If
    Else
        mstr单位申请编号 = ""
    End If
End Sub

Private Sub ccmb体检类型_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '自动保存。
    If ctbMain.Buttons(1).Enabled Then
'        ctxtSysNo.SetFocus
'        SendKeys "{F2}"
        mobjGUI_BeforeOperate "保存", blnCancel
    End If
End Sub


Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub
Private Sub ctxtAge_GotFocus()
    On Error Resume Next
    With ctxtAge
        .SelStart = 0
        .SelLength = Len(Trim(ctxtAge.Text))
    End With
End Sub
Private Sub ctxtAge_KeyPress(KeyAscii As Integer)
'    On Error GoTo errhandler
'    '判断是否为数字
'    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = 13 Then
'    Else
'        sffuncMsg "年龄只能为数字，请重新录入。", sf警告
'        KeyAscii = 0
'    End If
'    Exit Sub
'errhandler:
End Sub

Private Sub ctxtAge_LostFocus()
'    On Error GoTo errhandler
'    '移动焦点。
'    If ctxtAge <> "" Then
'        If Val(ctxtAge) > 150 Or Val(ctxtAge) <= 0 Then
'            sffuncMsg "年龄必须>0，而且不可能超过150。请重新输入年龄。", sf警告
'            ctxtAge.SetFocus
'        End If
'    End If
'    Exit Sub
'errhandler:
End Sub


Private Sub ctxtSysNo_Change()
    On Error Resume Next
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
End Sub

Private Sub ctxtTubeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt体检单号.SetFocus
    End If
End Sub

Private Sub ctxt体检单号_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '热健处理。
    If KeyCode = vbKeyF8 And mblnTakePhoto Then
        If cctlCatchPhoto.VideoIsOk Then
            cctlCatchPhoto.sub转换状态
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub


Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection           '工具栏上的按钮初始化集合。
   
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    Set mcol体检项目 = New Collection
    Set mcol收费项目 = New Collection
    '界面不可操作。
    cframJBXX.Enabled = False
    ciptBase.Enabled = False
    frmPhoto.Enabled = False
    ctbMain.Enabled = False
    
    '初始化界面通用对象。
    Set mobjGUI = New cls界面通用对象
    mobjGUI.pbln自动设置字典高度 = False
    With lcol工具栏按钮
        .Add "保存"
        .Add "|"
        .Add "体检项目(&T)102"
        .Add "|"
        .Add "另存照片(&A)111"
        .Add "载入照片(&E)103"
        .Add "|"
        .Add "打印"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
        Set .c状态栏 = csbMain
        Set .c录入板 = ciptBase
        Set .c字典表 = c字典表
    End With
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""
                
    '初始时，保存按钮不可用（必须输入合法的系统编号后才可用）。
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
    
    If pobj业务对象.业务设置("是否照像") = "是" Then
        mblnTakePhoto = True
    Else
        mblnTakePhoto = False
    End If
    
    '需要照相时初始化照相控件
    If mblnTakePhoto Then
        '初始化控件
        cctlCatchPhoto.funcInitVideo
        '判断初始化是否成功
        If cctlCatchPhoto.VideoIsOk = False Then
            sffuncMsg "照相设备初始化失败，请检查原因或到业务设置中设置不进行照相！", sf警告
        End If
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pobj业务对象.业务设置("试管编号自动生成") = "否" Then
        ctxtTubeNo.Visible = True
        ctxtTubeNo.TabIndex = 1
        clblTubeNo.Visible = False
        clblLetter.Visible = False
    Else
        ctxtTubeNo.Visible = False
        clblTubeNo.Visible = True
        clblLetter.Visible = True
    End If
    
    '其他初始化工作放在定时器完成。
    Timer1.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmEditRegister", "Form_Load", 6666, lstrError, False
    
    ctbMain.Enabled = True
End Sub


'功能：完成Form_Load余下的窗体初始化操作。
Private Sub Timer1_Timer()
    Dim lobj体检表模板集 As Object
    Dim lcol单位名称集 As New Collection
    Dim lcol体检表模板集 As New Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    Timer1.Enabled = False
    
    Set lobj体检表模板集 = CreateObject("体检对象.ClsMedicalExamTemplateSet")
    lobj体检表模板集.体检表类型 = 3
    
    '将所有的体检表模板加入到ccmb体检表名
    Set lcol体检表模板集 = lobj体检表模板集.元素集
    For i = 1 To lcol体检表模板集.Count
        ccmbTemplate.AddItem lcol体检表模板集(i)
    Next i
    Set lcol体检表模板集 = Nothing
    Set lobj体检表模板集 = Nothing
    
    If ccmbTemplate.ListCount = 0 Then
        sffuncMsg "本操作现在暂时无法进行！请先进入“体检表设置”操作界面，设置各类体检表的内容！", sf警告
    End If
    
    '加入单位名称下拉列表框。
    Set lcol单位名称集 = pobj业务对象.当日工作记忆簿.单位名称集
    For i = 1 To lcol单位名称集.Count
        ccmbUnit.AddItem lcol单位名称集(i)
    Next i
    Set lcol单位名称集 = Nothing
        
    '创建本窗体所需要的全局对象。
    Set mobj体检 = CreateObject("体检对象.clsMedicalExam")
    '修改：2002-10-10（设置系统编号名称）。
    If pstr系统编号名称 <> "" Then
        mobj体检.系统编号名称 = pstr系统编号名称
    End If
    mobj体检.系统编号 = mstr系统编号
    
    mstr系统编号固定部分 = mobj体检.系统编号固定部分
    
    If mstr系统编号 <> "" Then
        ctxtSysNo = mstr系统编号
        
        '显示初始系统编号的内容。
        If mobj体检.是否已存在 Then
            subShowRegisterInfo
            
            '保存按钮可用。
            ctbMain.Buttons(1).Enabled = True
            ctbMain.Buttons(5).Enabled = True
            ctbMain.Buttons(6).Enabled = True
            ctbMain.Buttons(8).Enabled = True
            
        End If
    Else
        ctxtSysNo = mstr系统编号固定部分
    End If
    
    If pobj业务对象.业务设置("是否快速登记") = "是" Then
        mbln快速录入 = True
    Else
        mbln快速录入 = False
    End If
    
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmEditRegister", "Timer1_Timer", 6666, lstrError, False
    End If
    
    '界面可操作。
    cframJBXX.Enabled = True
    ciptBase.Enabled = True
    frmPhoto.Enabled = True
    ctbMain.Enabled = True
    If ctbMain.Buttons(1).Enabled Then
        ctxtName.SetFocus
    Else
        ctxtSysNo.SetFocus
        ctxtSysNo.SelStart = Len(ctxtSysNo)
        ctxtSysNo.SelLength = 0
    End If
    
    Exit Sub
    
    Resume
End Sub

Private Sub ccmbTemplate_Click()
    Dim i As Long
    On Error GoTo errHandler
    
    If mobj体检.体检表.体检表名 = ccmbTemplate.Text Then Exit Sub
    
    '重新设置附加信息录入板。
    On Error Resume Next
    mobjGUI.sub初始化录入板 ccmbTemplate.Text
        
    '修改：2001-8-23（显示单位属性）。
    If mstr单位申请编号 <> "" Then
        sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
    End If
        
    On Error GoTo errHandler
    mobj体检.体检表.体检表名 = ccmbTemplate.Text
    
    
    '修改：2002-7-26（杨春）根据“是否年检表”选择体检表类型。
    On Error Resume Next
    Dim lobj体检表模板 As Object
    Set lobj体检表模板 = CreateObject("体检对象.clsMedicalExamTemplate")
    lobj体检表模板.体检表名 = ccmbTemplate.Text
    If lobj体检表模板.是否年检表 Then
        ccmb体检类型.ListIndex = 1
    Else
        ccmb体检类型.ListIndex = 0
    End If
    
    '修改：2002-10-10（杨春）嘉定定制：显示体检金额。
    ciptBase.Box1("体检金额").Text = lobj体检表模板.收费标准金额
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmEditRegister", "ccmbTemplate_Click", 6666, lstrError, False
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
    '弹出列表框
    gfsubShowComboList ccmbUnit
    
    Exit Sub
errHandler:
End Sub

Private Sub ccmbUnit_LostFocus()
    Dim i As Integer
    
    On Error GoTo errHandler
    If Trim(ccmbUnit.Text) = "" Then Exit Sub
    
    '判断录入的单位是否在列表中存在，不存在则加入列表
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '加入到列表框中
        ccmbUnit.AddItem ccmbUnit.Text
        
        '加载到工作记忆簿文件中
        pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
    Else
        '修改：2001-8-26（若单位申请编号不同，修改工作记忆簿）。
        If mstr单位申请编号 <> pobj业务对象.当日工作记忆簿.单位编号(ccmbUnit.Text) Then
            pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
        End If
    End If

    Exit Sub

errHandler:
End Sub

Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '单位定位返回的结果记录。
    
    On Error GoTo errHandler
    
    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
            mstr单位申请编号 = lobjRec!申请编号
        
            '显示该体检人员所在单位的其他属性。
            '修改：2001-8-23（杨春，新增）。
            On Error Resume Next
            sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
        End If
    End If
    
    '把焦点回到单位录入框。
    ccmbUnit.SetFocus
    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmEditRegister", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub



Private Sub ctxtName_GotFocus()
    On Error Resume Next
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
End Sub

Private Sub ctxtSysNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '移动焦点。
        ctxtName.SetFocus
    End If
    
End Sub
'功能：录入系统编号后，获取并显示该体检记录的登记信息。
Private Sub ctxtSysNo_LostFocus()
    On Error GoTo errHandler
    
    '若编号未变化，不作处理。
    If mobj体检 Is Nothing Then Exit Sub
    
    If mobj体检.系统编号 = ctxtSysNo.Text Then Exit Sub
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
    
    '显示体检信息。
    mobj体检.系统编号 = ctxtSysNo.Text
    subShowRegisterInfo
    
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    ctbMain.Buttons(6).Enabled = True
    ctbMain.Buttons(8).Enabled = True
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmEditRegister", "ctxtSysNo_LostFocus", 6666, lstrError, False
    
    ctxtSysNo.SetFocus
    Exit Sub
    Resume
End Sub

Private Sub ctxtSysNo_GotFocus()
    On Error Resume Next
    With ctxtSysNo
        '显示系统编号的固定部分。
        If ctxtSysNo.Text = "" Then
            ctxtSysNo.Text = mstr系统编号固定部分
        End If
        .SelStart = Len(Trim(ctxtSysNo.Text))
        .SelLength = 0
    End With

End Sub
Private Sub subClear()
    On Error Resume Next
    ctxtName.Text = ""
    ctxtAge.Text = ""
    ccmbUnit.Text = ""
    
    '修改：体检金额不清空。
    Dim ldbl体检金额 As Double
    ldbl体检金额 = ciptBase.Box1("体检金额").Text
    ciptBase.ClearContent
    ciptBase.Box1("体检金额").Text = ldbl体检金额

End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c字典表" Then
        c字典表.Visible = False
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '释放对象。
    Set mobj体检 = Nothing
    Set mobjGUI = Nothing
    
    '设置窗体没有加载标志。
    mblnInUse = False
    
    pstr系统编号名称 = ""
End Sub

'功能：处理工具栏上按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lstr系统编号 As String
    Dim i As Integer
    Dim lcol原体检项目 As Collection
    Dim lstrFile As String
    
    On Error GoTo errHandler
    Select Case Operate
    Case "保存"
         '判断是否需要照相。
        If mblnTakePhoto = True Then
            '判断是否照相
            If cctlCatchPhoto.Photo Is Nothing Then
                sffuncMsg "没有照相，请重新照相后保存！", sf警告
                Cancel = True
                Exit Sub
            End If
        End If
        
        '若不是快速录入，检查录入是否有错误。
        If mobj体检.体检表.附加信息.Count > 0 Then
            '修改：2001-9-12（解决最后录入非法字典内容，并保存时系统服务提示错误）。
            On Error Resume Next
            ciptBase.Box1(ciptBase.ActiveInputBoxIndex).LostFocus
            On Error GoTo errHandler
            
            If ciptBase.ItemsError.Count > 0 And Not mbln快速录入 Then
                sffuncMsg "请更正黄色录入框内容！", sf警告
                ciptBase.SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
        
        MousePointer = 11
        csbMain.Panels(1) = "正在保存体检登记信息，请稍侯..."
        
        '设置体检对象属性。
        With mobj体检
            .体检人员.姓名 = Trim(ctxtName.Text)
            .体检人员.性别 = Trim(ccmbSex.Text)
            .体检人员.单位名称 = Trim(ccmbUnit.Text)
            
            If pobj业务对象.业务设置("试管编号自动生成") = "否" Then
                .试管编号 = ctxtTubeNo.Text
            End If
            .体检单号 = ctxt体检单号
            
            If mblnTakePhoto Then
                .体检人员.像片 = cctlCatchPhoto.Photo
'                .体检人员.像片压缩 = cctlCatchPhoto.Photo
            End If
            If Val(ctxtAge) > 0 Then
                .体检人员.出生日期 = DateAdd("yyyy", -Val(ctxtAge), Date)
            End If
            .体检人员.年龄 = ctxtAge
            
            On Error Resume Next
            .体检人员.公民身份号码 = ciptBase.Box1("身份证号").Text
            .体检人员.卫生种类 = ciptBase.Box1("卫生种类").TrueText
            .体检人员.片区 = ciptBase.Box1("片区").TrueText
            .体检人员.行业类别 = ciptBase.Box1("行业类别").TrueText
            
            If .体检人员.单位申请编号 <> mstr单位申请编号 Then
                .体检人员.单位申请编号 = mstr单位申请编号
            End If
            .体检日期 = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            On Error GoTo errHandler
            
            '保存附加信息
            For i = 1 To ciptBase.ItemCount
                If ciptBase.InfoCollection(i).字典名称 <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection(i).名称, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
                Else
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
                End If
            Next i
            
            '设置为体检类别。
            If ccmb体检类型.Text = "初检" Then
                .体检类别 = P_EXAM_FIRST
            Else
                .体检类别 = P_EXAM_ANNUAL
            End If
            
        End With
        
        '修改体检表。
        If mcol体检项目.Count > 0 Then
            '获取体检表上已有的体检项目。
            Set lcol原体检项目 = mobj体检.体检表.体检项目集("")
            '删除去掉的。
            For i = 1 To lcol原体检项目.Count
                If Not sffunc判断集合键值是否存在(mcol体检项目, lcol原体检项目(i).体检项目编号) Then
                    mobj体检.体检表.Sub删除体检项目 lcol原体检项目(i).体检项目编号
                End If
            Next
            '添加新增项目
            For i = 1 To mcol体检项目.Count
                mobj体检.体检表.Sub添加体检项目 mcol体检项目(i)("编码")
            Next
            
        End If
        
        '执行保存。
        If mcol收费项目.Count = 0 Then
            pobj业务对象.Sub体检登记 mobj体检
        Else
            pobj业务对象.Sub体检登记 mobj体检, , , mcol收费项目
        End If
        If cchkClear = 1 Then
            subClear
            ctbMain.Buttons(1).Enabled = False
            ctbMain.Buttons(5).Enabled = False
            ctbMain.Buttons(6).Enabled = False
            ctbMain.Buttons(8).Enabled = False
        End If
        
        '恢复照相。
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "恢复" Then
                cctlCatchPhoto.sub转换状态
            End If
        End If
        
        '焦点回到系统编号。
        ctxtSysNo.SetFocus
        ctxtSysNo.SelStart = Len(ctxtSysNo.Text)
        ctxtSysNo.SelLength = 0
        MousePointer = 0
        csbMain.Panels(1) = ""
        
        '忽略界面通用对象对本操作以后的处理。
        Cancel = True
    
    Case "体检项目"
        Dim lobj体检模板 As Object
        
        '获取体检表上已有的体检项目。
        Set lcol原体检项目 = mobj体检.体检表.体检项目集("")
        
        '设置选择项目界面的属性。
        frmSelectItem.pstr体检表名称 = ccmbTemplate.Text
        Set frmSelectItem.pcol复查项目 = lcol原体检项目
        Set frmSelectItem.pcol收费项目 = mcol收费项目
        '启动选择项目界面。
        frmSelectItem.Show 1
        If frmSelectItem.pblnOk Then
            '获取选中的复查项目。
            Set mcol体检项目 = frmSelectItem.pcol复查项目
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
    Case "另存照片"
        ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片", vbDirectory) <> "" Then
            ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片"
        End If
        ccmdFile.FileName = ctxtSysNo.Text
        ccmdFile.ShowOpen
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            SavePicture cctlCatchPhoto.Photo, lstrFile
        End If
    Case "载入照片"
        ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片", vbDirectory) <> "" Then
            ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片"
        End If
        ccmdFile.FileName = ctxtSysNo.Text
        ccmdFile.ShowOpen
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            If InStr(lstrFile, ".") > 0 Then
                Set cctlCatchPhoto.Photo = LoadPicture(lstrFile)
                mblnTakePhoto = True
            End If
        End If
    Case "打印"
        Dim lcol编号 As Collection
        Set lcol编号 = New Collection
        '打印体检表
        lcol编号.Add ctxtSysNo.Text
        pobj业务对象.Sub打印文书 "体检表", lcol编号, True
    
    End Select
    
    Exit Sub
    
errHandler:
    '错误处理。
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    '恢复界面。
    MousePointer = 0
    csbMain.Panels(1) = ""
    ctxtSysNo.SetFocus
    Cancel = True
    Exit Sub
    Resume
End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal 名称 As String, ByVal 内容 As String, ByVal 保存内容 As String, ByVal IsError As Boolean)
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    
    On Error GoTo errHandler
    ldatBirth = ""
    Select Case 名称
    Case "身份证号"
        lstrIDCard = ciptBase.ItemText(Index)
        If lstrIDCard <> "" Then
            '正确时从身份证号中获取出生日期。
            sub根据公民身份号码获取生日和性别 lstrIDCard, ldatBirth, lstrSex
            If Not IsDate(ldatBirth) Then
                Err.Raise 6666, , "身份证号不合法！"
            End If
            
            '查找是否需要录入出生日期，需要时自动根据身份证号填写出生日期
            On Error Resume Next
            If IsDate(ldatBirth) Then
                ciptBase.Box1("出生日期").Text = ldatBirth
                ctxtAge = DateDiff("yyyy", ldatBirth, Date)
            End If
            If lstrSex <> "" Then
                ccmbSex.Text = lstrSex
            End If
        End If
    Case "卫生种类"
        Dim lstrItemText  As String
        '设置行业类别录入框的字典。
        For i = 1 To ciptBase.InfoCollection.Count
            If ciptBase.InfoCollection(i).Title = "行业类别" Then
                If Not ciptBase.InfoCollection(Index + 1).DictRecordSet Is Nothing Then
                    If ciptBase.InfoCollection(Index + 1).DictRecordSet.EOF Then
                    Else
                        mobjGUI.sub初始化字典表 i, "Parent=" & ciptBase.InfoCollection(Index + 1).DictRecordSet("InnerId")
                    End If
                End If
                ciptBase.pblnTemp = True
                lstrItemText = ciptBase.Box1(i - 1).Text
                ciptBase.Box1(i - 1).Text = ""
                ciptBase.Box1(i - 1).Text = lstrItemText
                ciptBase.pblnTemp = False
                
                Exit For
            End If
        Next
    Case "工龄"
        '有效性判断。
        If 内容 <> "" Then
            If Val(内容) > 100 Then
                Err.Raise 6666, , "工龄不能大于100！"
            End If
            If Val(内容) >= Val(ctxtAge) Then
                Err.Raise 6666, , "工龄>=年龄，这是非法的数据！"
            End If
        End If
        
    End Select
    Exit Sub
errHandler:
    '错误处理。
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemSetFocus Index
    Exit Sub
    Resume


End Sub

'功能：显示体检对象“mobj体检”的属性在界面上。
Private Sub subShowRegisterInfo()
    Dim lcolInfo As New Collection '体检附加项目集合。
    Dim lstr试管编号 As String
    Dim lstrItem As String
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errHandler
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
    
    If Not mobj体检.是否已存在 Then
        '体检记录不存在时给出提示
        ctxtSysNo.SelStart = Len(ctxtSysNo)
        ctxtSysNo.SelLength = 0
        subClear
        Err.Raise 6666, , "该体检编号不存在，请重新输入体检编号。"
    End If
    '判断是否已下体检结论。
    If mobj体检.体检状态 = P_ENDED_STATUS Then
        Err.Raise 6666, , "已下体检结论，不允许修改体检登记信息。"
    End If
    '判断是否复查体检记录。
    If mobj体检.体检类别 = P_EXAM_AGAIN Then
        Err.Raise 6666, , "不允许修改复查记录。"
    End If
    '查找到该人则填写登记信息
    With mobj体检
    
        ctxtName.Text = .体检人员.姓名
        If .体检人员.性别 = "" Then
            ccmbSex.ListIndex = 0
        Else
            ccmbSex.ListIndex = gffuncItemIsInComboBox(ccmbSex, .体检人员.性别)
        End If
        If IsDate(.体检人员.出生日期) Then
            ctxtAge.Text = DateDiff("yyyy", .体检人员.出生日期, Date)
        Else
            ctxtAge.Text = ""
        End If
        ccmbUnit.Text = .体检人员.单位名称
        If ccmbTemplate.Text <> .体检表.体检表名 Then
            ccmbTemplate.Text = .体检表.体检表名
            '重新初始化附加表格
            On Error Resume Next
            mobjGUI.sub初始化录入板 ccmbTemplate.Text
            On Error GoTo errHandler
        End If
        lstr试管编号 = .试管编号
        If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
            clblLetter.Caption = Left(lstr试管编号, InStr(1, lstr试管编号, ":") - 1)
            clblTubeNo.Caption = Right(lstr试管编号, Len(lstr试管编号) - InStr(1, lstr试管编号, ":"))
        Else
            ctxtTubeNo = lstr试管编号
        End If
        If IsDate(.体检日期) Then
            cdtpDate.Value = Format(.体检日期, "yyyy-mm-dd")
        Else
            cdtpDate.Value = Date
        End If
        '修改：2001-12-29（显示体检类型）。
        If .体检类别 = P_EXAM_ANNUAL Then
            ccmb体检类型.ListIndex = 1
        Else
            ccmb体检类型.ListIndex = 0
        End If
        
        '修改：2001-8-23。
        mstr单位申请编号 = .体检人员.单位申请编号
        
        ctxt体检单号 = .体检单号
    End With
    
            
    '获取所有附加项目及其结果。
    Set lcolInfo = mobj体检.体检表.附加信息
    
    '填写附加信息结果。
    sub填录入板值 ciptBase, mobjGUI, lcolInfo
    
    '获得并显示照片。
    If Not mobj体检.体检人员.像片 Is Nothing Then
        Set cctlCatchPhoto.Photo = mobj体检.体检人员.像片
    Else
        cctlCatchPhoto.subClear
    End If
    
    '当还未开始体检时可修改体检表。
    If mobj体检.体检状态 = P_LOGIN_STATUS Then
        ccmbTemplate.Enabled = True
    Else
        ccmbTemplate.Enabled = False
    End If
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    ctbMain.Buttons(6).Enabled = True
    ctbMain.Buttons(8).Enabled = True
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面部件", "FrmEditRegister", "subShowRegisterInfo", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
