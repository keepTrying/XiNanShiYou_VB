VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "录入控件.ocx"
Begin VB.Form FrmRegisterFirst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "初检登记"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9363.768
   ScaleMode       =   0  'User
   ScaleWidth      =   10560
   StartUpPosition =   3  '窗口缺省
   Begin 录入控件.ctlInputDictGrid c字典表 
      Height          =   4455
      Left            =   5825
      TabIndex        =   25
      Top             =   2552
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7858
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
      Interval        =   3
      Left            =   2160
      Top             =   600
   End
   Begin 录入控件.ctlInputFrame ciptBase 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   7858
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      BorderStyle     =   0
      Caption         =   "Frame1"
      Rows            =   6
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
      datatypeInputBox0001=   2
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
      PermitNullInputBox0001=   0   'False
      TriggerstrInputBox0001=   ""
      允许多选InputBox0001=   0   'False
      ErrColor        =   12648447
   End
   Begin VB.Frame cframJBXX 
      BackColor       =   &H80000013&
      Caption         =   "登记基本信息："
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1485
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   10215
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6030
         TabIndex        =   26
         Top             =   420
         Width           =   255
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         ItemData        =   "FrmRegisterFirst.frx":0000
         Left            =   4440
         List            =   "FrmRegisterFirst.frx":0002
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker cdtpDate 
         Height          =   300
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   23592961
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.TextBox ctxtName 
         Height          =   300
         Left            =   240
         TabIndex        =   0
         Top             =   1080
         Width           =   2010
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "FrmRegisterFirst.frx":0004
         Left            =   2520
         List            =   "FrmRegisterFirst.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   840
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "定位(&L)"
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox ctxtAge 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   15
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表："
         Height          =   180
         Index           =   7
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label clblTubeNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "保存后请看状态栏"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6285
         TabIndex        =   22
         Top             =   450
         Width           =   1650
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5520
         TabIndex        =   21
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   270
         Width           =   900
      End
      Begin VB.Label clblSysNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管编号："
         Height          =   180
         Index           =   1
         Left            =   5520
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
         Height          =   180
         Index           =   2
         Left            =   8280
         TabIndex        =   17
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   5
         Left            =   4440
         TabIndex        =   14
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   3600
         TabIndex        =   13
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame frmPhoto 
      Caption         =   "照像："
      ForeColor       =   &H00800000&
      Height          =   4365
      Left            =   5800
      TabIndex        =   23
      Top             =   2640
      Width           =   4575
      Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
         Height          =   3720
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   6562
         BackColor       =   0
         FontSize        =   9.75
         OriginalSize    =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7080
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15849
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctlbTool 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1032
      ButtonWidth     =   820
      ButtonHeight    =   926
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchkClear 
         Caption         =   "保存后清空"
         Height          =   435
         Left            =   6885
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   3480
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
End
Attribute VB_Name = "FrmRegisterFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：邓恒
'最后修改：杨春

Private mobj体检 As Object                   '体检对象
Private mobj体检表模板 As Object             '体检表模板对象
Private mobj体检表 As Object                 '体检表对象

'业务设置。
Private mblnTakePhoto As Boolean             '是否需要照相
Private mbln快速录入 As Boolean

Private mblnInUse As Boolean
Private WithEvents mobjGUI As cls界面通用对象 '界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mcolTubeNo As New Collection          '试管编号集

Private mstr单位申请编号 As String

Private mblnSys As Boolean

Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub ccmbUnit_Click()
    Dim i As Integer
    Dim lcolInfo As New Collection
    
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
        '修改：2001-8-23。
        On Error Resume Next
        mstr单位申请编号 = pobj业务对象.当日工作记忆簿.单位编号(ccmbUnit.Text)
        sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
        
        
    End If

    Exit Sub

errHandler:
End Sub

Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '自动保存。
    If ctlbTool.Buttons(4).Enabled Then
'        ctxtName.SetFocus
'        SendKeys "{F2}"
        mobjGUI_BeforeOperate "保存", blnCancel
    End If
End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c字典表" Then
        c字典表.Visible = False
    End If

End Sub


Private Sub ctxtAge_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    '判断是否为数字
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = 13 Then
    Else
        sffuncMsg "年龄只能为数字，请重新录入。", sf警告
        KeyAscii = 0
    End If
End Sub

Private Sub ctxtAge_LostFocus()
    On Error GoTo errHandler
    '移动焦点。
    If ctxtAge <> "" Then
        If Val(ctxtAge) > 150 Or Val(ctxtAge) <= 0 Then
            sffuncMsg "年龄必须>0，而且不可能超过150。请重新输入年龄。", sf警告
            ctxtAge.SetFocus
        End If
    End If
    Exit Sub

errHandler:
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnTakePhoto Then
        '重新初始化照相控件。
        cctlCatchPhoto.funcInitVideo
    End If
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    gfsubHideComboList ccmbUnit
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

'功能：窗体初始化。
Private Sub Form_Load()
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then
        Exit Sub
    End If
    MousePointer = 11
    '设置窗体正在使用的标志。
    mblnInUse = True
    csbMain.Panels(1) = "窗体正在初始化，请稍侯..."
    
    '界面不可操作。
    cframJBXX.Enabled = False
    ciptBase.Enabled = False
    frmPhoto.Enabled = False
    ctlbTool.Enabled = False
       
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    mobjGUI.pbln自动设置字典高度 = False
    
    '设置工具栏上所需要的各种按钮。
    Dim lcol工具栏按钮 As New Collection           '工具栏上的按钮初始化集合。
    With lcol工具栏按钮
        .Add "清空"
        .Add "修改"
        .Add "|"
        .Add "保存"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlbTool
        Set .c录入板 = ciptBase
        Set .c字典表 = c字典表
        
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcol工具栏按钮, ""
    End With
    
    Set mobj体检 = CreateObject("体检对象部件.clsMedicalExam")
    Set mobj体检表模板 = CreateObject("体检对象部件.ClsMedicalExamTemplate")
    
    '清空
    subClear

    '为了加快窗体加载速度，余下初始化工作放在定时器中完成。
    Timer1.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "Form_Load", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = "窗体初始化失败！"
    ctlbTool.Enabled = True
    Exit Sub
    Resume
End Sub


'功能：为了加快窗体加载速度，完成余下初始化工作。
Private Sub Timer1_Timer()
    Dim lobj体检表模板集 As Object           '体检表模板集对象
    Dim lcol体检表模板集 As Collection
    Dim lcol单位名称集 As Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    Timer1.Enabled = False
    
    '加入单位名称
    Set lcol单位名称集 = pobj业务对象.当日工作记忆簿.单位名称集
    ccmbUnit.Clear
    For i = 1 To lcol单位名称集.Count
        ccmbUnit.AddItem lcol单位名称集(i)
    Next i
    Set lcol单位名称集 = Nothing
    
    '将所有的非复查体检表模板加入到ccmb体检表名
    Set lobj体检表模板集 = CreateObject("体检对象部件.ClsMedicalExamTemplateSet")
    lobj体检表模板集.体检表类型 = 3
    Set lcol体检表模板集 = lobj体检表模板集.元素集
    For i = 1 To lcol体检表模板集.Count
        ccmbTemplate.AddItem lcol体检表模板集(i)
    Next i
    Set lcol体检表模板集 = Nothing
    Set lobj体检表模板集 = Nothing
    
    '设置第一个体检表为缺省的体检表。
    mblnSys = True
    If ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.ListIndex = 0
    Else
        sffuncMsg "本操作现在暂时无法进行！请先进入“体检表设置”操作界面，设置各类体检表的内容！", sf警告
    End If
    mblnSys = False
    
    '获取业务设置。
    If pobj业务对象.业务设置("是否照像") = "是" Then
        mblnTakePhoto = True
    Else
        mblnTakePhoto = False
    End If
    
    If pobj业务对象.业务设置("是否快速登记") = "是" Then
        mbln快速录入 = True
    Else
        mbln快速录入 = False
    End If
    
    '根据体检表名设置附加信息录入板。
    sub设置体检表
    
    '获取当前体检表已经使用的试管编号字母。
    If mobj体检.体检表.试管编号字母 <> "" Then
        '当前体检表已固定使用了某个字母。
        clblLetter.Caption = mobj体检.体检表.试管编号字母
        cvscLetter.Enabled = False
    Else
        mobj体检.体检表.试管编号字母 = clblLetter.Caption
    End If
    
    '需要照相时初始化照相控件
    If mblnTakePhoto Then
        '初始化控件
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    '分配系统编号
    clblSysNo.Caption = mobj体检.Func分配系统编号
    
    '恢复界面可操作。
    cframJBXX.Enabled = True
    ciptBase.Enabled = True
    frmPhoto.Enabled = True
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "Form_Load", 6666, lstrError, False
    Else
        ctxtName.SetFocus
    End If
    ctlbTool.Enabled = True
    mblnSys = False
    csbMain.Panels(1) = ""
    
    '判断初始化是否成功。
    If mblnTakePhoto Then
        If Not cctlCatchPhoto.VideoIsOk Then
            csbMain.Panels(1) = "照相设备初始化失败，请检查原因或到业务设置中设置不进行照相！"
        End If
    End If
    MousePointer = 0
    
    Exit Sub
    Resume
End Sub


Private Sub cvscLetter_Change()
    On Error GoTo errHandler
    
    '点击滚动条，获得相应的字母。
    clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
        
    Exit Sub
errHandler:
    'sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "cvscLetter_Scroll", Err.Number, Err.Description, False
End Sub

Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '移动焦点。
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If

End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal 名称 As String, ByVal 内容 As String, ByVal 保存内容 As String, ByVal IsError As Boolean)
    Dim lstrIDCard As String
    Dim ldatBirth As String
    Dim lstrSex As String
    Dim i As Integer
    
    On Error GoTo errHandler
    ldatBirth = ""
    Select Case 名称
    Case "身份证号"
        lstrIDCard = 内容
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
        Dim lstrItemText As String
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
        
    Case Else
    End Select
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    ciptBase.ItemSetFocus Index
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj体检 = Nothing
    Set mobj体检表模板 = Nothing
    '关闭相机。
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    mblnInUse = False

End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    '移动焦点。
    If KeyCode = 13 Then
        ctxtAge.SetFocus
    End If
    Exit Sub
errHandler:
End Sub

'功能：选择体检表。
Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    
    If mobj体检.体检表.体检表名 = ccmbTemplate.Text Or mblnSys Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1).Text = "正在获取体检表模板信息，请稍侯..."
    
    sub设置体检表
    
    '修改：2001-8-23（显示单位属性）。
    On Error Resume Next
    If mstr单位申请编号 <> "" Then
        sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
    End If
    
    ctxtName.SetFocus
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    Err.Clear
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "ccmbTemplate_Click", 6666, lstrError, False
    
    csbMain.Panels(1).Text = ""
    MousePointer = 0
    Exit Sub
    Resume
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbUnit
    csbMain.Panels(1) = "要清空本地记录的单位名称清单，请删除文件“c:\temp\当日体检工作记忆簿.ini”。"
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    '移动焦点。
    If KeyCode = 13 Then
        '若录入板没有录入项目，则直接保存。
        If mobj体检表模板.基本附加项目集.Count > 0 Then
            ciptBase.ItemSetFocus 0
        Else
            ciptBase_LastLostFocus
        End If
    Else
        mstr单位申请编号 = ""
    End If
    Exit Sub
errHandler:
End Sub

Private Sub ccmbUnit_LostFocus()
    On Error GoTo errHandler
    Dim i As Integer
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
    Dim lobjRec As Object  '单位定位返回的结果记录。
    Dim lcolInfo As Collection
    
    On Error GoTo errHandler
    
    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ccmbUnit.Text = lobjRec("单位名称")
            mstr单位申请编号 = lobjRec!申请编号
            
            '显示该体检人员所在单位的其他属性。
            '修改：2001-8-23（杨春，新增）。
            On Error Resume Next
            sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
        End If
    End If
    
    '把焦点回到单位录入框。
    ccmbUnit.SetFocus
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub

Private Sub ctxtAge_GotFocus()
    On Error GoTo errHandler
    With ctxtAge
        .SelStart = 0
        .SelLength = Len(Trim(ctxtAge.Text))
    End With
    Exit Sub
errHandler:
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '移动焦点。
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub ctxtName_GotFocus()
    On Error GoTo errHandler
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
    Exit Sub
errHandler:

End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '移动焦点。
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub
Private Sub subClear()
    On Error Resume Next

    'clblTubeNo.Caption = ""
    ctxtName.Text = ""
    ctxtAge.Text = ""
    cdtpDate.Value = Date
    ccmbUnit.Text = ""
    ciptBase.ClearContent
    Set cctlCatchPhoto.Photo = Nothing
    
End Sub

'功能：处理工具栏上按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    On Error GoTo errHandler
    
    Select Case Operate
    Case "清空"
        '清空界面。
        subClear
        
        '忽略界面通用对象余下的处理。
        Cancel = True
        
    Case "修改"
        Dim lstr上一个号 As String
        
        '设置最后录入的系统编号。
        If mobj体检.系统编号 = "" Then
            On Error Resume Next
            lstr上一个号 = mobj体检.func获取系统编号的前一个号(clblSysNo.Caption)
            On Error GoTo errHandler
        Else
            lstr上一个号 = mobj体检.系统编号
        End If
        '关闭相机。
        If mblnTakePhoto Then
            cctlCatchPhoto.subDisconnect
        End If
        
        FrmEditRegister.系统编号 = lstr上一个号
        '启动修改界面。
        FrmEditRegister.Move Me.Left, Me.Top
        FrmEditRegister.Show 1
        
        '重新开启相机。
        If mblnTakePhoto Then
            cctlCatchPhoto.funcInitVideo
        End If
        
        '忽略界面通用对象余下的处理。
        Cancel = True
        
    Case "保存"
        '检查录入是否有错误。
        If mobj体检表模板.基本附加项目集.Count > 0 Then
            '修改：2001-9-12（杨春）。
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
         '大于150岁报错。
        If Val(ctxtAge.Text) > 150 Or (ctxtAge.Text <> "" And Val(ctxtAge.Text) <= 0) Then
            sffuncMsg "输入的年龄不合法，请重新输入合法的年龄（<150）。", sf警告
            With ctxtAge
                .SelStart = 0
                .SetFocus
            End With
            Cancel = True
            Exit Sub
        End If
        '判断是否需要照相。
        If mblnTakePhoto = True Then
            '判断是否照相
            If cctlCatchPhoto.PhotoIsOk = False Then
                sffuncMsg "没有照相，请重新照相后保存！", sf警告
                Cancel = True
                Exit Sub
            End If
        End If
        MousePointer = 11
        csbMain.Panels(1) = "正在保存体检登记信息，请稍侯..."
        
        '窗体暂时不可操作。
        ctlbTool.Enabled = False
        cframJBXX.Enabled = False
        ciptBase.Enabled = False
        frmPhoto.Enabled = False
        
        '设置体检对象属性。
        With mobj体检
            .系统编号 = clblSysNo.Caption
            .体检表.体检表名 = ccmbTemplate.Text
            .体检表.试管编号字母 = clblLetter.Caption
            .体检人员.姓名 = Trim(ctxtName.Text)
            .体检人员.性别 = ccmbSex.Text
            .体检人员.单位名称 = Trim(ccmbUnit.Text)
            
            
            .体检日期 = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            If mblnTakePhoto Then
                .体检人员.像片 = cctlCatchPhoto.Photo
            End If
            If Val(ctxtAge) > 0 Then
                .体检人员.出生日期 = DateAdd("yyyy", -Val(ctxtAge), Date)
            End If
            On Error Resume Next
            .体检人员.公民身份号码 = ciptBase.Box1("身份证号").Text
            .体检人员.卫生种类 = ciptBase.Box1("卫生种类").TrueText
            .体检人员.片区 = ciptBase.Box1("片区").TrueText
            .体检人员.行业类别 = ciptBase.Box1("行业类别").TrueText
            
            If .体检人员.单位申请编号 <> mstr单位申请编号 Then
                .体检人员.单位申请编号 = mstr单位申请编号
            End If
            On Error GoTo errHandler
            
            '设置赋加项目结果。
            For i = 1 To ciptBase.ItemCount
                '若按字典录入，设值附加项目值编号。
                If ciptBase.InfoCollection(i).字典名称 <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection(i).名称, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
                Else
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection(i).名称, ciptBase.Box1(i - 1).Text
                End If
            Next
        End With
        
        '通过体检管理业务对象执行保存。
        pobj业务对象.Sub体检登记 mobj体检
        
        csbMain.Panels(2) = "上次保存的体检系统编号：" & mobj体检.系统编号 & "，试管编号：" & mobj体检.试管编号
        If mobj体检.收费批号 <> "" Then
            csbMain.Panels(2) = csbMain.Panels(2) & "，收费批号：" & mobj体检.收费批号
        End If
        
        If cchkClear = 1 Then
            subClear
        End If
        
        '恢复照相。
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "恢复" Then
                cctlCatchPhoto.sub转换状态
            End If
        End If
        
        '新分配系统编号
        clblSysNo.Caption = mobj体检.Func分配系统编号
        
        '恢复窗体可操作。
        ctlbTool.Enabled = True
        cframJBXX.Enabled = True
        ciptBase.Enabled = True
        frmPhoto.Enabled = True
        
        '试管字母不能再选择。
        cvscLetter.Enabled = False
        
        ctxtName.SetFocus
        
        MousePointer = 0
        csbMain.Panels(1) = ""
        
        '忽略界面通用对象余下的处理。
        Cancel = True
    Case "退出"
        If Trim(ctxtName) <> "" Then
            If Not sffuncMsg("你确认要退出本界面，并且不保存当前你录入的体检人员登记信息吗？", sf询问) Then
                Cancel = True
                Exit Sub
            End If
        End If
        If clblSysNo.Caption <> "" Then
            '退回系统编号。
            mobj体检.sub退回系统编号 clblSysNo.Caption
        End If
    End Select
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    If Operate = "保存" Then
        '恢复窗体可操作。
        ctlbTool.Enabled = True
        cframJBXX.Enabled = True
        ciptBase.Enabled = True
        frmPhoto.Enabled = True
    End If
    Cancel = True
    Exit Sub
    Resume

End Sub


'功能：根据体检表下拉列表框当前内容设置体检表对象的体检表名属性，
'      并据此判断是否可以设置试管编号,初始化附加信息。
Private Sub sub设置体检表()
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer
    
    On Error GoTo errHandler
    '获取新试管编号。
    mobj体检.体检表.体检表名 = ccmbTemplate.Text
    
    '根据体检表模板获取该体检表所有可用的字母。
    If mobj体检表模板.体检表名 <> ccmbTemplate.Text Then
        mobj体检表模板.体检表名 = ccmbTemplate.Text
    End If
    
    '试管编号字母为空时cvscLetter可用
    If mobj体检.体检表.试管编号字母 = "" Then
        '将字母按逗号分开，加入mcoltubeNo中
        lstrTubeNo = mobj体检表模板.试管字母编号
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
            
            '赋值给clblLetter。
            clblLetter.Caption = mcolTubeNo(1)
            cvscLetter.Min = 1
            cvscLetter.Max = mcolTubeNo.Count
            cvscLetter.Enabled = True
            cvscLetter.Value = 1
            
        Else
            ctlbTool.Buttons(4).Enabled = False
            If ccmbTemplate.Text <> "" Then
                '提示该体检表无可用的字母。
                Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
            Else
                Exit Sub
            End If
        End If
    Else
        '有字母，不能选择字母。
        clblLetter.Caption = mobj体检.体检表.试管编号字母
        cvscLetter.Enabled = False
        
    End If
    
    '初始化附加信息。
    On Error Resume Next
    mobjGUI.sub初始化录入板 ccmbTemplate.Text
    
    ctlbTool.Buttons(4).Enabled = True
    Exit Sub
    
errHandler:
    
    sfsub错误处理 "体检界面部件", "FrmRegisterFirst", "sub设置体检表", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

