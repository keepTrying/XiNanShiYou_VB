VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "录入控件.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmRegisterLater 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "补录体检登记信息"
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
   ScaleHeight     =   7455
   ScaleWidth      =   10470
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   3120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
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
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchk刷条码 
         Caption         =   "刷条码"
         Height          =   375
         Left            =   6840
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox cchkClear 
         Caption         =   "保存后清空"
         Height          =   375
         Left            =   8520
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Value           =   1  'Checked
         Width           =   1530
      End
   End
   Begin VB.Frame cframJBXX 
      BackColor       =   &H80000013&
      Caption         =   "登记基本信息:"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   6195
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   10400
      Begin VB.TextBox ctxt体检单号 
         Height          =   315
         Left            =   8760
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox ctxtTubeNo 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame frmPhoto 
         Caption         =   "照像："
         ForeColor       =   &H00800000&
         Height          =   4305
         Left            =   5640
         TabIndex        =   29
         Top             =   1560
         Width           =   4695
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3570
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   6297
            BackColor       =   0
            FontSize        =   9.75
            OriginalSize    =   -1  'True
         End
      End
      Begin VB.TextBox ctxtSysNo 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2100
      End
      Begin 录入控件.ctlInputDictGrid c字典表 
         Height          =   4335
         Left            =   5640
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   7646
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
      Begin 录入控件.ctlInputFrame ciptBase 
         Height          =   4455
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   5385
         _ExtentX        =   9499
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
         BorderStyle     =   0
         Caption         =   "Frame1"
         Rows            =   6
         Cols            =   27
         DistanceofRow   =   0
         BorderStyle     =   0
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
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   4200
         TabIndex        =   5
         Top             =   1080
         Width           =   3360
      End
      Begin VB.TextBox ctxtAge 
         Height          =   300
         Left            =   3480
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "定位(&L)"
         Height          =   375
         Left            =   7560
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1020
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "frmRegisterLater.frx":0000
         Left            =   2280
         List            =   "frmRegisterLater.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   960
      End
      Begin VB.TextBox ctxtName 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2100
      End
      Begin VB.Frame Frame2 
         Height          =   75
         Left            =   30
         TabIndex        =   16
         Top             =   1440
         Width           =   10335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检单号："
         Height          =   180
         Index           =   9
         Left            =   8760
         TabIndex        =   34
         Top             =   840
         Width           =   900
      End
      Begin VB.Label clbl体检类型 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "初检"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8760
         TabIndex        =   32
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检类型："
         Height          =   180
         Index           =   8
         Left            =   8760
         TabIndex        =   31
         Top             =   240
         Width           =   900
      End
      Begin VB.Label clblDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7200
         TabIndex        =   28
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label clblTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2280
         TabIndex        =   26
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表："
         Height          =   180
         Index           =   7
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   22
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   5
         Left            =   4200
         TabIndex        =   21
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
         Height          =   180
         Index           =   2
         Left            =   7200
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
         Caption         =   "系统编号："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Index           =   4
         Left            =   2280
         TabIndex        =   17
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton ccmdPre 
      Appearance      =   0  'Flat
      Caption         =   "<"
      Height          =   450
      Left            =   945
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin VB.CommandButton ccmdFirst 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Height          =   450
      Left            =   450
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin VB.CommandButton ccmdNext 
      Appearance      =   0  'Flat
      Caption         =   ">"
      Height          =   450
      Left            =   1455
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin VB.CommandButton ccmdLast 
      Appearance      =   0  'Flat
      Caption         =   ">>"
      Height          =   450
      Left            =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   1560
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRegisterLater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春
'最后修改：杨春

Private WithEvents mobjGUI As cls界面通用对象 '界面通用对象，用来初始化工具栏，控制录入板。
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj体检 As Object                    '体检对象，用来保存体检登记信息。

Private mstr系统编号固定部分 As String
Private mbln快速录入 As Boolean              '业务设置：是否快速录入（若是快速录入，不作录入合法性检查）。
Private mblnTakePhoto As Boolean             '业务设置：是否照相。
Private mblnChangedPhoto As Boolean

Private mstr单位申请编号 As String

Private mblnInUse As Boolean

'修改：2003-4-15（增加该模块级全局对象，为了获取刷条码操作记忆值）。
Private mobj记忆  As cls用户操作记忆

'功能：获取本窗体是否已启动的标志。
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchk刷条码_Click()
    On Error GoTo errhandler
    
    If cchk刷条码.Value = 1 Then
        '清空系统编号输入框。
        ctxtSysNo = ""
    Else
        ctxtSysNo = mstr系统编号固定部分
    End If
    mobj体检.系统编号 = ctxtSysNo.Text
    ctxtSysNo.SelStart = Len(ctxtSysNo)
    ctxtSysNo.SelLength = 0
    ctxtSysNo.SetFocus
    Exit Sub
errhandler:
End Sub

Private Sub ccmbUnit_Click()
    On Error GoTo errhandler
    Dim i As Integer
    If ccmbUnit.Text = "" Then Exit Sub  '为空时不加入列表
    '判断录入的单位是否在列表中存在，不存在则加入列表
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '加入到列表框中
        ccmbUnit.AddItem ccmbUnit.Text
        
        '加载到工作记忆簿文件中
        pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
    Else
        '修改：2001-8-23(显示单位属性)。
        On Error Resume Next
        mstr单位申请编号 = pobj业务对象.当日工作记忆簿.单位编号(ccmbUnit.Text)
        sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
    
    End If
    Exit Sub
errhandler:
End Sub

Private Sub ciptBase_ItemLostFocus(Index As Integer)
    On Error Resume Next
    If ciptBase.InfoCollection(Index + 1).Title = "卫生种类" Then
        'mobjGUI_ItemLostFocus Index, "卫生种类", ciptBase.ItemText(Index), ciptBase.ItemTrueText(Index), False
    End If

End Sub

Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '自动保存。
    If ctbMain.Buttons(3).Enabled Then
'        ctxtSysNo.SetFocus
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

Private Sub ctxtTubeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub ctxt体检单号_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '若录入板没有录入项目，则直接保存。
        If mobj体检.体检表.附加信息.Count > 0 Then
            ciptBase.SetFocus
        Else
            ciptBase_LastLostFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtSysNo.SetFocus
    ctxtSysNo.SelStart = Len(ctxtSysNo)
    ctxtSysNo.SelLength = 0
    
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


Private Sub Form_Load()
    Dim lcolInfo As Collection
    Dim i As Integer
    
    On Error GoTo errhandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1) = "窗体正在初始化，请稍侯..."
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    '界面暂时不可操作。
    cframJBXX.Enabled = False
    ctbMain.Enabled = False
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    mobjGUI.pbln自动设置字典高度 = False
    
    Set lcolInfo = New Collection
    With lcolInfo
        .Add "查询(&Q)105"
        .Add "|"
        .Add "保存"
        .Add "另存照片(&A)111"
        .Add "更改照片(&E)103"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
        Set .c录入板 = ciptBase
        Set .c字典表 = c字典表
        
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcolInfo, ""
    End With

    '加入单位名称
    Set lcolInfo = pobj业务对象.当日工作记忆簿.单位名称集
    For i = 1 To lcolInfo.Count
        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    '创建体检对象。
    Set mobj体检 = CreateObject("体检对象.clsMedicalExam")
    mstr系统编号固定部分 = mobj体检.系统编号固定部分
    mobj体检.系统编号 = mstr系统编号固定部分
    
    ctxtSysNo = mstr系统编号固定部分
    ctxtSysNo.SelLength = Len(ctxtSysNo)
    ctxtSysNo.SelStart = 0
    
    ctbMain.Buttons(3).Enabled = False
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    If pobj业务对象.业务设置("是否快速登记") = "是" Then
        mbln快速录入 = True
    Else
        mbln快速录入 = False
    End If

    If pobj业务对象.业务设置("是否照像") = "是" Then
        mblnTakePhoto = True
    Else
        mblnTakePhoto = False
    End If
    
    
    '需要照相时初始化照相控件
    If mblnTakePhoto Then
        '初始化控件
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pobj业务对象.业务设置("试管编号自动生成") = "否" Then
        ctxtTubeNo.Enabled = True
    Else
        ctxtTubeNo.Enabled = False
    End If
        
    
    '创建记忆对象。
    On Error Resume Next
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = um用户编号
    mobj记忆.业务名 = "体检补录登记"
    Dim lstrOption As String
    lstrOption = mobj记忆.记忆项值("刷条码")
    If lstrOption = "是" Then
        cchk刷条码.Value = 1
    End If
    
    Err.Clear
    
errhandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "frmRegisterLater", "Form_Load", 6666, lstrError, False
    End If
    '界面可操作。
    cframJBXX.Enabled = True
    ctbMain.Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    
    Exit Sub
    Resume
End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtAge.SetFocus
    End If
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error GoTo errhandler
    '获得焦点时弹出下拉框
    gfsubShowComboList ccmbUnit
    Exit Sub
errhandler:
End Sub

Private Sub ccmbUnit_LostFocus()
    On Error GoTo errhandler
    Dim i As Integer
    If Trim(ccmbUnit.Text) = "" Then Exit Sub  '为空时不加入列表
    
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
errhandler:
    
End Sub

Private Sub ccmdLocateUnit_Click()
    On Error GoTo errhandler
    Dim lobjRec As Object  '单位定位返回的结果记录。

    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ccmbUnit.Text = lobjRec("单位名称")
            mstr单位申请编号 = lobjRec!申请编号
            
            '修改：2001-8-23（显示单位属性）。
            On Error Resume Next
            sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
            
        End If
    End If
    
    '把焦点回到单位录入框。
    ccmbUnit.SetFocus
    SendKeys vbTab
    Exit Sub
errhandler:
    'sfsub错误处理 "体检界面部件", "frmRegisterLater", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub



Private Sub ctxtAge_GotFocus()
    On Error GoTo errhandler
    With ctxtAge
        .SelStart = 0
        .SelLength = Len(Trim(ctxtAge.Text))
    End With
    Exit Sub
errhandler:
    'sfsub错误处理 "体检界面部件", "frmRegisterLater", "ctxtAge_GotFocus", Err.Number, Err.Description, False
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub ctxtAge_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler
    '判断是否为数字
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = 13 Then
    Else
        sffuncMsg "年龄只能为数字，请重新录入。", sf警告
        KeyAscii = 0
    End If
    Exit Sub
    
errhandler:
End Sub

Private Sub ctxtAge_LostFocus()
    On Error GoTo errhandler
    '移动焦点。
    If ctxtAge <> "" Then
        If Val(ctxtAge) > 150 Or Val(ctxtAge) <= 0 Then
            sffuncMsg "年龄必须>0，而且不可能超过150。请重新输入年龄。", sf警告
            ctxtAge.SetFocus
        End If
    End If
    Exit Sub
errhandler:
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub

Private Sub ctxtSysNo_GotFocus()
    On Error Resume Next
    With ctxtSysNo
        '显示系统编号的固定部分。
        If ctxtSysNo.Text = "" And cchk刷条码.Value = 0 Then
            ctxtSysNo.Text = mstr系统编号固定部分
        End If
        .SelStart = Len(Trim(ctxtSysNo.Text))
        .SelLength = 0
    End With
End Sub

Private Sub ctxtSysNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '显示体检人员的基本信息，移动焦点。
    If KeyCode = 13 Then
        If ctxtTubeNo.Enabled Then
            ctxtTubeNo.SetFocus
        Else
            ctxtName.SetFocus
        End If
    End If
End Sub

Private Sub ctxtSysNo_LostFocus()
    Dim lcol体检附加项目 As New Collection
    Dim lstr试管编号 As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo errhandler
    
    '若编号未变化，不作处理。
    If mobj体检.系统编号 = ctxtSysNo.Text Or (Len(ctxtSysNo.Text) = Len(mstr系统编号固定部分) And mstr系统编号固定部分 <> "") Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1) = "正在搜索体检人员信息，请稍侯..."
    
    ctbMain.Buttons(3).Enabled = False
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    mobj体检.系统编号 = Trim(ctxtSysNo.Text)
    ctxtName.Text = ""
    cctlCatchPhoto.subClear
    mblnChangedPhoto = False
    
    If Not mobj体检.是否已存在 Then
        '体检记录不存在时给出提示
        Err.Raise 6666, , "该体检编号不存在，请重新输入体检编号。"
    Else
        '判断是否已下体检结论。
        If mobj体检.体检状态 = P_ENDED_STATUS Then
            Err.Raise 6666, , "已下体检结论，不允许补录体检登记信息。"
        End If
        '判断是否复查体检记录。
        If mobj体检.体检类别 = P_EXAM_AGAIN Then
            Err.Raise 6666, , "复查记录不允许作补录。"
        End If
        
        '查找到该人则填写登记信息在界面上。
        With mobj体检
            ctxt体检单号 = .体检单号
            
            ctxtName.Text = .体检人员.姓名
            i = gffuncItemIsInComboBox(ccmbSex, .体检人员.性别)
            ccmbSex.ListIndex = i
            
            If IsDate(.体检人员.出生日期) Then
                ctxtAge.Text = DateDiff("yyyy", .体检人员.出生日期, Date)
            Else
                ctxtAge.Text = ""
            End If
            ccmbUnit.Text = .体检人员.单位名称
            If clblTemplate.Caption <> .体检表.体检表名 Then
                clblTemplate.Caption = .体检表.体检表名
                '重新初始化附加信息录入板。
                On Error Resume Next
                mobjGUI.sub初始化录入板 clblTemplate.Caption
                On Error GoTo errhandler
            End If
            ctxtTubeNo.Text = .试管编号
            
            If IsDate(.体检日期) Then
                clblDate.Caption = Format(.体检日期, "yyyy-mm-dd")
            Else
                clblDate.Caption = .体检日期
            End If
            '修改：2001-8-23。
            mstr单位申请编号 = .体检人员.单位申请编号
            
            '修改：2001-12-29（显示体检类型）。
            If .体检类别 = P_EXAM_ANNUAL Then
                clbl体检类型.Caption = "年检"
            Else
                clbl体检类型.Caption = "初检"
            End If
        End With
        
                
        '获取所有附加项目及其结果。
        Set lcol体检附加项目 = mobj体检.体检表.附加信息
        
        '填写附加信息结果。
        sub填录入板值 ciptBase, mobjGUI, lcol体检附加项目
        
        '获得并显示照片。
        If Not mobj体检.体检人员.像片 Is Nothing Then
            Set cctlCatchPhoto.Photo = mobj体检.体检人员.像片
        Else
            cctlCatchPhoto.subClear
        End If
        
        If ctxtTubeNo.Enabled Then
            ctxtTubeNo.SetFocus
        Else
            ctxtName.SetFocus
        End If
    End If
    ctbMain.Buttons(3).Enabled = True
    ctbMain.Buttons(4).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    
errhandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmRegisterLater", "ctxtSysNo_LostFocus", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    If ctxtSysNo.Enabled Then
        ctxtSysNo.SetFocus
    End If
    ctxtSysNo = mstr系统编号固定部分
    ctxtSysNo.SelStart = Len(ctxtSysNo)
    ctxtSysNo.SelLength = 0
    Exit Sub
    Resume
End Sub
Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt体检单号.SetFocus
    Else
        mstr单位申请编号 = ""
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj体检 = Nothing
    Set mobjGUI = Nothing
    
    '关闭相机。
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    
    '修改：2003-4-15（保存操作记忆）。
    mobj记忆.sub覆盖记忆值 "刷条码", IIf(cchk刷条码.Value = 1, "是", "否")
    
    Set mobj记忆 = Nothing
    
    mblnInUse = False
End Sub

'功能：处理工具栏上按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Dim i As Integer
    Dim lstr流水号 As String
    Dim lstrFile As String
    
    On Error GoTo errhandler
    
    Select Case Operate
    
    Case "查询"
        '显示该体检人员的基本信息。
        Dim lstr系统编号 As String
        frm查找人员.pstr体检类型 = "初检"
        frm查找人员.Show 1, Me
        lstr系统编号 = frm查找人员.pstr系统编号
        If lstr系统编号 <> "" Then
            ctxtSysNo.Text = lstr系统编号
            '显示体检人员基本信息。
            ctxtSysNo_LostFocus
        End If
    
    
    Case "保存"
         '大于150岁报错。
        If Val(ctxtAge.Text) > 150 Or (ctxtAge.Text <> "" And Val(ctxtAge.Text) <= 0) Then
            sffuncMsg "输入的年龄不合法，请重新输入。", sf警告
            ctxtAge.SetFocus
            Cancel = True
            Exit Sub
        End If
        If Trim(ctxtTubeNo.Text) = "" Then
            sffuncMsg "试管编号必须输入！", sf警告
            ctxtTubeNo.SetFocus
            Cancel = True
            Exit Sub
        End If
        
        '检查录入是否有错误。
        If mobj体检.体检表.附加信息.Count > 0 Then
            '修改：2001-9-12（杨春）。
            On Error Resume Next
            ciptBase.Box1(ciptBase.ActiveInputBoxIndex).LostFocus
            On Error GoTo errhandler
            
            If ciptBase.ItemsError.Count > 0 And Not mbln快速录入 Then
                sffuncMsg "请更正黄色录入框内容！", sf警告
                Cancel = True
                Exit Sub
            End If
        End If
        MousePointer = 11
        csbMain.Panels(1) = "正在保存体检登记信息，请稍侯..."
        
        '产生试管编号并保存
        With mobj体检
            .体检人员.姓名 = Trim(ctxtName.Text)
            .体检人员.性别 = ccmbSex.Text
            .体检人员.单位名称 = Trim(ccmbUnit.Text)
            
            If pobj业务对象.业务设置("试管编号自动生成") = "否" Then
                .试管编号 = ctxtTubeNo.Text
            End If
            .体检单号 = ctxt体检单号.Text
            
            If Val(ctxtAge) > 0 Then
                .体检人员.出生日期 = DateAdd("yyyy", -Val(ctxtAge), Date)
            End If
            If mblnTakePhoto Or mblnChangedPhoto Then
                .体检人员.像片 = cctlCatchPhoto.Photo
            End If
            On Error Resume Next
            .体检人员.公民身份号码 = ciptBase.Box1("身份证号").Text
            .体检人员.卫生种类 = ciptBase.Box1("卫生种类").TrueText
            .体检人员.片区 = ciptBase.Box1("片区").TrueText
            .体检人员.行业类别 = ciptBase.Box1("行业类别").TrueText
            
            If .体检人员.单位申请编号 <> mstr单位申请编号 Then
                .体检人员.单位申请编号 = mstr单位申请编号
            End If
            
            On Error GoTo errhandler
            '保存附加信息
            For i = 1 To ciptBase.ItemCount
                'If ciptBase.Box1(i - 1).TrueText <> ciptBase.Box1(i - 1).Text And ciptBase.Box1(i - 1).Text <> "" Then
                If ciptBase.InfoCollection(i).字典名称 <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection(i).名称, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
                Else
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
                End If
            Next i
            
            .试管编号 = ctxtTubeNo.Text
        End With
        
        pobj业务对象.Sub体检登记 mobj体检
        
        If cchkClear = 1 Then
            subClear
            mobj体检.系统编号 = ""
            ctbMain.Buttons(3).Enabled = False
            ctbMain.Buttons(4).Enabled = False
            ctbMain.Buttons(5).Enabled = False
            
        End If
        
        '恢复照相。
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "恢复" Then
                cctlCatchPhoto.sub转换状态
            End If
        End If
        
        mblnChangedPhoto = False
        
        ctxtSysNo.SelStart = Len(ctxtSysNo)
        ctxtSysNo.SelLength = 0
        ctxtSysNo.SetFocus
        
        MousePointer = 0
        csbMain.Panels(1) = ""
        Cancel = True
        
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
    Case "更改照片"
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
                mblnChangedPhoto = True
            End If
        End If
    End Select
    
    Exit Sub
errhandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "frmRegisterLater", "mobjGUI_BeforeOperate", 6666, lstrError, False
    End If
    MousePointer = 0
    csbMain.Panels(1) = ""
    Cancel = True
    Exit Sub
    Resume
    Exit Sub
End Sub
'清空
Private Sub subClear()
    On Error Resume Next
    If cchk刷条码.Value = 1 Then
        ctxtSysNo.Text = ""
    Else
        ctxtSysNo.Text = mstr系统编号固定部分
    End If
    clblTemplate.Caption = ""
    ctxtTubeNo.Text = ""
    ctxtName.Text = ""
    If ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
    ctxtAge.Text = ""
    ccmbUnit.Text = ""
    cctlCatchPhoto.subClear
    ciptBase.ClearContent
    
End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal 名称 As String, ByVal 内容 As String, ByVal 保存内容 As String, ByVal IsError As Boolean)
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    
    On Error GoTo errhandler
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
errhandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterLater", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemBox(Index).Text = ""
    ciptBase.ItemSetFocus Index
    Exit Sub
    Resume

End Sub
