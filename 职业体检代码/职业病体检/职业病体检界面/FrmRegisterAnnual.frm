VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "录入控件.ocx"
Begin VB.Form FrmRegisterAnnual 
   BorderStyle     =   0  'None
   Caption         =   "体检登记--年检登记"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8229.114
   ScaleMode       =   0  'User
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   3600
      Top             =   480
   End
   Begin VB.ListBox clstPersonList 
      Height          =   1680
      Left            =   5280
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Frame cfram基本信息 
      BackColor       =   &H80000013&
      Caption         =   "登记基本信息:"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   5595
      Left            =   60
      TabIndex        =   21
      Top             =   1920
      Width           =   10380
      Begin VB.Frame frmPhoto 
         Caption         =   "照像："
         ClipControls    =   0   'False
         ForeColor       =   &H00800000&
         Height          =   4035
         Left            =   5760
         TabIndex        =   40
         Top             =   1560
         Width           =   4575
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3720
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   6562
            BackColor       =   0
            FontSize        =   9.75
            OriginalSize    =   -1  'True
         End
      End
      Begin VB.CommandButton ccmd单位定位 
         Caption         =   "定位(&T)"
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   945
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Top             =   1200
         Width           =   3120
      End
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6480
         TabIndex        =   10
         Top             =   600
         Width           =   345
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   3120
      End
      Begin 录入控件.ctlInputDictGrid c字典表 
         Height          =   3735
         Left            =   5640
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   6588
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
      Begin MSComCtl2.DTPicker cdtpDate 
         Height          =   315
         Left            =   8760
         TabIndex        =   39
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   72024065
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin 录入控件.ctlInputFrame ciptBase 
         Height          =   3780
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6668
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
         Cols            =   27
         DistanceofRow   =   0
         BorderStyle     =   0
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
      Begin VB.Label clblAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1920
         TabIndex        =   37
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label clblSex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label clblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "性别："
         Height          =   180
         Left            =   1320
         TabIndex        =   34
         Top             =   990
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   990
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "体检表："
         Height          =   180
         Left            =   2760
         TabIndex        =   30
         Top             =   315
         Width           =   720
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "保存后请看状态栏"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6840
         TabIndex        =   29
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6120
         TabIndex        =   28
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   315
         Width           =   900
      End
      Begin VB.Label clblSysNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管编号："
         Height          =   180
         Index           =   1
         Left            =   6120
         TabIndex        =   25
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
         Height          =   180
         Index           =   2
         Left            =   8760
         TabIndex        =   24
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   5
         Left            =   2760
         TabIndex        =   23
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   1920
         TabIndex        =   22
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.Frame cframSearch 
      Caption         =   "查找年检人员："
      ForeColor       =   &H00800000&
      Height          =   1020
      Left            =   45
      TabIndex        =   20
      Top             =   840
      Width           =   10395
      Begin VB.OptionButton coptChoise 
         Caption         =   "身份证号"
         Height          =   240
         Index           =   2
         Left            =   4320
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox ctxtId 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   7
         Top             =   600
         Width           =   2025
      End
      Begin VB.ComboBox ccmbQueryUnit 
         Height          =   300
         Left            =   4920
         TabIndex        =   2
         Top             =   225
         Width           =   3735
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "FrmRegisterAnnual.frx":0000
         Left            =   3240
         List            =   "FrmRegisterAnnual.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox ctxtHealthNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   2025
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "健康证号"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton ccmdSearch 
         Caption         =   "显示人员(F5)"
         Height          =   360
         Left            =   8760
         TabIndex        =   8
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox ctxtName 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   990
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "姓名"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "定位(&L)"
         Height          =   375
         Left            =   8760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原单位"
         Height          =   180
         Index           =   7
         Left            =   4320
         TabIndex        =   32
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   2760
         TabIndex        =   31
         Top             =   300
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   7455
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16007
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchkClear 
         Caption         =   "保存后清空"
         Height          =   345
         Left            =   5160
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Value           =   1  'Checked
         Width           =   1410
      End
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
Attribute VB_Name = "FrmRegisterAnnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：邓恒
'最后修改：杨春

Private mobj旧体检 As Object                   '年检人员的最近次的体检。
Private mobj体检 As Object                     '新体检对象，提供获取系统编号和试管编号，保存登记信息的方法。
Private mobj体检集 As Object                   '体检集，用来定位需要年检的体检人员信息。
Private mobj体检表模板 As Object               '体检表模板，获取所有的非复查体检表模板名称。
Private WithEvents mobjGUI As cls界面通用对象  '界面通用对象，用来初始化工具栏，控制录入板控件。
Attribute mobjGUI.VB_VarHelpID = -1

'业务设置。
Private mblnTakePhoto As Boolean               '业务设置‘是否照相’。
Private mbln快速录入 As Boolean

Private mcolTubeNo As New Collection           '当前体检表可选的试管字母。

Private mstr单位申请编号 As String             '单位定位出申请编号。
Private mblnInUse As Boolean

'功能：记载当前窗体是否已加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub ccmbTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        cdtpDate.SetFocus
    End If
End Sub

Private Sub ccmbUnit_Click()
    On Error GoTo errHandler
    Dim i As Integer
    
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

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
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
    If ctbMain.Buttons(4).Enabled Then
        ccmbUnit.SetFocus
        SendKeys "{F2}"
        'mobjGUI_BeforeOperate "保存", blnCancel
    End If
End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c字典表" Then
        c字典表.Visible = False
    End If

End Sub

'功能：点击一条编号，显示出该系统编号的体检人员的信息。
Private Sub clstPersonList_DblClick()
    Dim lobj体检人员  As Object 'clsPersonExamed.
    Dim lobj体检 As Object      'clsMedicalExam
    Dim lstrItem As String      '选中行内容、健康档案编号、系统编号。
    On Error GoTo errHandler
    
    With clstPersonList
        '算出健康档案编号。
        lstrItem = .List(.ListIndex)
        lstrItem = Left(lstrItem, InStr(lstrItem, " ") - 1)
        
        '获取该人最近一次体检记录的系统编号。
        If coptChoise(0).Value Then
            lstrItem = func根据健康档案编号获取系统按号(lstrItem)
        End If
        
        '显示体检人员信息。
        SubGetPersonInfo lstrItem
        
        '本列表框消失。
        .Visible = False
        
    End With
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "clstPersonList_Click", 6666, lstrError, False
    
End Sub

Private Sub clstPersonList_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And clstPersonList.ListIndex >= 0 Then
        clstPersonList_DblClick
    End If
End Sub

Private Sub ctxtId_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
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
    gfsubHideComboList ccmbQueryUnit
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
   
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    MousePointer = 11
    csbMain.Panels(1) = "窗体正在初始化，请稍侯..."
    
    '界面不可操作。
    cframSearch.Enabled = False
    cfram基本信息.Enabled = False
    ctbMain.Enabled = False
    
    
    Set mobj旧体检 = CreateObject("体检对象部件.clsMedicalExam")
    Set mobj体检 = CreateObject("体检对象部件.clsMedicalExam")
    Set mobj体检集 = CreateObject("体检对象部件.clsMedicalExamSet")
    Set mobj体检表模板 = CreateObject("体检对象部件.ClsMedicalExamTemplate")
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
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
        Set .c工具栏 = ctbMain
        Set .c录入板 = ciptBase
        Set .c字典表 = c字典表
        'Set .c状态栏 = csbMain
        
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcol工具栏按钮, ""
    End With
    
    '清空
    subClear
    
    '为了加快窗体加载速度，余下初始化工作放在定时器中完成。
    Timer1.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "Form_Load", 6666, lstrError, False
    '恢复工具栏可用。
    ctbMain.Enabled = True
    MousePointer = 0
    csbMain.Panels(1) = "窗体初始化失败！"
End Sub

'功能：完成form_load余下的初始化工作。
Private Sub Timer1_Timer()
    Dim lobj体检表模板集 As Object  '体检表模板集，获取所有的非复查体检表模板名称。
    Dim lcolInfo As Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '定时器不再起作用。
    Timer1.Enabled = False
    
    '从当日工作及已簿中获取当天录入过的单位名称。
    Set lcolInfo = pobj业务对象.当日工作记忆簿.单位名称集
    For i = 1 To lcolInfo.Count
        ccmbQueryUnit.AddItem lcolInfo(i)
        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    '将所有的非复查体检表模板加入到体检表下拉列表框中。
    Set lobj体检表模板集 = CreateObject("体检对象部件.ClsMedicalExamTemplateSet")
    lobj体检表模板集.体检表类型 = 3
    Set lcolInfo = lobj体检表模板集.元素集
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
    Next
    Set lobj体检表模板集 = Nothing
    
    '根据业务设置判断是否照相。
    If pobj业务对象.业务设置("是否照像") = "是" Then
        mblnTakePhoto = True
    Else
        mblnTakePhoto = False
    End If
    
    '需要照相时初始化照相控件。
    If mblnTakePhoto Then
        '初始化控件。
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pobj业务对象.业务设置("是否快速登记") = "是" Then
        mbln快速录入 = True
    Else
        mbln快速录入 = False
    End If
     
    '恢复界面可操作。
    cframSearch.Enabled = True
    cfram基本信息.Enabled = True
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "Timer1_Timer", 6666, lstrError, False
    End If
    ctbMain.Enabled = True
    MousePointer = 0
    csbMain.Panels(1) = ""
    If cframSearch.Enabled Then
        ctxtName.SetFocus
    End If
End Sub


'功能：自动弹出列表框
Private Sub ccmbQueryUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbQueryUnit
    Exit Sub
errHandler:
End Sub

Private Sub ccmbQueryUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ccmbQueryUnit_LostFocus()
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '判断录入的单位是否在列表中存在，不存在则加入列表。
    i = gffuncItemIsInComboBox(ccmbQueryUnit, ccmbQueryUnit.Text)
    
    If i = -1 Then
        '加到ccmbQueryUnit中。
        ccmbQueryUnit.AddItem ccmbQueryUnit.Text
    End If
    
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbQueryUnit.SetFocus
    End If
End Sub

Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    '选择体检表
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    MousePointer = 11
    csbMain.Panels(1).Text = "正在获取体检表模板信息，请稍侯..."
    
    '获取新试管编号。
    If mobj体检.体检表.体检表名 <> ccmbTemplate.Text Then
        mobj体检.体检表.体检表名 = ccmbTemplate.Text
        
        '根据体检表模板获取该体检表所有可用的字母。
        mobj体检表模板.体检表名 = ccmbTemplate.Text
        
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
                
                '赋值给clblLetter
                clblLetter.Caption = mcolTubeNo(1)
                cvscLetter.Enabled = True
                cvscLetter.Min = 1
                cvscLetter.Max = mcolTubeNo.Count
                cvscLetter.Value = 1
            Else
                ctbMain.Buttons(4).Enabled = False
                '提示该体检表无可用的字母。
                Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
            End If
        Else
            '有字母，不能选择字母。
            clblLetter.Caption = mobj体检.体检表.试管编号字母
            cvscLetter.Enabled = False
        End If
        
        '初始化附加信息。
        On Error Resume Next
        mobjGUI.sub初始化录入板 ccmbTemplate.Text
        
        '修改：2001-8-23（显示单位属性）。
        If mstr单位申请编号 <> "" Then
            sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
        End If
        
        '重新填写附加信息值。
        If mobj体检表模板.基本附加项目集.Count > 0 Then
            Set lcolInfo = mobj旧体检.体检表.附加信息
            If lcolInfo.Count > 0 Then
                sub填录入板值 ciptBase, mobjGUI, lcolInfo
            End If
        End If
        DoEvents
    End If
    ctbMain.Buttons(4).Enabled = True
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "ccmbTemplate_Click", 6666, lstrError, False
    ciptBase.pblnTemp = False
    Exit Sub
    Resume
End Sub

'自动弹出列表框
Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbUnit
    Exit Sub
errHandler:
    'sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "ccmbUnit_GotFocus", Err.Number, Err.Description, False
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
'调用单位定位
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '单位定位返回的结果记录。

    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ccmbQueryUnit.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
        End If
    End If
    
    '把焦点回到单位录入框。
    ccmbQueryUnit.SetFocus
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub

Private Sub ccmdSearch_Click()
    Dim lobj体检人员 As Object  '用来获取体检人员最近次体检记录。
    Dim lobj体检 As Object      '体检人员最近次体检。
    Dim lobjRec As Object       'Recordset，体检集对象返回的元素集。
    Dim lstr系统编号 As String  '健康证号对应的体检系统编号。
    Dim i As Integer
    
    On Error GoTo errHandler
    MousePointer = 11
    csbMain.Panels(1) = "正在查找满足指定条件的体检人员，请稍候..."
    
    '先清空界面。
    subClear
    
    If coptChoise(0).Value Then
        If Trim(ctxtName.Text) = "" And ccmbSex.Text = "" And Trim(ccmbQueryUnit.Text) = "" Then
            Err.Raise 6666, , "必须输入姓名、性别、单位名称！"
        End If
        '设置体检集对象的搜索定位条件属性。
        With mobj体检集
            .subClear
            .姓名 = Trim(ctxtName.Text)
            .性别 = ccmbSex.Text
            .单位名称 = Trim(ccmbQueryUnit.Text)
        End With
        
        '获取满足指定定位条件的体检记录。
        Set lobjRec = mobj体检集.元素集("distinct 健康档案编号,姓名,单位名称")
        If lobjRec.recordcount = 0 Then
            '没找到相应体检人员。
            Err.Raise 6666, , "该体检人员还没有在本体检中心体检过，无法进行年检登记。请选择初检登记。"
        Else
            If lobjRec.recordcount > 1 Then
                '查找到多条记录时弹出list，并加入到体检人员列表框中。
                clstPersonList.Clear
                Do While Not lobjRec.EOF
                    clstPersonList.AddItem lobjRec("健康档案编号") & " " & lobjRec("姓名") & " " & IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
                    lobjRec.movenext
                Loop
                '列表可见
                clstPersonList.Visible = True
                clstPersonList.SetFocus
                
            Else
                '只找到一个人,获取该人最近一次体检记录的系统编号。
                lstr系统编号 = func根据健康档案编号获取系统按号(lobjRec!健康档案编号)
            End If
        End If
        
        lobjRec.Close
        
    ElseIf coptChoise(1).Value Then
        '将健康证号转换成系统编号。
        If Trim(ctxtHealthNo.Text) = "" Then
            Err.Raise 6666, , "你必须输入健康证号！"
        End If
        lstr系统编号 = pobj业务对象.Func根据健康证条码号获取体检系统编号(Trim(ctxtHealthNo.Text))
        If lstr系统编号 = "" Then
            Err.Raise 6666, , "你输入的健康证号没有对应的体检记录。"
        End If
    Else
        '按身份证号查询。
        If Trim(ctxtId.Text) = "" Then
            Err.Raise 6666, , "你必须输入身份证号！"
        End If
        '设置体检集对象的搜索定位条件属性。
        With mobj体检集
            .subClear
            .身份证号 = Trim(ctxtId.Text)
        End With
        
        '获取满足指定定位条件的体检记录。
        Set lobjRec = mobj体检集.元素集("系统编号,姓名,单位名称")
        If lobjRec.recordcount = 0 Then
            '没找到相应体检人员。
            Err.Raise 6666, , "该体检人员还没有在本体检中心体检过，无法进行年检登记。请选择初检登记。"
        ElseIf lobjRec.recordcount = 1 Then
            lstr系统编号 = lobjRec("系统编号")
        Else
            '查找到多条记录时弹出list，并加入到体检人员列表框中。
            clstPersonList.Clear
            Do While Not lobjRec.EOF
                clstPersonList.AddItem lobjRec("系统编号") & " " & lobjRec("姓名") & " " & IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
                lobjRec.movenext
            Loop
            '列表可见
            clstPersonList.Visible = True
            clstPersonList.SetFocus
            
        End If
    End If
    
    '显示该体检人员的基本信息。
    If lstr系统编号 <> "" Then
        SubGetPersonInfo lstr系统编号
    End If
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "ccmdSearch_Click", 6666, lstrError, False
        
        If coptChoise(0).Value Then
            ctxtName.SetFocus
        ElseIf coptChoise(1).Value Then
            ctxtHealthNo.SetFocus
        Else
            ctxtId.SetFocus
        End If
    End If
    Set lobj体检人员 = Nothing
    Set lobj体检 = Nothing
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
    Case vbKeyF5
        '显示人员。
        ccmdSearch_Click
    Case vbKeyF8
        If mblnTakePhoto Then
            If cctlCatchPhoto.VideoIsOk Then
                cctlCatchPhoto.sub转换状态
            End If
        End If
    
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

'功能：窗体初始化。

'调用单位定位
Private Sub ccmd单位定位_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '单位定位返回的结果记录。

    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
            mstr单位申请编号 = lobjRec!申请编号
            
            If mstr单位申请编号 <> "" Then
                '修改：2001-8-23（显示单位属性）。
                On Error Resume Next
                sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
            End If
        End If
    End If
    
    '把焦点回到单位录入框。
    ccmbUnit.SetFocus
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "ccmd单位定位_Click", 6666, lstrError, False
End Sub


Private Sub clstPersonList_LostFocus()
    On Error Resume Next
    clstPersonList.Visible = False
End Sub


Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal 名称 As String, ByVal 内容 As String, ByVal 保存内容 As String, ByVal IsError As Boolean)
    On Error GoTo errHandler
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    

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
                clblAge.Caption = DateDiff("yyyy", ldatBirth, Date)
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
            If Val(内容) >= Val(clblAge) Then
                Err.Raise 6666, , "工龄>=年龄，这是非法的数据！"
            End If
        End If
        
    End Select
    Exit Sub

errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemBox(Index).Text = ""
    ciptBase.ItemSetFocus Index
End Sub

Private Sub coptChoise_Click(Index As Integer)
    On Error GoTo errHandler
    ctxtName.Enabled = False
    ccmbSex.Enabled = False
    ccmbQueryUnit.Enabled = False
    ccmdLocateUnit.Enabled = False
    ctxtId.Enabled = False
    ctxtHealthNo.Enabled = False
    
    If coptChoise(0).Value Then
        '选择输入姓名。
        ctxtName.Enabled = True
        ccmbSex.Enabled = True
        ccmbQueryUnit.Enabled = True
        ccmdLocateUnit.Enabled = True
        ctxtName.SetFocus
    ElseIf coptChoise(1).Value Then
        '选择输入健康档案编号。
        ctxtHealthNo.Enabled = True
        ctxtHealthNo.SetFocus
    ElseIf coptChoise(2).Value Then
        '选择输入身份证号。
        ctxtId.Enabled = True
        ctxtId.SetFocus
    End If
    
    Exit Sub
errHandler:
    'sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "coptChoise_Click", Err.Number, Err.Description, False
End Sub

Private Sub ctxtHealthNo_GotFocus()
    On Error Resume Next
    With ctxtHealthNo
        .SelStart = 0
        .SelLength = Len(Trim(ctxtHealthNo.Text))
    End With
End Sub

Private Sub ctxtHealthNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ctxtName_GotFocus()
    On Error Resume Next
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub

Private Sub cvscLetter_Change()
    On Error Resume Next
    '点击滚动条，获得相应的字母。
    If mcolTubeNo.Count > 0 Then
        clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
    End If
End Sub

'功能：清空界面。
Private Sub subClear()
    On Error Resume Next
    ccmbTemplate.Text = ""
    clblLetter.Caption = ""
    cdtpDate.Value = Date
    clblName.Caption = ""
    clblSex.Caption = ""
    clblAge.Caption = ""
    ccmbUnit.Text = ""
    ciptBase.ClearContent
    cfram基本信息.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set mobj体检 = Nothing
    Set mobj体检集 = Nothing
    Set mobj体检表模板 = Nothing
    '关闭相机。
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    
    mblnInUse = False
End Sub


'功能：处理工具栏上按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim i As Integer
    Dim lstr流水号 As String
    On Error GoTo errHandler
    
    Select Case Operate
    Case "清空"
        subClear
        Cancel = True
    Case "修改"
        Dim lstr上一个号 As String
        
        '关闭相机。
        If mblnTakePhoto Then
            cctlCatchPhoto.subDisconnect
        End If
        
        '设置最后录入的系统编号。
        lstr上一个号 = mobj体检.系统编号
        If Not mobj体检.是否已存在 And lstr上一个号 <> "" Then
            lstr上一个号 = mobj体检.func获取系统编号的前一个号(lstr上一个号)
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
        '判断是否需要照相。
        If mblnTakePhoto = True Then
            '判断是否照相
            If cctlCatchPhoto.Photo Is Nothing Then
                Err.Raise 6666, , "没有照相，请重新照相后保存！"
            End If
        End If
        
        '若不是快速录入，检查录入是否有错误。
        If mobj体检.体检表.附加信息.Count > 0 Then
            '修改：2001-9-12（杨春）。
            On Error Resume Next
            ciptBase.Box1(ciptBase.ActiveInputBoxIndex).LostFocus
            On Error GoTo errHandler
            
            If ciptBase.ItemsError.Count > 0 And Not mbln快速录入 Then
                Err.Raise 6666, , "请更正黄色录入框内容！"
            End If
        End If
        MousePointer = 11
        csbMain.Panels(1) = "正在保存体检登记信息，请稍侯..."
        
        '产生试管编号并保存
        With mobj体检
            If .体检表.体检表名 <> ccmbTemplate.Text Then
                .体检表.体检表名 = ccmbTemplate.Text
            End If
            If .体检表.试管编号字母 <> clblLetter.Caption Then
                .体检表.试管编号字母 = clblLetter.Caption
            End If
            .体检人员.姓名 = clblName.Caption
            .体检人员.性别 = clblSex.Caption
            .体检人员.单位名称 = ccmbUnit.Text
            
            If mblnTakePhoto Then
                .体检人员.像片 = cctlCatchPhoto.Photo
            End If
            If Val(clblAge.Caption) > 0 Then
                .体检人员.出生日期 = DateAdd("yyyy", -Val(clblAge.Caption), Date)
            End If
            
            On Error Resume Next
            .体检人员.公民身份号码 = ciptBase.Box1("身份证号").Text
            .体检人员.卫生种类 = ciptBase.Box1("卫生种类").TrueText
            .体检人员.片区 = ciptBase.Box1("片区").TrueText
            .体检人员.行业类别 = ciptBase.Box1("行业类别").TrueText
            If .体检人员.单位申请编号 <> mstr单位申请编号 Then
                .体检人员.单位申请编号 = mstr单位申请编号
            End If
            
            '保存附加信息
            For i = 1 To ciptBase.ItemCount
                'If ciptBase.Box1(i - 1).TrueText <> ciptBase.Box1(i - 1).Text And ciptBase.Box1(i - 1).Text <> "" Then
                If ciptBase.InfoCollection(i).字典名称 <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection(i).名称, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
                Else
                    .体检表.Sub填附加信息值 ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
                End If
            Next i
            
            '设置为年检。
            .体检类别 = P_EXAM_ANNUAL
            .体检日期 = Format(cdtpDate.Value, "yyyy-mm-dd")
        End With
        On Error GoTo errHandler
        
        pobj业务对象.Sub体检登记 mobj体检
        
        csbMain.Panels(2) = "上次保存的体检系统编号：" & mobj体检.系统编号 & "，试管编号：" & mobj体检.试管编号 & "。"
        
        If cchkClear = 1 Then
            subClear
        End If
        
        clblSysNo.Caption = ""
        
        '恢复照相。
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "恢复" Then
                cctlCatchPhoto.sub转换状态
            End If
        End If
        
        '保存按钮不可用。
        ctbMain.Buttons(4).Enabled = False
        
        If coptChoise(0).Value Then
            ctxtName.SetFocus
        Else
            ctxtHealthNo.SetFocus
        End If
        
        '试管字母不能再选择。
        cvscLetter.Enabled = False
        
        Cancel = True
        csbMain.Panels(1) = ""
        MousePointer = 0
     Case "退出"
        '若新增体检记录没有保存，退回系统编号。
        If mobj体检.系统编号 <> "" And Not mobj体检.是否已存在 Then
            If Not sffuncMsg("你确认要退出本界面，并且不保存当前你录入的体检人员登记信息吗？", sf询问) Then
                Cancel = True
                Exit Sub
            End If
            
            '退回系统编号。
            mobj体检.sub退回系统编号 mobj体检.系统编号
        End If
        
        '取消界面通用对象对退出按钮的处理。
        Set mobjGUI.Form = Nothing
        Unload Me
    End Select
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Cancel = True
    Exit Sub
    Resume
    Exit Sub

End Sub


'功能：显示指定系统编号的体检人员的信息在界面上。
Private Sub SubGetPersonInfo(ByVal para系统编号 As String)
    Dim lcolInfo As New Collection
    Dim i As Integer
    Dim j As Integer
    Dim lstrTemp As String
    Dim lstrTubeNo As String
    Dim lstrSysNo As String
    
    
    On Error GoTo errHandler
    MousePointer = 11
    csbMain.Panels(1) = "正在显示当前年检人员的信息，请稍侯..."
    
    '界面暂时不可操作。
    ctbMain.Enabled = False
    cfram基本信息.Enabled = False
    
    '先退回旧系统编号。
    If Not mobj体检.是否已存在 And mobj体检.系统编号 <> "" Then
        mobj体检.sub退回系统编号 mobj体检.系统编号
    End If
    
    '创建旧体检对象。
    Set mobj旧体检 = CreateObject("体检对象部件.clsMedicalExam")
    mobj旧体检.系统编号 = para系统编号
    
    '获得上次年检的体检表
    If ccmbTemplate.Text <> mobj旧体检.体检表.体检表名 Then
        ccmbTemplate.Text = mobj旧体检.体检表.体检表名
    
        '重新初始化录入板。
        On Error Resume Next
        mobjGUI.sub初始化录入板 mobj旧体检.体检表.体检表名
        On Error GoTo errHandler
    End If
    
    '获取旧体检记录的附加信息。
    Set lcolInfo = mobj旧体检.体检表.附加信息
    
    '填写附加信息值
    sub填录入板值 ciptBase, mobjGUI, lcolInfo
    
    '显示基本信息。
    With mobj旧体检.体检人员
        clblName.Caption = .姓名
        clblSex.Caption = .性别
        clblAge.Caption = .年龄
        ccmbUnit.Text = .单位名称
        ccmbUnit_LostFocus
    
        '相片
        '获得并显示照片。
        If Not .像片 Is Nothing Then
            Set cctlCatchPhoto.Photo = .像片
        Else
            cctlCatchPhoto.subClear
        End If
        
        '修改：2001-8-23。
        On Error Resume Next
        mstr单位申请编号 = .单位申请编号
'        If mstr单位申请编号 <> "" Then
'            sub显示单位属性 ciptBase, mstr单位申请编号
'        End If
        On Error GoTo errHandler
    End With
    
    '分配新的系统编号
    lstrSysNo = mobj体检.Func分配系统编号
    mobj体检.系统编号 = lstrSysNo
    clblSysNo.Caption = lstrSysNo
    
    '健康档案不变。
    mobj体检.体检人员.健康档案编号 = mobj旧体检.体检人员.健康档案编号
    
    
    '设置年检的体检表名，从而获取新试管编号。
    mobj体检.体检表.体检表名 = ccmbTemplate.Text
    
    '获取本体检表的当天已使用的试管编号字母。
    clblLetter.Caption = mobj体检.体检表.试管编号字母
    If clblLetter.Caption = "" Then
        
        '该次体检登记是当天的第一个，从体检表模板对象中获取所有可选的字幕。
        mobj体检表模板.体检表名 = ccmbTemplate.Text
        lstrTubeNo = mobj体检表模板.试管字母编号
        
        '将字母按逗号分开，加入mcoltubeNo中。
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
            '字母可以选择。
            cvscLetter.Enabled = True
            cvscLetter.Min = 1
            cvscLetter.Max = mcolTubeNo.Count
            cvscLetter.Value = 1
        Else
            ctbMain.Buttons(4).Enabled = False
            '提示该体检表无可用的字母。
            Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
        End If
    Else
        '有字母，不能选择字母。
        cvscLetter.Enabled = False
    End If
    
    '录入区可以操作。
    cfram基本信息.Enabled = True
    
    '保存按钮可用。
    ctbMain.Buttons(4).Enabled = True
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "SubGetPersonInfo", Err.Number, Err.Description, True
    End If
    
    '恢复界面可操作。
    ctbMain.Enabled = True
    cframSearch.Enabled = True
    MousePointer = 0
    csbMain.Panels(1) = ""
    
    Exit Sub
    Resume
End Sub
    
Private Function func根据健康档案编号获取系统按号(ByVal para健康档案编号 As String) As String
    Dim lobj体检人员  As Object 'clsPersonExamed.
    Dim lobj体检 As Object      'clsMedicalExam
    Dim lstr系统编号 As String
    
    On Error GoTo errHandler
    
    '获取该人最近一次体检记录。
    '创建体检人员对象。
    Set lobj体检人员 = CreateObject("体检对象部件.clsPersonExamed")
    lobj体检人员.健康档案编号 = para健康档案编号
    Set lobj体检 = lobj体检人员.Func获取本人最近一次体检
    If Not lobj体检 Is Nothing Then
        lstr系统编号 = lobj体检.系统编号
    Else
        Err.Raise 6666, , "该体检人员还没有在本体检中心体检过，无法进行年检登记。请选择初检登记。"
    End If
            
    func根据健康档案编号获取系统按号 = lstr系统编号
    
    Exit Function
errHandler:
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "func根据健康档案编号获取系统按号", Err.Number, Err.Description, True
    Exit Function
    Resume

End Function
