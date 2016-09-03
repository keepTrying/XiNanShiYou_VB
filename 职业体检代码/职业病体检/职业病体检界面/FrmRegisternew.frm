VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "体检登记"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   13590
   ClipControls    =   0   'False
   Icon            =   "FrmRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9837.104
   ScaleMode       =   0  'User
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox Check身份证 
      Caption         =   "刷二代身份证"
      Height          =   255
      Left            =   8520
      TabIndex        =   38
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "保存后清空"
      Height          =   345
      Left            =   8520
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   6600
      Top             =   360
   End
   Begin VB.Frame cfram基本信息 
      Caption         =   "登记基本信息（非快速录入时黄色为必录项，快速录入时只需照相):"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   8055
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   12540
      Begin VB.TextBox ctxt民族 
         Height          =   300
         Left            =   4800
         TabIndex        =   51
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox ctxt文化程度 
         Height          =   300
         Left            =   3480
         TabIndex        =   49
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox Ctxt婚否 
         Height          =   300
         Left            =   2520
         TabIndex        =   47
         Text            =   "Combo1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox ctxt身份证号 
         Height          =   300
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   2130
      End
      Begin VB.ComboBox ccmb体检时期 
         Height          =   300
         Left            =   8160
         TabIndex        =   43
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox ccmb体检类别 
         Height          =   300
         Left            =   2640
         TabIndex        =   40
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox clblsysno 
         Height          =   270
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox cchk录入单位名称 
         Caption         =   "录入单位名称"
         Height          =   255
         Left            =   10800
         TabIndex        =   35
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox ctxt份数 
         Height          =   270
         Left            =   6600
         TabIndex        =   33
         Text            =   "1"
         Top             =   7320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox ctxt体检单号 
         Height          =   315
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox ctxtTubeNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6000
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox ccmb体检类型 
         Height          =   300
         ItemData        =   "FrmRegister.frx":0442
         Left            =   3960
         List            =   "FrmRegister.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   7440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox ctxtAge 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3480
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox ccmbSex 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         ItemData        =   "FrmRegister.frx":045C
         Left            =   2520
         List            =   "FrmRegister.frx":0466
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   840
      End
      Begin VB.TextBox ctxtName 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2130
      End
      Begin VB.Frame frmPhoto 
         Caption         =   "照像："
         ClipControls    =   0   'False
         ForeColor       =   &H00800000&
         Height          =   4275
         Left            =   7560
         TabIndex        =   26
         Top             =   2880
         Width           =   4660
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Height          =   2175
            Left            =   1560
            ScaleHeight     =   2115
            ScaleWidth      =   1635
            TabIndex        =   39
            Top             =   720
            Width           =   1695
         End
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3570
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   6297
            BackColor       =   0
            FontSize        =   9.75
            OriginalSize    =   -1  'True
         End
      End
      Begin VB.CommandButton ccmd单位定位 
         Caption         =   "定位(&T)"
         Height          =   375
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   7800
         TabIndex        =   7
         Top             =   1320
         Width           =   3240
      End
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6360
         TabIndex        =   11
         Top             =   1320
         Width           =   345
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   4440
         TabIndex        =   0
         Top             =   720
         Width           =   3480
      End
      Begin 录入控件.ctlInputDictGrid c字典表 
         Height          =   3255
         Left            =   4440
         TabIndex        =   25
         Top             =   3600
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
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
         Left            =   10080
         TabIndex        =   2
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   68091904
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin 录入控件.ctlInputFrame ciptBase 
         Height          =   975
         Left            =   1440
         TabIndex        =   10
         Top             =   6120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1720
         BackColor       =   15791081
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
         Caption         =   ""
         Rows            =   1
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "民族："
         Height          =   180
         Left            =   4800
         TabIndex        =   50
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "文化程度："
         Height          =   180
         Left            =   3480
         TabIndex        =   48
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "婚否："
         Height          =   180
         Left            =   2520
         TabIndex        =   46
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "身份证号："
         Height          =   180
         Left            =   240
         TabIndex        =   44
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "体检时期："
         Height          =   180
         Left            =   8160
         TabIndex        =   42
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "体检人员类别："
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "注：刷条码前请确保文本框中内容为空"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   3060
      End
      Begin VB.Label clbl份数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "份数："
         Height          =   180
         Left            =   6480
         TabIndex        =   32
         Top             =   7080
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检单号："
         Height          =   180
         Index           =   7
         Left            =   4200
         TabIndex        =   30
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label clbl旧体检日期 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   29
         Top             =   6720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上次体检日期："
         Height          =   180
         Index           =   4
         Left            =   2280
         TabIndex        =   28
         Top             =   6480
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检类型："
         Height          =   180
         Index           =   3
         Left            =   4080
         TabIndex        =   27
         Top             =   7200
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Left            =   2520
         TabIndex        =   24
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表："
         Height          =   180
         Left            =   4440
         TabIndex        =   22
         Top             =   480
         Width           =   720
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "保存后请看状态栏"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6600
         TabIndex        =   21
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6000
         TabIndex        =   20
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管编号："
         Height          =   180
         Index           =   1
         Left            =   6000
         TabIndex        =   18
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
         Height          =   180
         Index           =   2
         Left            =   10080
         TabIndex        =   17
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   5
         Left            =   7800
         TabIndex        =   16
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   15
         Top             =   1080
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   1680
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar cstbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   8985
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23918
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
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：邓恒
'最后修改：杨春
Public pstr系统编号 As String

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

'新选择的体检项目、收费项目
Private mcol体检项目 As New Collection
Private mcol收费项目 As New Collection               'item:编号,key：编号。

Public pstr系统编号名称 As String '修改：2002-10-10（杨春）为嘉定定制增加该属性。

Private mobj记忆  As cls用户操作记忆
Private mstr默认年龄 As String


'功能：记载当前窗体是否已加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchkClear_Click()
    On Error Resume Next
    ctxtName.SetFocus
End Sub

Private Sub cchk录入单位名称_Click()
    Dim lblnVisible As Boolean
    On Error Resume Next
    If cchk录入单位名称.Value = 1 Then
        lblnVisible = True
    Else
        lblnVisible = False
    End If
    ccmbUnit.Visible = lblnVisible
    ccmd单位定位.Visible = lblnVisible
    Label2(5).Visible = lblnVisible
    ctxtName.SetFocus
End Sub

Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex = "" And ccmbSex.ListCount > 0 Then
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
            ctxt体检单号.SetFocus
        End If
    End If
End Sub

'功能：控制不能输入体检表名称，只能选择。
'创建：2002-11-28（杨春）。
Private Sub ccmbTemplate_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
        KeyAscii = 0
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
        If ctxt份数.Visible Then
            ctxt份数.SetFocus
        Else
            If ciptBase.Visible Then
                ciptBase.SetFocus
            End If
        End If
    Else
        mstr单位申请编号 = ""
    End If
        
End Sub

Private Sub ccmb体检类别_click()
    Dim lobj体检表模板集 As Object
    Dim lcolInfo As Collection
    Dim i As Integer
     '将体检类别加入组合框中
    'Set lobj体检类别 = CreateObject("体检对象.clsmedicalexamtemplateset")
    'lobj体检类别.体检表类别 = 1
    'Set lcol类别 = lobj体检类别.体检类别
    'ccmb体检类别.AddItem ""
    'For i = 1 To lcol类别.recordCount
    '    ccmb体检类别.AddItem lcol类别("类别")
    '    ccmb体检类别.ItemData(ccmb体检类别.NewIndex) = lcol类别("编号")
    '    lcol类别.movenext
    'Next
    'ccmb体检类别.ListIndex = 0
    'Set lobj体检类别 = Nothing
   
    
    '将所有的非复查体检表模板加入到体检表下拉列表框中。再加体检类别条件
    ccmbTemplate.Clear
    Set lobj体检表模板集 = CreateObject("体检对象.ClsMedicalExamTemplateSet")
    lobj体检表模板集.体检表类型 = 3
    lobj体检表模板集.体检表类别 = ccmb体检类别.ItemData(ccmb体检类别.ListIndex)
    Set lcolInfo = lobj体检表模板集.元素集
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
    Next
    ccmbTemplate.Text = ccmbTemplate.List(0)
    
    Set lobj体检表模板集 = Nothing
    Call ccmbTemplate_Click
    
    
End Sub

Private Sub ccmb体检类型_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    If KeyCode = 13 Then
        ctxt份数.SetFocus
    End If
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub


Private Sub Check身份证_Click()
    If Check身份证.Value = 0 Then
        cctlCatchPhoto.Visible = True
        Picture1.Visible = False
        cctlCatchPhoto.funcInitVideo
        'Check身份证.Enabled = True
        cctlCatchPhoto.Enabled = True
    Else
        cctlCatchPhoto.Visible = False
        Picture1.Visible = True
    End If
End Sub



Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '自动保存。
    If ctbMain.Buttons(6).Enabled Then
        ctxtName.SetFocus
        SendKeys "{F2}"
    End If
End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c字典表" Then
        c字典表.Visible = False
    End If

End Sub


Private Sub clblsysno_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub clblsysno_LostFocus()
    If Len(clblSysNo.Text) < 5 Then
            MsgBox "系统编号错误，请检查！", vbInformation, "系统提示"
            Exit Sub
    End If
    mobj体检.系统编号 = Trim(clblSysNo.Text)
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt体检单号.SetFocus
    End If

End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub

Private Sub ctxtTubeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        cdtpDate.SetFocus

    End If
End Sub



Private Sub ctxt份数_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '若录入板没有录入项目，则直接保存。
        If ciptBase.Visible Then
            ciptBase.SetFocus
            ciptBase.ItemSetFocus 0
        End If
    End If
End Sub

Private Sub ctxt体检单号_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ccmbUnit.Visible Then
            ccmbUnit.SetFocus
        Else
            If ctxt份数.Visible Then
                ctxt份数.SetFocus
            Else
                If ciptBase.Visible Then
                    ciptBase.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnTakePhoto Then
        '重新初始化照相控件。
        cctlCatchPhoto.funcInitVideo
    End If
    ctxtName.SetFocus
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    gfsubHideComboList ccmbUnit
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
   
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    MousePointer = 11
    
    '界面不可操作。
'    cfram基本信息.Enabled = False
    ctbMain.Enabled = False
    
    Set mcol收费项目 = New Collection
    Set mcol体检项目 = New Collection
    
    Set mobj旧体检 = CreateObject("体检对象.clsMedicalExam")
    
    Set mobj体检 = CreateObject("体检对象.clsMedicalExam")
    '修改：2002-10-10（设置系统编号名称）。
    If pstr系统编号名称 <> "" Then
        mobj体检.系统编号名称 = pstr系统编号名称
    End If
    
    Set mobj体检集 = CreateObject("体检对象.clsMedicalExamSet")
    Set mobj体检表模板 = CreateObject("体检对象.ClsMedicalExamTemplate")
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    mobjGUI.pbln自动设置字典高度 = False
    
    '设置工具栏上所需要的各种按钮。
    Dim lcol工具栏按钮 As New Collection           '工具栏上的按钮初始化集合。
    With lcol工具栏按钮
        .Add "清空"
        .Add "|"
        .Add "体检项目(&T)102"
        .Add "载入照片(&E)103"
        .Add "|"
        .Add "保存"
        .Add "修改"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
        Set .c录入板 = ciptBase
        Set .c字典表 = c字典表
        Set .c状态栏 = cstbMain
        
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcol工具栏按钮, ""
    End With
    
    '清空
    subClear
    cdtpDate.Value = Date
    '新分配系统编号
    'clblsysno.Caption = mobj体检.Func分配系统编号
    clblSysNo.Text = ""
    'mobj体检.系统编号 = Trim(clblsysno.Text)
    Check身份证.Value = 1
    'cctlCatchPhoto.Visible = False
    'cctlCatchPhoto.Visible = True
    
    If pobj业务对象.业务设置("试管编号自动生成") = "否" Then
        ctxtTubeNo.Visible = True
        ctxtTubeNo.TabIndex = 1
        clblTubeNo.Visible = False
        clblLetter.Visible = False
        cvscLetter.Visible = False
    Else
        ctxtTubeNo.Visible = False
        clblTubeNo.Visible = True
        clblLetter.Visible = True
        cvscLetter.Visible = True
    End If
    
    DoEvents
    ccmb体检类型.Visible = True
    Label2(3).Visible = True
    clbl旧体检日期.Visible = True

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
    cstbMain.Panels(1) = lstrError
End Sub

'功能：完成form_load余下的初始化工作。
Private Sub Timer1_Timer()
    Dim lobj体检表模板集 As Object  '体检表模板集，获取所有的非复查体检表模板名称。
    Dim lcolInfo As Collection
    Dim lcol类别 As Object
    Dim i As Integer
    Dim lobj体检类别 As Object
    On Error GoTo errHandler
    
    '定时器不再起作用。
    Timer1.Enabled = False
    
    '从当日工作及已簿中获取当天录入过的单位名称。
    Set lcolInfo = pobj业务对象.当日工作记忆簿.单位名称集
    For i = 1 To lcolInfo.Count
        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    
    '将体检类别加入组合框中
    Set lobj体检类别 = CreateObject("体检对象.clsmedicalexamtemplateset")
    lobj体检类别.体检表类别 = 1
    Set lcol类别 = lobj体检类别.体检类别
    'ccmb体检类别.AddItem ""
    'ccmb体检类别.ListIndex = 0
    For i = 1 To lcol类别.recordcount
        ccmb体检类别.AddItem lcol类别("类别")
        ccmb体检类别.ItemData(ccmb体检类别.NewIndex) = lcol类别("编号")
        lcol类别.movenext
    Next
    ccmb体检类别.ListIndex = 0
    ccmb体检类别.Text = ccmb体检类别.List(0)
    'ccmb体检类别.ListIndex = 0
    Set lobj体检类别 = Nothing
   
    
    '将所有的非复查体检表模板加入到体检表下拉列表框中。
    'Set lobj体检表模板集 = CreateObject("体检对象.ClsMedicalExamTemplateSet")
    'lobj体检表模板集.体检表类型 = 3
    
    'lobj体检表模板集.体检表类别 = ccmb体检类别.ItemData(ccmb体检类别.ListIndex)
    'Set lcolInfo = lobj体检表模板集.元素集
    'For i = 1 To lcolInfo.Count
    '    ccmbTemplate.AddItem lcolInfo(i)
    'Next
    'ccmbTemplate.Text = ccmbTemplate.List(0)
    'Set lobj体检表模板集 = Nothing
    
    
    
    
    '根据业务设置判断是否照相。
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
    
    '只有初检，而且快速登记才可以批量登记。
    If Not mbln快速录入 Or pstr系统编号 <> "" Then
        clbl份数.Visible = False
        ctxt份数.Visible = False
    End If
    
    ccmb体检类型.ListIndex = 0
    
    If ccmbTemplate.ListCount > 0 Then
        'ccmbTemplate.ListIndex = 0
        ccmbTemplate.Text = ccmbTemplate.List(0)
        subChangeTemplate
        
    End If
    
    '需要照相时初始化照相控件。
    If mblnTakePhoto And Check身份证 = False Then
        '初始化控件。
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pstr系统编号 <> "" Then
        '是年检登记。
        '显示体检人员基本信息。
        SubGetPersonInfo pstr系统编号
    End If
    
    On Error Resume Next
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = "*"
    mobj记忆.业务名 = "体检管理"
    mstr默认年龄 = mobj记忆.记忆项值("体检年龄")
'    If mstr默认年龄 <> "" And ctxtAge = "" Then
'        ctxtAge = mstr默认年龄
'    End If
    
    If mobj记忆.记忆项值("体检登记时录入单位名称") = "" Or mobj记忆.记忆项值("体检登记时录入单位名称") = "是" Then
        cchk录入单位名称.Value = 1
    Else
        cchk录入单位名称.Value = 0
    End If
    cfram基本信息.Enabled = True
    ctbMain.Enabled = True
    cctlCatchPhoto.Visible = False
    If Check身份证.Value = 0 Then
        'frmPhoto.Visible = True
        cctlCatchPhoto.Visible = True
        Picture1.Visible = False
    Else
        cctlCatchPhoto.Visible = False
        'frmPhoto.Visible = False
        Picture1.Visible = True
    End If
    MousePointer = 0
    Exit Sub
    
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "Timer1_Timer", 6666, lstrError, False
    
    '恢复界面可操作。
    cfram基本信息.Enabled = True
    ctbMain.Enabled = True
    MousePointer = 0

End Sub


Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    '选择体检表
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    MousePointer = 11
    
    subChangeTemplate
    
    ctbMain.Buttons(6).Enabled = True
'    If ctxtTubeNo.Visible Then
'        ctxtTubeNo.SetFocus
'    Else
'        ctxtName.SetFocus
'    End If
    MousePointer = 0
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "ccmbTemplate_Click", 6666, lstrError, False
    
    Exit Sub
    Resume
End Sub

Private Sub subChangeTemplate()
    On Error GoTo errHandler
    '选择体检表
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    '获取新试管编号。
    If mobj体检.体检表.体检表名 <> ccmbTemplate.Text Then
        mobj体检.体检表.体检表名 = ccmbTemplate.Text

        '根据体检表模板获取该体检表所有可用的字母。
        mobj体检表模板.体检表名 = ccmbTemplate.Text

        If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
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
                    ctbMain.Buttons(6).Enabled = False
                    '提示该体检表无可用的字母。
                    Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
                End If
            Else
                '有字母，不能选择字母。
                clblLetter.Caption = mobj体检.体检表.试管编号字母
                cvscLetter.Enabled = False
            End If
        Else
            clblLetter.Caption = mobj体检表模板.试管字母编号
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

        '修改：2002-7-26（杨春）根据“是否年检表”选择体检表类型。
        If mobj体检表模板.是否年检表 Then
            ccmb体检类型.ListIndex = 1
        Else
            ccmb体检类型.ListIndex = 0
        End If

        '修改：2002-10-10（杨春）嘉定定制：显示体检金额。
        On Error Resume Next
        ciptBase.Box1("体检金额").Text = mobj体检表模板.收费标准金额
'

    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "subChangeTemplate", 6666, lstrError, True
    
    Exit Sub
    Resume
End Sub

'自动弹出列表框
Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
'    gfsubShowComboList ccmbUnit
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
        If mstr单位申请编号 <> pobj业务对象.当日工作记忆簿.单位编号(ccmbUnit.Text) And mstr单位申请编号 <> "" Then
            pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
        End If
    End If
    Exit Sub
errHandler:
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
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
    
    '把焦点回到单位录入框。保存能保存新单位定位信息。
    ccmbUnit.SetFocus
    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "ccmd单位定位_Click", 6666, lstrError, False
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
                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
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
            If Val(内容) >= Val(ctxtAge.Text) Then
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
    'cdtpDate.Value = Date
    ctxtName.Text = ""
    'ctxtAge = mstr默认年龄
    ctxtAge = ""

    ccmbUnit.Text = ""
    
    ctxtTubeNo = ""
    ctxt体检单号 = ""
    
    '修改：2002-10-10（杨春）嘉定定制：体检金额不清空。
    Dim ldbl体检金额 As Double
    ldbl体检金额 = ciptBase.Box1("体检金额").Text
    ciptBase.ClearContent
    ciptBase.Box1("体检金额").Text = ldbl体检金额
    
    clbl旧体检日期.Caption = ""
    Label2(4).Visible = False
    clbl旧体检日期.Visible = False
    Set cctlCatchPhoto.Photo = Nothing
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '若新增体检记录没有保存，退回系统编号。
    If Not mobj体检 Is Nothing Then
        If mobj体检.系统编号 <> "" And Not mobj体检.是否已存在 Then
            '退回系统编号。
            mobj体检.sub退回系统编号 mobj体检.系统编号
        End If
    End If
    mobj记忆.sub覆盖记忆值 "体检登记时录入单位名称", IIf(cchk录入单位名称.Value = 1, "是", "否")
     
    Set mobj体检 = Nothing
    Set mobj体检集 = Nothing
    Set mobj体检表模板 = Nothing
    '关闭相机。
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    mblnInUse = False
    pstr系统编号名称 = ""
End Sub


'功能：处理工具栏上按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    Dim lstr流水号 As String
    Dim lstr系统编号 As String
    Dim lcol原体检项目 As Collection
    
    On Error GoTo errHandler
    
    Select Case Operate
    
    Case "清空"
        subClear
        '清空健康档案编号，表示新录入的体检人员。
        mobj体检.体检人员.健康档案编号 = ""
        
        Cancel = True
    
    Case "保存"
        '如果系统编号长度小于5，初步判定是操作失误
        If Len(clblSysNo.Text) < 5 Then
            MsgBox "系统编号错误，请检查！", vbInformation, "系统提示"
            Exit Sub
        End If
        '判断是否需要照相。
        If mblnTakePhoto = True Then
            '判断是否照相
            If cctlCatchPhoto.Photo Is Nothing Then
                Err.Raise 6666, , "你在“业务设置”中设置了要照像，但是现在你没有照相，无法保存。解决办法：" & Chr(13) & Chr(10) & "（1） 请按“取像”按钮照相后保存！" & Chr(13) & Chr(10) & "（2）若你不准备照相，请先进入“业务设置”设置不照相。"
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
        
        '产生试管编号并保存
        With mobj体检
            If .体检表.体检表名 <> ccmbTemplate.Text Then
                .体检表.体检表名 = ccmbTemplate.Text
            End If
            '修改：2004-1-9（试管编号可以输入）
            If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
                If .体检表.试管编号字母 <> clblLetter.Caption Then
                    .体检表.试管编号字母 = clblLetter.Caption
                End If
            Else
                .体检表.试管编号字母 = clblLetter.Caption
                .试管编号 = ctxtTubeNo.Text
            End If
            
            .体检人员.姓名 = ctxtName
            .体检人员.性别 = ccmbSex.Text
            .体检人员.单位名称 = ccmbUnit.Text
            
            If mblnTakePhoto Then
                .体检人员.像片 = cctlCatchPhoto.Photo
'                .体检人员.像片压缩 = cctlCatchPhoto.Photo
            End If
            If Val(ctxtAge.Text) > 0 Then
'                If Val(ctxtAge.Text) > 200 Then
'                    Err.Raise 6666, , "年龄超过系统允许的最大数：200。"
'                End If
                .体检人员.出生日期 = DateAdd("yyyy", -Val(ctxtAge.Text), Date)
            Else
                '如果输入字符，则记忆该年龄。
                mobj记忆.sub覆盖记忆值 "体检年龄", ctxtAge.Text
                mstr默认年龄 = ctxtAge.Text
            End If
            .体检人员.年龄 = ctxtAge.Text
            
            On Error Resume Next
            .体检人员.公民身份号码 = ciptBase.Box1("身份证号").Text
            .体检人员.卫生种类 = ciptBase.Box1("卫生种类").TrueText
            .体检人员.行业类别 = ciptBase.Box1("行业类别").TrueText
            .体检人员.片区 = ciptBase.Box1("片区").TrueText
            .体检人员.文化程度 = ctxt文化程度.Text
            .体检人员.民族 = ctxt民族.Text
            If ccmbUnit.Text = "" Then
                .体检人员.单位申请编号 = ""
            Else
                If .体检人员.单位申请编号 <> mstr单位申请编号 Then
                    '给单位编号重新赋值，可以重新获取其卫生种类、行业类别、片区。
                    .体检人员.单位申请编号 = mstr单位申请编号
                End If
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
            
            '设置为体检类别。
            If ccmb体检类型.Text = "初检" Then
                .体检类别 = P_EXAM_FIRST
            Else
                .体检类别 = P_EXAM_ANNUAL
            End If
            .体检日期 = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            '修改：2004-1-9（增加体检单号）
            .体检单号 = ctxt体检单号.Text

        End With
        
        On Error GoTo errHandler
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
        
        If mcol收费项目.Count > 0 Then
            pobj业务对象.Sub体检登记 mobj体检, , , mcol收费项目, Val(ctxt份数)
        Else
            pobj业务对象.Sub体检登记 mobj体检, , , , Val(ctxt份数)
        End If
        cstbMain.Panels(1) = "上次保存的体检系统编号：" & mobj体检.系统编号 & "，试管编号：" & mobj体检.试管编号
        If mobj体检.收费批号 <> "" Then
            cstbMain.Panels(1) = cstbMain.Panels(1) & "，收费批号：" & mobj体检.收费批号
        End If
        
        
        If cchkClear = 1 Then
            subClear
        End If
        
        '生成新的系统编号。
        'clblsysno.Caption = mobj体检.Func分配系统编号
        clblSysNo.Text = ""
        mobj体检.系统编号 = Trim(clblSysNo.Text)
        mobj体检.体检表.体检表名 = ccmbTemplate.Text
        
        Set mcol体检项目 = New Collection
        Set mcol收费项目 = New Collection
        '恢复照相。
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "恢复" Then
                cctlCatchPhoto.sub转换状态
            End If
            
        End If
        
        '试管字母不能再选择。
        cvscLetter.Enabled = False
        ctxtName.SetFocus

        frmRegisterManage.sub查询并显示
        
        Cancel = True
        MousePointer = 0
    
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
    Case "载入照片"
        Dim lstrFile As String
        ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片", vbDirectory) <> "" Then
            ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片"
        End If
        ccmdFile.FileName = Trim(clblSysNo.Text)
        ccmdFile.ShowOpen
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            If InStr(lstrFile, ".") > 0 Then
                Set cctlCatchPhoto.Photo = LoadPicture(lstrFile)
                mblnTakePhoto = True
            End If
        End If
    Case "修改"
        Dim lobjRec As Object
        '获取当天最后的号。
        If Val(Right(Trim(clblSysNo.Text), Len(Trim(clblSysNo.Text)) - Len(mobj体检.系统编号固定部分))) > 1 Then
            FrmEditRegister.系统编号 = mobj体检.系统编号固定部分 & Format(Val(Right(Trim(clblSysNo.Text), Len(Trim(clblSysNo.Text)) - Len(mobj体检.系统编号固定部分))) - 1, String(Len(Trim(clblSysNo.Text)) - Len(mobj体检.系统编号固定部分), "0"))
        Else
            FrmEditRegister.系统编号 = ""
        End If
        FrmEditRegister.Show 1, Me
    End Select
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
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
    
    '界面暂时不可操作。
    ctbMain.Enabled = False
    
    '先退回旧系统编号。
    If Not mobj体检.是否已存在 And mobj体检.系统编号 <> "" Then
        mobj体检.sub退回系统编号 mobj体检.系统编号
    End If
    
    '创建旧体检对象。
    Set mobj旧体检 = CreateObject("体检对象.clsMedicalExam")
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
        ctxtName.Text = .姓名
        ccmbSex.Text = .性别
        ctxtAge.Text = .年龄
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
        
        On Error GoTo errHandler
    End With
    
    '修改：2001-12-30（显示上次体检日期）。
    Label2(4).Visible = True
    clbl旧体检日期.Visible = True
    clbl旧体检日期.Caption = mobj旧体检.体检日期
    
    '修改：2002-1-6（若时间间隔超过18个月，自动设置为初检）。
    If IsDate(clbl旧体检日期.Caption) Then
        If DateDiff("m", clbl旧体检日期.Caption, Now) >= 18 Then
            ccmb体检类型.ListIndex = 0
        Else
            '不到18个月，自动设置为年检。
            ccmb体检类型.ListIndex = 1
        End If
    End If
    '分配新的系统编号
    lstrSysNo = mobj体检.Func分配系统编号
    mobj体检.系统编号 = lstrSysNo
    clblSysNo.Text = lstrSysNo
    
    '健康档案不变。
    mobj体检.体检人员.健康档案编号 = mobj旧体检.体检人员.健康档案编号
    
    
    '设置年检的体检表名，从而获取新试管编号。
    mobj体检.体检表.体检表名 = ccmbTemplate.Text
    
    If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
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
                ctbMain.Buttons(6).Enabled = False
                '提示该体检表无可用的字母。
                Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
            End If
        Else
            '有字母，不能选择字母。
            cvscLetter.Enabled = False
        End If
    Else
        ctxtTubeNo = mobj体检.试管编号
    End If
    '保存按钮可用。
    ctbMain.Buttons(6).Enabled = True
    Err.Clear
    
errHandler:
    '恢复界面可操作。
    ctbMain.Enabled = True
    MousePointer = 0
    If Err <> 0 Then
        sfsub错误处理 "体检界面部件", "FrmRegisterAnnual", "SubGetPersonInfo", Err.Number, Err.Description, True
    End If
    
    Exit Sub
    Resume
End Sub
    
