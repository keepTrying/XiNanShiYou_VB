VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetMedicalExamTemplate 
   Caption         =   "体检表设置"
   ClientHeight    =   7980
   ClientLeft      =   615
   ClientTop       =   1260
   ClientWidth     =   10410
   ClipControls    =   0   'False
   Icon            =   "frmSetMedicalExamTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   7980
   ScaleWidth      =   10410
   Begin VB.Frame cfraMedicalTemplateName 
      Appearance      =   0  'Flat
      Caption         =   "已有体检表(选中可以修改)"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   2955
      Begin VB.ListBox clstTemplate 
         BackColor       =   &H00FFFFFF&
         Height          =   6540
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "新增体检表"
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   3000
      TabIndex        =   10
      Top             =   960
      Width           =   9105
      Begin VB.ComboBox tijian_human_leixing 
         Height          =   300
         Left            =   2640
         TabIndex        =   37
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox tijian_leibie 
         Height          =   300
         Left            =   5280
         TabIndex        =   47
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox cchkAnnual 
         Caption         =   "是否年检表"
         Height          =   180
         Left            =   3480
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox ctxtLetter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox ctxtName 
         Height          =   300
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "汉字"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox cchkAgain 
         Caption         =   "是否复查体检表"
         Height          =   180
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox ctxtNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox ccmbSheet 
         Height          =   300
         Left            =   4920
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin TabDlg.SSTab ctabMain 
         Height          =   5430
         Left            =   60
         TabIndex        =   5
         Top             =   1560
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   9578
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "  体检项目  "
         TabPicture(0)   =   "frmSetMedicalExamTemplate.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "  体检结论  "
         TabPicture(1)   =   "frmSetMedicalExamTemplate.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame1(3)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "  体检表附加信息  "
         TabPicture(2)   =   "frmSetMedicalExamTemplate.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "  收费标准和诊断处理意见  "
         TabPicture(3)   =   "frmSetMedicalExamTemplate.frx":0496
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1(1)"
         Tab(3).ControlCount=   1
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   5055
            Index           =   0
            Left            =   -75000
            TabIndex        =   32
            Top             =   310
            Width           =   7890
            Begin MSComctlLib.TreeView ctrwSelectedItem 
               Height          =   4365
               Left            =   120
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   7699
               _Version        =   393217
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Appearance      =   1
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
            Begin VB.CommandButton ccmdDeleteItem 
               Caption         =   ">>"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3435
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   1080
               Width           =   650
            End
            Begin VB.CommandButton ccmdAddItem 
               Caption         =   "<<"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3435
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   600
               Width           =   650
            End
            Begin MSComctlLib.TreeView ctrwAllItem 
               Height          =   4365
               Left            =   4200
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   480
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   7699
               _Version        =   393217
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Appearance      =   1
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "可选择的体检项目（双击可以加入）"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   4200
               TabIndex        =   36
               Top             =   240
               Width           =   2880
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "选定的体检项目（双击可以去掉）"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   2700
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   4905
            Index           =   3
            Left            =   0
            TabIndex        =   27
            Top             =   310
            Width           =   8865
            Begin MSComctlLib.TreeView ctrwAllConclusion 
               Height          =   4215
               Left            =   4320
               TabIndex        =   40
               Top             =   480
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   7435
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               FullRowSelect   =   -1  'True
               Appearance      =   1
            End
            Begin VB.CommandButton ccmdDeleteConclusion 
               Caption         =   ">>"
               Height          =   350
               Left            =   3600
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   1320
               Width           =   650
            End
            Begin VB.CommandButton ccmdAddConclusion 
               Caption         =   "<<"
               Height          =   350
               Left            =   3600
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   780
               Width           =   650
            End
            Begin MSComctlLib.TreeView ctrwSelectedConclusion 
               Height          =   4215
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   7435
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               FullRowSelect   =   -1  'True
               Appearance      =   1
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "所有可选的体检结论"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   11
               Left            =   4560
               TabIndex        =   31
               Top             =   240
               Width           =   1620
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "选定的结论"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   10
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   900
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   5055
            Index           =   2
            Left            =   -75000
            TabIndex        =   18
            Top             =   310
            Width           =   7905
            Begin VB.CommandButton ccmdDown 
               Caption         =   "下移"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   2280
               Width           =   650
            End
            Begin VB.CommandButton ccmdUp 
               Caption         =   "上移"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1800
               Width           =   650
            End
            Begin VB.ListBox clstSelectedBase 
               Height          =   4050
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   22
               Top             =   480
               Width           =   3135
            End
            Begin VB.ListBox clstAllBase 
               BackColor       =   &H00FFFFFF&
               Height          =   3840
               Left            =   4080
               TabIndex        =   21
               Top             =   480
               Width           =   3435
            End
            Begin VB.CommandButton ccmdAddBase 
               Caption         =   "<<"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   780
               Width           =   650
            End
            Begin VB.CommandButton ccmdDeleteBase 
               Caption         =   ">>"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   1200
               Width           =   650
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "选定的应填附加项目（必录项前打钩）"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   6
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   3060
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "可选的附加项目"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   7
               Left            =   4080
               TabIndex        =   25
               Top             =   240
               Width           =   1260
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   5025
            Index           =   1
            Left            =   -75000
            TabIndex        =   15
            Top             =   310
            Width           =   7665
            Begin VB.ListBox clstDisposalIdea 
               Height          =   4260
               Left            =   3720
               Style           =   1  'Checkbox
               TabIndex        =   44
               Top             =   480
               Width           =   3255
            End
            Begin VB.TextBox ctxtCharge 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   480
               Width           =   3195
            End
            Begin VB.ListBox clstCharge 
               Height          =   4020
               Left            =   120
               TabIndex        =   16
               Top             =   840
               Width           =   3195
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "选择诊断处理意见"
               Height          =   180
               Index           =   0
               Left            =   3780
               TabIndex        =   45
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "选择收费标准"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   1080
            End
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检人员类型"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   2640
         TabIndex        =   48
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检类别"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   5280
         TabIndex        =   46
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管字母："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   6120
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检单名称："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3720
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表名称："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表代号："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   6600
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   2520
      Top             =   480
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   9
      Top             =   7590
      Visible         =   0   'False
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17833
            Key             =   "Msg"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检类别"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmSetMedicalExamTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：刘浩

Private mobj体检表模板 As Object               '当前正在操作的体检表模板。
Private WithEvents mobjGUI  As cls界面通用对象 '界面通用对象，用来初始化工具栏。
Attribute mobjGUI.VB_VarHelpID = -1

Private mblnInUse As Boolean                   '对应属性pblnInUse。
Private mblnSys As Boolean

'功能：表明当前窗体是否已家载。
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchkAgain_Click()
    On Error Resume Next
    If cchkAgain.Value = 1 Then
        cchkAnnual.Value = 0
        cchkAnnual.Enabled = False
    Else
        cchkAnnual.Enabled = True
    End If
    
End Sub

Private Sub cchkAgain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        clstDisposalIdea.SetFocus
    End If
End Sub

Private Sub ccmbSheet_GotFocus()
    On Error Resume Next
    If ccmbSheet.Text = "" And ccmbSheet.ListCount > 0 Then
        ccmbSheet.ListIndex = 0
    End If
End Sub

Private Sub clstCharge_Click()
    Dim lstrTemp As String
    
    On Error Resume Next
    If mblnSys Then Exit Sub
    
    lstrTemp = ""
    If clstCharge.ListIndex <> -1 Then
        lstrTemp = clstCharge.List(clstCharge.ListIndex)
        If InStr(lstrTemp, " ") > 0 Then lstrTemp = Left(lstrTemp, InStr(lstrTemp, " ") - 1)
    End If
    ctxtCharge.Text = lstrTemp
    
End Sub



Private Sub ctabMain_Click(PreviousTab As Integer)
    subResizeTab
End Sub

Private Sub ctrwAllConclusion_DblClick()
    On Error Resume Next
    If ccmdAddConclusion.Enabled Then
        If Not ctrwAllConclusion.SelectedItem.Parent Is Nothing Then
            ccmdAddConclusion_click
        End If
    End If
End Sub

Private Sub ctrwAllConclusion_NodeClick(ByVal Node As MSComctlLib.Node)
    ccmdAddConclusion.Enabled = True
End Sub

Private Sub ctrwAllItem_DblClick()
    On Error Resume Next
    ccmdAddItem_Click
End Sub



Private Sub ctrwAllItem_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    ccmdAddItem.Enabled = True
End Sub

Private Sub ctrwSelectedConclusion_DblClick()
    On Error Resume Next
    If ccmdDeleteConclusion.Enabled Then
        If Not ctrwSelectedConclusion.SelectedItem.Parent Is Nothing Then
            ccmdDeleteConclusion_click
        End If
    End If
End Sub

Private Sub ctrwSelectedConclusion_NodeClick(ByVal Node As MSComctlLib.Node)
    ccmdDeleteConclusion.Enabled = True
End Sub

Private Sub ctrwSelectedItem_DblClick()
    On Error Resume Next
    ccmdDeleteItem_Click

End Sub

Private Sub ctrwSelectedItem_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    ccmdDeleteItem.Enabled = True
End Sub

Private Sub ctxtLetter_KeyPress(KeyAscii As Integer)
    '不能输入汉字，不能用Ctrl-V。
    On Error Resume Next
    If KeyAscii < 0 Or KeyAscii = 22 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ctxtNo_KeyPress(KeyAscii As Integer)
    '不能输入汉字，不能用Ctrl-V。
    On Error Resume Next
    If KeyAscii < 0 Or KeyAscii = 22 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtName.SetFocus
End Sub

'功能：加载窗体时,被始化窗体上的控件(状态栏,工具栏,各文本框.)
'      因为附加项目列表框,体检项目列表框,体检结论列表框'的列表项目数据会随各体检表的不同而有变化,
'      所以在体检表模板列表框的click事件中初始化。
Private Sub Form_Load()
    Dim lcolInfo As Collection
    
    On Error GoTo errHandler
    If mblnInUse Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1).Text = "窗体正在初始化，请稍侯.."
    
    '界面暂时不能操作。
    Me.Enabled = False
        
    '创建界面通用对象，通过该对象初始化工具栏。
    Set lcolInfo = New Collection
    With lcolInfo
        .Add "新增(&A)102"
        .Add "|"
        .Add "保存"
        .Add "删除"
        .Add "复制(&C)118"
        .Add "|"
        .Add "退出"
    End With
    Set mobjGUI = New cls界面通用对象
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
        .subInitialize lcolInfo, ""
    End With
    
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    '2012-05-22 翁乔 ↓↓↓
    '界面权限设置
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_业务设置_体检表设置_新增") = False Then
        ctbMain.Buttons(1).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_业务设置_体检表设置_保存") = False Then
        ctbMain.Buttons(3).Visible = False
        ctbMain.Buttons(2).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_业务设置_体检表设置_复制") = False Then
        ctbMain.Buttons(5).Visible = False
        ctbMain.Buttons(6).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_业务设置_体检表设置_删除") = False Then
        ctbMain.Buttons(4).Visible = False
    End If
    Set lobjTmp = Nothing
    '2012-05-22 ↑↑↑
    '标准版不提供收费功能。
    ctabMain.TabVisible(3) = False
    '屏蔽功能 'german
    ctabMain.TabVisible(2) = False
    ctabMain.TabVisible(1) = False
    
    '余下初始化工作在定时器中完成。
    Timer1.Enabled = True
    
    mblnInUse = True
    mblnSys = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetMedicalExamTemplate", "form_load", 6666, lstrError, False
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    '恢复界面可以操作。
    Me.Enabled = True
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Frame2.Width = Me.ScaleWidth - Frame2.Left - 60
    cfraMedicalTemplateName.Height = Me.ScaleHeight - cfraMedicalTemplateName.Top - 60
    Frame2.Height = Me.ScaleHeight - Frame2.Top - 60
    ctabMain.Width = Frame2.Width - ctabMain.Left - 60
    ctabMain.Height = Frame2.Height - ctabMain.Top - 60
    
    subResizeTab
End Sub

Private Sub subResizeTab()
    On Error Resume Next
    Select Case ctabMain.Tab
    Case 0
        Frame1(0).Width = ctabMain.Width - Frame1(0).Left - 60
        Frame1(0).Height = ctabMain.Height - Frame1(0).Top - 60
        ctrwSelectedItem.Height = Frame1(0).Height - ctrwSelectedItem.Top - 120
        ctrwAllItem.Height = Frame1(0).Height - ctrwAllItem.Top - 120
        
    Case 1
        Frame1(3).Width = ctabMain.Width - Frame1(3).Left - 60
        Frame1(3).Height = ctabMain.Height - Frame1(3).Top - 60
        ctrwSelectedConclusion.Height = Frame1(3).Height - ctrwSelectedConclusion.Top - 120
        ctrwAllConclusion.Height = Frame1(3).Height - ctrwAllConclusion.Top - 120
    Case 2
        Frame1(2).Width = ctabMain.Width - Frame1(2).Left - 60
        Frame1(2).Height = ctabMain.Height - Frame1(2).Top - 60
        clstSelectedBase.Height = Frame1(2).Height - clstSelectedBase.Top - 120
        clstAllBase.Height = Frame1(2).Height - clstAllBase.Top - 120
    Case 3
        Frame1(1).Width = ctabMain.Width - Frame1(1).Left - 60
        Frame1(1).Height = ctabMain.Height - Frame1(1).Top - 60
        clstCharge.Height = Frame1(1).Height - clstCharge.Top - 120
        clstDisposalIdea.Height = Frame1(1).Height - clstDisposalIdea.Top - 120
    End Select
End Sub

'功能：为了提高窗体的加载速度，把部分初始化工作放在定时器里完成。
Private Sub Timer1_Timer()
    Dim lobj体检表模板集  As Object
    Dim lobj体检项目集 As Object
    Dim lobjItem As Object
    Dim lobjRec As Object           '执行sql语句的结果。
    Dim lcolInfo As New Collection
    Dim objFrame As Frame
    Dim i As Integer
    
    On Error GoTo errHandler
    
    Timer1.Enabled = False
    
    '创建体检表模板对象。
    Set mobj体检表模板 = CreateObject("职业病对象.ClsMedicalExamTemplate")
    
    '创建体检表模板集对象，从而获取所有体检表模板名称。
    Set lobj体检表模板集 = CreateObject("职业病对象.ClsMedicalExamTemplateSet")
    'lobj体检表模板集.体检表类型 = 1
    Set lcolInfo = lobj体检表模板集.元素集
    
    '初始化体检表名称列表框。
    clstTemplate.Clear
    For i = 1 To lcolInfo.Count
        clstTemplate.AddItem lcolInfo(i)
    Next i
    
    '通过字典对象获取所有诊断处理意见，并初始化诊断处理意见列表框。
    clstDisposalIdea.Clear
    Set lobjRec = pobjDict.Fetch("诊断处理意见字典")
    If lobjRec Is Nothing Then
        Err.Raise 6666, , "使用字典管理对象.Fetch方法出错。请重新注册“字典管理.dll”"
    Else
        While Not lobjRec.EOF
            clstDisposalIdea.AddItem lobjRec("名称")
            lobjRec.movenext
        Wend
    End If
    '设置当前体检表模板为体检表名称框的第一项。
    If clstTemplate.ListCount = 0 Then
        '界面所有部分均不可操作。
        For Each objFrame In Frame1
            objFrame.Enabled = False
        Next
        ccmbSheet.Enabled = False
        ctxtNo.Enabled = False
        ctxtLetter.Enabled = False
        cchkAgain.Enabled = False
        ctbMain.Buttons(3).Enabled = False
        ctbMain.Buttons(4).Enabled = False
        ctbMain.Buttons(5).Enabled = False
    Else
        'clstTemplate.ListIndex = 0
    End If
    
    '获取所有体检单类型。
    Set lcolInfo = pobj业务对象.func获取所有体检单类型 'CreateObject("职业病对象.clsManageMedicalExam")
    ccmbSheet.Clear
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next
    
    '修改：2001-11-2（杨春）可选得体检项目显示在一颗树中。
    Set lobj体检项目集 = CreateObject("职业病对象.clsTestItemSet")
    
    '获取所有体检大类、体检项目。
    Set lobjRec = pobjDict.Fetch("职业病体检科室字典")
    '显示体检大类在ctrvItem中（其中节点的key=体检大类id）；
    ctrwAllItem.Nodes.Clear
    Do While Not lobjRec.EOF
        '通过"lobj体检项目集"依次获取各大类的体检项目。
        lobj体检项目集.体检大类 = lobjRec("InnerID")
        Set lobjItem = lobj体检项目集.体检项目
        If Not lobjItem.EOF Then
            ctrwAllItem.Nodes.Add , , "I" & lobjRec("InnerID"), lobjRec("编号") & " " & lobjRec("名称")
        End If
        '显示体检项目在ctrvItem中（其中节点的key=编码,parent=体检大类）。
        Do While Not lobjItem.EOF
            ctrwAllItem.Nodes.Add "I" & lobjRec("InnerID"), tvwChild, "I" & lobjItem("编码"), lobjItem("编码") & " " & lobjItem("名称")
            lobjItem.movenext
        Loop
        
        lobjRec.movenext
    Loop
    
    '通过字典对象，初始化所有体检结论大类
    Set lobjRec = pobjDict.Fetch("体检结论字典", "Parent=0")
    If lobjRec Is Nothing Then
        Err.Raise 6666, , "使用字典管理对象.Fetch方法获取体检结论字典内容时出错。请重新注册“字典管理.dll”"
    End If
    ctrwAllConclusion.Nodes.Clear
    While Not lobjRec.EOF
        'key:R+InnderID。
        ctrwAllConclusion.Nodes.Add , , "R" & lobjRec("InnerID").Value, lobjRec("名称").Value
        lobjRec.movenext
    Wend
    
    '获取体检结论。
    Set lobjRec = pobjDict.Fetch("体检结论字典", "Parent<>0")
    If lobjRec Is Nothing Then
        Err.Raise 6666, , "使用字典管理对象.Fetch方法获取体检结论字典内容时出错。请重新注册“字典管理.dll”"
    End If
    While Not lobjRec.EOF
        'key:I+InnerID。
        On Error Resume Next
        ctrwAllConclusion.Nodes.Add "R" & lobjRec("Parent").Value, tvwChild, "I" & lobjRec("InnerID").Value, lobjRec("名称").Value
        On Error GoTo errHandler
        lobjRec.movenext
    Wend
    
    
    If ctabMain.TabVisible(3) Then
        '获取所有收费标准：名称，总额。
        Set lcolInfo = pobj业务对象.所有体检收费标准
        mblnSys = True
        clstCharge.Clear
        For i = 1 To lcolInfo.Count
            clstCharge.AddItem Format(lcolInfo(i)("名称"), String(50, " ")) & " " & lcolInfo(i)("总额")
        Next
        mblnSys = False
    End If
    
    '修改：2003-6-27（杨春）增加业务设置“是否使用体检单”。
    If pobj业务对象.业务设置("是否使用体检单") <> "否" Then
        'Label1(0).Visible = True   '临时屏蔽-----------------------
        'ccmbSheet.Visible = True   '临时屏蔽-----------------------
        'Label2.Top = 1320          '临时屏蔽-----------------------
    End If

    
    '初始处于新增状态。
    Dim lblnCancel As Boolean
    mobjGUI_BeforeOperate "新增", lblnCancel
    
    '恢复界面可以操作。
    Me.Enabled = True
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    ctxtName.SetFocus
    
    'german
    '移植者:何啸天
    '日期:3.16
    '从字典表中获得选取的类型
    Set lobjRec = pobjDict.FetchEx("体检类型字典")
    tijian_leibie.Clear
    tijian_leibie.AddItem ""
    For i = 1 To lobjRec.recordcount
        tijian_leibie.AddItem lobjRec("名称")
        tijian_leibie.ItemData(tijian_leibie.NewIndex) = lobjRec("编号")
        lobjRec.movenext
    Next
    tijian_leibie.ListIndex = 0
    
    Set lobjRec = pobjDict.FetchEx("体检人类别字典")
    tijian_human_leixing.Clear
    tijian_human_leixing.AddItem ""
    For i = 1 To lobjRec.recordcount
        tijian_human_leixing.AddItem lobjRec("名称")
        tijian_human_leixing.ItemData(tijian_human_leixing.NewIndex) = lobjRec("编号")
        lobjRec.movenext
    Next
    tijian_human_leixing.ListIndex = 0
    
    Exit Sub
    
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetMedicalExamTemplate", "Timer1_Timer", 6666, lstrError, False
    mblnSys = False
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    '恢复界面可以操作。
    Me.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub
'功能:添加一项目体检项目
Private Sub ccmdAddItem_Click()
    Dim lstrCode As String '体检项目的编码。
    Dim i As Long
    Dim llngIndex As Long
    Dim llngChildren As Long
    
    On Error GoTo errHandler
    
    If ctrwAllItem.SelectedItem Is Nothing Then Exit Sub
    
    If ctrwAllItem.SelectedItem.Parent Is Nothing Then
        '添加一类项目。
        llngChildren = ctrwAllItem.SelectedItem.Children  '子节点数。
        llngIndex = ctrwAllItem.SelectedItem.Child.Index  '第一个子节点的索引。
        For i = llngIndex To llngIndex + llngChildren - 1
            lstrCode = ctrwAllItem.Nodes(i).Key                 '体检项目的编码。
            lstrCode = Right(lstrCode, Len(lstrCode) - 1)
            sub添加单个体检项目 lstrCode
        Next
    Else
        '添加单个项目。
        lstrCode = ctrwAllItem.SelectedItem.Key
        lstrCode = Right(lstrCode, Len(lstrCode) - 1)
        
        sub添加单个体检项目 lstrCode
    End If
    ccmdAddItem.Enabled = False
    ccmdDeleteItem.Enabled = False
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "ccmdAddItem_Click", 6666, lstrError, False
End Sub

'输入：paraItemCode 体检项目编码。
Private Sub sub添加单个体检项目(ByVal paraItemCode As String)
    Dim lobj体检项目 As Object
    
    '在体检表对象中添加项目。
    mobj体检表模板.Sub添加体检项目 paraItemCode
    
    Set lobj体检项目 = mobj体检表模板.体检项目集(paraItemCode)
    
    '在选中体检项目树中添加项目。忽略重复插入的错误。
    On Error Resume Next
    '插入一级字节点.
    ctrwSelectedItem.Nodes.Add , , "I" & lobj体检项目.体检大类, ctrwAllItem.Nodes("I" & lobj体检项目.体检大类).Text
    '插入项目。
    ctrwSelectedItem.Nodes.Add "I" & lobj体检项目.体检大类, tvwChild, "I" & lobj体检项目.编码, lobj体检项目.编码 & " " & lobj体检项目.名称

End Sub
Private Sub ccmdDeleteItem_Click()
    '删除选中的一项目体检项目到当前体检模板的体检项目列表框
    Dim lintSelectedIndex As Integer
    Dim lstrCode As String
    Dim lobjParent As Node
    Dim llngChildren  As Long
    Dim llngIndex As Long
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwSelectedItem.SelectedItem Is Nothing Then Exit Sub
    
    If ctrwSelectedItem.SelectedItem.Parent Is Nothing Then
        '删除一类。
        llngChildren = ctrwSelectedItem.SelectedItem.Children  '子节点数。
        llngIndex = ctrwSelectedItem.SelectedItem.Child.Index  '第一个子节点的索引。
        For i = llngIndex To llngIndex + llngChildren - 1
            lstrCode = ctrwSelectedItem.Nodes(i).Key                 '体检项目的编码。
            lstrCode = Right(lstrCode, Len(lstrCode) - 1)
            mobj体检表模板.Sub删除体检项目 lstrCode
        Next
        
         '删除大类根节点。
        ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Key
    Else
        '删除单个项目。
        lstrCode = ctrwSelectedItem.SelectedItem.Key
        lstrCode = Right(lstrCode, Len(lstrCode) - 1)
        mobj体检表模板.Sub删除体检项目 lstrCode
        
        '删除选中项目树中项目。
        Set lobjParent = ctrwSelectedItem.SelectedItem.Parent
        ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Key
        
        '若某大类被删除完毕，删除大类根节点。
        If lobjParent.Children = 0 Then
            ctrwSelectedItem.Nodes.Remove lobjParent.Key
        End If
    End If
    ccmdDeleteItem.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "ccmdDeleteItem_Click", 6666, lstrError, False
End Sub

'功能：添加一项体检结论。
Private Sub ccmdAddConclusion_click()
    Dim lobjNode As Node
    Dim llngChildren  As Long
    Dim llngIndex As Long
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwAllConclusion.SelectedItem Is Nothing Then Exit Sub
    
    
    Set lobjNode = ctrwAllConclusion.SelectedItem
    
    If lobjNode.Parent Is Nothing Then
        '选择大类。
        
        llngChildren = lobjNode.Children  '子节点数。
        
        If llngChildren > 0 Then
            llngIndex = lobjNode.Child.Index  '第一个子节点的索引。
            On Error Resume Next
            '添加上级节点。
            ctrwSelectedConclusion.Nodes.Add , , lobjNode.Key, lobjNode.Text
            
            '依次添加该大类下所有体检结论。
            For i = llngIndex To llngIndex + llngChildren - 1
                Set lobjNode = ctrwAllConclusion.Nodes(i)
                '在对象中添加选中的体检结论。
                mobj体检表模板.sub添加体检结论 Right(lobjNode.Key, Len(lobjNode.Key) - 1)
                '添加叶节点。
                ctrwSelectedConclusion.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
            Next
        End If
    Else '选择叶节点。
        
        '在对象中添加选中的体检结论。
        mobj体检表模板.sub添加体检结论 Right(lobjNode.Key, Len(lobjNode.Key) - 1)
        
        '在选中体检结论列表框中添加。
        On Error Resume Next
        '添加上级节点。
        ctrwSelectedConclusion.Nodes.Add , , lobjNode.Parent.Key, lobjNode.Parent.Text
        '添加叶节点。
        ctrwSelectedConclusion.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
        On Error GoTo errHandler
        
    End If
    
    '若选择体检结论超过10个，提示。
    If ctrwSelectedConclusion.Nodes.Count >= 20 Then
        sffuncMsg "请注意，你选择了这么多的体检结论，可能导致在批量录入体检结果后保存的时间会很长（因自动下体检结论）。请慎重选择该体检表可能下的体检结论。", sf警告
    End If
    
    ccmdAddConclusion.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "ccmdAddConclusion_click", 6666, lstrError, False
End Sub
'功能：删除一项体检结论。
Private Sub ccmdDeleteConclusion_click()
    Dim lobjNode As Node
    Dim llngChildren As Long
    Dim llngIndex As Long
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwSelectedConclusion.SelectedItem Is Nothing Then Exit Sub
    
    Set lobjNode = ctrwSelectedConclusion.SelectedItem
    
    If lobjNode.Parent Is Nothing Then
        '删除大类。
        
        llngChildren = lobjNode.Children  '子节点数。
        If llngChildren > 0 Then
            llngIndex = lobjNode.Child.Index  '第一个子节点的索引。
            
            '依次添加该大类下所有体检结论。
            For i = llngIndex To llngIndex + llngChildren - 1
                '在对象中删除选中的体检结论。
                mobj体检表模板.sub删除体检结论 Right(ctrwSelectedConclusion.Nodes(i).Key, Len(ctrwSelectedConclusion.Nodes(i).Key) - 1)
            Next
        End If
    Else '删除叶节点。
    
        '在对象中添加选中的体检结论。
        mobj体检表模板.sub删除体检结论 Right(lobjNode.Key, Len(lobjNode.Key) - 1)
        
        '若上级节点的叶节点被删除光了，整个大类要被删除。
        If lobjNode.Parent.Children = 1 Then
            Set lobjNode = lobjNode.Parent
        End If
        
    End If
    
    On Error Resume Next
    '在选中体检结论中删除。
    ctrwSelectedConclusion.Nodes.Remove lobjNode.Key
    
    ccmdDeleteConclusion.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "ccmdDeleteConclusion_click", 6666, lstrError, False
    
End Sub

'功能：添加附加项目。
Private Sub ccmdAddBase_click()
    Dim llngAllIndex As Long
    On Error GoTo errHandler
    
    '在对象中添加附加项目。
    mobj体检表模板.sub添加附加项目 clstAllBase.List(clstAllBase.ListIndex), False
    
    '向选中附加项目列表框中添加项目。
    With clstSelectedBase
        .AddItem clstAllBase.List(clstAllBase.ListIndex)
        .Selected(.NewIndex) = False
        .ListIndex = .NewIndex
    End With
    
    '从所有附加项目列表框中删除项目。
    With clstAllBase
        llngAllIndex = .ListIndex
        .RemoveItem llngAllIndex
        If .ListCount = 0 Then
            ccmdAddBase.Enabled = False
        ElseIf .ListCount > llngAllIndex Then
            .ListIndex = llngAllIndex
        Else
            .ListIndex = .ListCount - 1
        End If
    End With
    
    ccmdDeleteBase.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "ccmdAddBase_click", 6666, lstrError, False
End Sub

'功能：删除附加项目。
Private Sub ccmdDeleteBase_click()
    On Error GoTo errHandler
    
    '在对象中删除附加项目。
    mobj体检表模板.sub删除附加项目 clstSelectedBase.List(clstSelectedBase.ListIndex)
    
    '从所有附加项目列表框中删除项目。
    clstAllBase.AddItem clstSelectedBase.List(clstSelectedBase.ListIndex)
    clstAllBase.ListIndex = clstAllBase.NewIndex
    
    '从选中附加项目列表框中删除项目。
    clstSelectedBase.RemoveItem clstSelectedBase.ListIndex
    If clstSelectedBase.ListCount > 0 Then
        clstSelectedBase.ListIndex = 0
    Else
        ccmdDeleteBase.Enabled = False
    End If
    ccmdAddBase.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "ccmdDeleteBase_click", 6666, lstrError, False
End Sub

Private Sub clstAllConclusion_DblClick()
    On Error Resume Next
    ccmdAddConclusion_click
End Sub



Private Sub clstSelectedBase_Click()
    On Error GoTo errHandler
    
    '改变某一附加项目的 "是否必录"状态。
    With clstSelectedBase
        mobj体检表模板.sub添加附加项目 .List(.ListIndex), .Selected(.ListIndex)
    End With
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "clstSelectedBase_Click", 6666, lstrError, False
End Sub

'功能：同删除附加项目。
Private Sub clstSelectedBase_DblClick()
    On Error Resume Next
    ccmdDeleteBase_click
End Sub

'功能：同添加附加项目。
Private Sub clstAllBase_dblclick()
    On Error Resume Next
    ccmdAddBase_click
End Sub

'功能：同删除体检结论。
Private Sub clstSelectedConclusion_DblClick()
    On Error Resume Next
    ccmdDeleteConclusion_click
End Sub


'功能：通过对当前体检表模板的关键属性属值,获得体检表的其它属性,并显示在界面上。
Private Sub clstTemplate_Click()
    Dim objFrame As Frame
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    If clstTemplate.ListIndex = -1 Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1).Text = "正在获取体检表模板设置信息，请稍侯..."
    
    '为体检表模板对象的体检表名做初始化工作，从而获取该体检表的其他属性。
    mobj体检表模板.体检表名 = clstTemplate.List(clstTemplate.ListIndex)
    
    'german
    '添加者:何啸天
    '功能：预先将已经存在的信息数据填充在界面上，应用于新添加的控件
    Set lobjRec = dafuncGetData("select * from 职业病体检_体检表模板基本信息表 where(体检表名称='" + mobj体检表模板.体检表名 + "')")
    If (lobjRec.recordcount > 0) Then
        tijian_leibie.Text = lobjRec("体检类别")
        tijian_human_leixing.Text = lobjRec("体检人员类型")
    Else
        MsgBox "数据读取出现严重错误", 16, "信息"
        Exit Sub
    End If
    
    '把体检表模板对象的属性显示在界面上。
    subShowTemplate
    
    '选择了体检表模板名后,界面上的控件变为可操作状态.
    For Each objFrame In Frame1
        objFrame.Enabled = True
    Next
    'modify by lanchao 2015-03-15 将Buttons(3).Enabled = false -->Buttons(3).Enabled = True
    ctbMain.Buttons(3).Enabled = True
    ctbMain.Buttons(4).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    Exit Sub
    
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "clstTemplate_Click", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1).Text = "获取体检表模板设置信息失败。"
    Exit Sub
    Resume
End Sub

Private Sub ctxtLetter_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        cchkAgain.SetFocus
    End If
    Exit Sub
errHandler:
    
End Sub


Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ccmbSheet.Visible Then
            ccmbSheet.Enabled = True
            ccmbSheet.SetFocus
        Else
            ctxtNo.SetFocus
        End If
    End If

End Sub

'输入了体检表模板名后,界面内容随体检表模板名变化发生显示不同的内容.如果输入的体检表名已经存在,效果等同于点击
'体检表列表框中的体检表名.如果不存在,则视为新建体检表.
Private Sub ctxtName_LostFocus()
    On Error Resume Next
    If clstTemplate.ListIndex = -1 Then
        Frame2.Caption = "新增体检表：" & Trim(ctxtName)
    Else
        Frame2.Caption = "修改体检表：" & Trim(ctxtName)
    End If
    ctbMain.Buttons(3).Enabled = True
End Sub

Private Sub ctxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Err.Clear
    On Error GoTo errorHandler
    If KeyCode = 13 Then
        ctxtLetter.SetFocus
    End If
    Exit Sub
errorHandler:
    
End Sub

Private Sub ccmbSheet_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        ctxtNo.SetFocus
    End If
    Exit Sub
errHandler:
    
End Sub

'功能：清空界面。
Private Sub subClear()
    clstTemplate.ListIndex = -1
    ctxtName = ""
    ctbMain.Buttons(3).Enabled = True
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctxtName.Enabled = True
    ccmbSheet.Enabled = True
    ctxtNo.Enabled = True
    ctxtLetter.Enabled = True
    cchkAgain.Enabled = True
    mobj体检表模板.体检表名 = ""
    Frame2.Caption = "新增体检表："
    On Error Resume Next
    ctxtName.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj体检表模板 = Nothing
    Set mobjGUI = Nothing
    mblnInUse = False
End Sub

'功能：选中体检附加项目上移。
Private Sub ccmdUp_Click()
    Dim lstrItem As String
    Dim lblnSelected As Boolean
    Dim llngIndex As Long
    
    On Error GoTo errHandler
    
    '体检附加项目上移。
    With clstSelectedBase
        llngIndex = .ListIndex
        
        '在对象中上移附加项目。
        mobj体检表模板.sub附加项目上移 .List(llngIndex)
        
        If llngIndex > 0 Then
            '先记录选中项目内容、是否选中。
            lstrItem = .List(llngIndex)
            lblnSelected = .Selected(llngIndex)
            '先移出。
            .RemoveItem llngIndex
            '再加入。
            .AddItem lstrItem, llngIndex - 1
            If lblnSelected Then
                .Selected(llngIndex - 1) = True
            End If
        End If
    End With
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetMedicalExamTemplate", "ccmdUp_Click", 6666, lstrError, False
End Sub

'功能：选中体检附加项目下移。
Private Sub ccmdDown_Click()
    Dim lstrItem As String
    Dim lblnSelected As Boolean
    Dim llngIndex As Long
    
    On Error GoTo errHandler
    
    '体检附加项目下移。
    With clstSelectedBase
        llngIndex = .ListIndex
        
        '在对象中下移附加项目。
        mobj体检表模板.sub附加项目下移 .List(llngIndex)
        
        If llngIndex < .ListCount - 1 Then
            '先记录选中项目内容、是否选中。
            lstrItem = .List(llngIndex)
            lblnSelected = .Selected(llngIndex)
            '先移出。
            .RemoveItem llngIndex
            '再加入。
            .AddItem lstrItem, llngIndex + 1
            If lblnSelected Then
                .Selected(llngIndex + 1) = True
            End If
            .ListIndex = llngIndex + 1
        End If
    End With
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetMedicalExamTemplate", "ccmdUp_Click", 6666, lstrError, False

End Sub


'功能：处理工具栏上的按钮。
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim objFrame As Object
    Dim lstrTemp As String
    Dim lstrError As String
    Dim i As Long
    
    On Error GoTo errHandler
    
    Select Case Operate
        Case "新增"
            clstTemplate.ListIndex = -1
            ctxtName = ""
            For Each objFrame In Frame1
                objFrame.Enabled = True
            Next
            ctbMain.Buttons(3).Enabled = True
            ctbMain.Buttons(4).Enabled = False
            ctbMain.Buttons(5).Enabled = False
            ctxtName.Enabled = True
            ccmbSheet.Enabled = True
            ctxtNo.Enabled = True
            ctxtLetter.Enabled = True
            cchkAgain.Enabled = True
            mobj体检表模板.体检表名 = ""
            subShowTemplate
            Frame2.Caption = "新增体检表："
            On Error Resume Next
            ctxtName.SetFocus
            
        Case "删除"
            '询问。
            If sffuncMsg("你确认要删除体检表模板“" & mobj体检表模板.体检表名 & "”吗？", sf询问) Then
                '从库中删除。
                mobj体检表模板.Sub删除模板
                '从界面上删除。
                clstTemplate.RemoveItem clstTemplate.ListIndex
                If clstTemplate.ListCount > 0 Then
                    clstTemplate.ListIndex = 0
                Else
                    '新增。
                    subClear
                    ctbMain.Buttons(4).Enabled = False
                    ctbMain.Buttons(5).Enabled = False
                    ctbMain.Buttons(3).Enabled = False
                    ctxtName.Text = ""
                    ctxtName.SetFocus
                    Frame2.Caption = "新增体检表："
                End If
                
            End If
            Cancel = True
        Case "复制"
            Dim lstrNewName As String '复制后的新体检表名称。
            lstrNewName = Trim(ctxtName.Text) & "1"
            lstrNewName = InputBox("输入新体检表名(输入完毕后，可以修改体检表设置，但必须按保存按钮复制才算完毕):", "复制体检表", lstrNewName)
            lstrNewName = Trim(lstrNewName)
            If lstrNewName <> "" Then
                ctxtName.Text = lstrNewName
                mobj体检表模板.Sub复制模板 lstrNewName
                clstTemplate.ListIndex = -1
                ctxtNo = ""
                ctbMain.Buttons(4).Enabled = False
                ctbMain.Buttons(5).Enabled = False
                Frame2.Caption = "新增体检表：" & lstrNewName
                csbMain.Panels("Msg") = "复制后的新体检表，必须按“保存”按钮保存，复制才算完毕。” "
                
            End If
            
            Cancel = True
            
        Case "保存"
            Dim lbln是否已存在 As Boolean '判断保存的体检表设置是否已在库中存在。
            Dim lobjFrame As Variant
            lstrError = ""
            If ctrwSelectedItem.Nodes.Count = 0 Then
                lstrError = "没有选择体检项目。" & Chr(13) & Chr(10)
            End If
'            If Trim(ctxtNo) = "" Then '去除头部和尾部的空格
'                lstrError = lstrError & "没有输入体检表代号。" & Chr(13) & Chr(10)
'            End If
'            If Trim(ctxtLetter) = "" Then '去除头部和尾部的空格
'                lstrError = lstrError & "没有输入试管字母。" & Chr(13) & Chr(10)
'            End If
'            If ctrwSelectedConclusion.Nodes.Count = 0 Then
'                lstrError = lstrError & "没有选择体检结论。" & Chr(13) & Chr(10)
'            End If
            '判断体检表名称是否唯一。
            Dim lobjTemp As Object
            Set lobjTemp = CreateObject("职业病对象.ClsMedicalExamTemplate")
            
            'MsgBox CStr(clstTemplate.ListIndex), , "Message"
            
            If clstTemplate.ListIndex = -1 Then '如果是新增体检模板的话，那么此值为-1，新增时，模板列表框为失去焦点状态
                '新增。
                lobjTemp.体检表名 = Trim(ctxtName)
                If lobjTemp.是否已存在 Then
                    lstrError = lstrError & "体检表名称已存在。"
                Else
                    '设置体检表名称。
                    If lstrError = "" Then
                        mobj体检表模板.sub更换体检表名称 Trim(ctxtName) '未操作 是否存在
                        '将数据库中的表名全部替换
                    End If
                End If
            ElseIf Trim(clstTemplate.Text) <> Trim(ctxtName.Text) Then
                lobjTemp.体检表名 = Trim(ctxtName)
                If lobjTemp.是否已存在 Then
                    lstrError = lstrError & "体检表名称已存在。"
                Else
                    '设置体检表名称。
                    If lstrError = "" Then
                        mobj体检表模板.sub更换体检表名称 Trim(ctxtName) '职业病对象 - clsMedicalExamTemplate
                    End If
                End If
            End If
            If lstrError <> "" Then
                sffuncMsg "系统无法保存，因为：" & Chr(13) & Chr(10) & lstrError, sf警告
                Cancel = True
                Exit Sub
            End If
            
            Cancel = True
            If (tijian_human_leixing.Text = "") Then
                MsgBox "请您把体检人员类型选择好后，再保存", 16, "消息"
                Exit Sub
            End If
                
            If (tijian_leibie.Text = "") Then
                MsgBox "请您把体检类别选择好后，再保存", 16, "消息"
                Exit Sub
            End If
            
            MousePointer = 11
            csbMain.Panels(1).Text = "正在保存，请稍侯..."
            
            '界面暂时不能操作。
            ctbMain.Enabled = False
            cfraMedicalTemplateName.Enabled = False
            For Each lobjFrame In Frame1
                lobjFrame.Enabled = False
            Next
            ctxtName.Enabled = False
            ctxtLetter.Enabled = False
            ctxtNo.Enabled = False
            cchkAgain.Enabled = False
            
            '设置体检表模板属性。
            ccmbSheet.ListIndex = 0 '----------体检单默认设置------------
            With mobj体检表模板
                '.代号 = Trim(ctxtNo.Text)
                .代号 = "13" '默认屏蔽代号
                .体检单名称 = Trim(ctxtName.Text)
                '.试管字母编号 = Trim(ctxtLetter.Text)
                .试管字母编号 = "B"
                .是否复查体检表 = IIf(cchkAgain.Value = 1, True, False)
                
                '修改：2002-7-26（增加“是否年检表”属性）。
                .是否年检表 = IIf(cchkAnnual.Value = 1, True, False)
                
                '获取收费标准名称。
                .收费标准 = ctxtCharge.Text
                .诊断处理意见 = Empty
                .体检人员类型_ger = tijian_human_leixing.Text 'german
                .体检类别_ger = tijian_leibie.Text 'german

            End With
            '获取选中的诊断处理意见。
            lstrTemp = ""
            For i = 0 To clstDisposalIdea.ListCount - 1
                If clstDisposalIdea.Selected(i) Then
                    lstrTemp = lstrTemp & clstDisposalIdea.List(i) & ","
                End If
            Next i
            mobj体检表模板.诊断处理意见 = lstrTemp
            
            '保存体检表模板。若是新建，添加体检表名名称到列表中。
            lbln是否已存在 = mobj体检表模板.是否已存在
            
            mobj体检表模板.sub保存模板
            
            If Not lbln是否已存在 Then
                clstTemplate.AddItem Trim(ctxtName.Text)
                mblnSys = True
                clstTemplate.ListIndex = clstTemplate.NewIndex
                mblnSys = False
                ctxtNo = mobj体检表模板.代号
            Else
                '判断是否修改了体检表名称。
                If mobj体检表模板.体检表名 <> Trim(ctxtName.Text) Then
                    '修改库中体检表名称。
                    mobj体检表模板.sub更换体检表名称 Trim(ctxtName.Text)
                    
                    '修改列表中体检表名称。
                    mblnSys = True
                    i = clstTemplate.ListIndex
                    clstTemplate.RemoveItem i
                    If i < clstTemplate.ListCount - 1 Then
                        clstTemplate.AddItem Trim(ctxtName.Text), i
                    Else
                        clstTemplate.AddItem Trim(ctxtName.Text)
                    End If
                    mblnSys = False
                End If
            End If
            
            MsgBox "保存成功！", vbOKOnly + vbInformation, "系统提示"
            subClear
            '恢复界面可以操作。
            ctbMain.Enabled = True
            cfraMedicalTemplateName.Enabled = True
            For Each lobjFrame In Frame1
                lobjFrame.Enabled = True
            Next
            If clstTemplate.ListIndex >= 0 Then
                If Trim(clstTemplate.Text) <> Trim(ctxtName.Text) Then
                    i = clstTemplate.ListIndex
                    clstTemplate.RemoveItem i
                    If i = clstTemplate.ListCount Then
                        clstTemplate.AddItem Trim(ctxtName.Text)
                    Else
                        clstTemplate.AddItem Trim(ctxtName.Text), i
                    End If
                    clstTemplate.ListIndex = i
                End If
            End If
            ctxtName.Enabled = True
            ctxtLetter.Enabled = True
            ctxtNo.Enabled = True
            cchkAgain.Enabled = True
            
            MousePointer = 0
            csbMain.Panels(1).Text = "保存成功！"
            Cancel = True
    End Select
    
    Exit Sub
errHandler:
    mblnSys = False
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检表模板设置", "frmSetmedicalExamTemplate", "mobjGUI_BeforeOperate", 6666, lstrError, False
    If Operate = "保存" Then
        '恢复界面可以操作。
        ctbMain.Enabled = True
        cfraMedicalTemplateName.Enabled = True
        For Each lobjFrame In Frame1
            lobjFrame.Enabled = True
        Next
        ctxtName.Enabled = True
        ctxtLetter.Enabled = True
        ctxtNo.Enabled = True
        cchkAgain.Enabled = True
        MousePointer = 0
        csbMain.Panels(1).Text = "保存失败！"
        Cancel = True
    End If
End Sub

'功能：根据体检表模板对象的属性显示体检表模板设置信息。
Private Sub subShowTemplate()
    Dim lobj体检项目集 As Object '职业病对象部件.clsTestItemSet
    Dim lobjRec As Object        '执行sql语句的结果。
    Dim lcolInfo As Collection   '体检表模板对象的“体检结论集”、“基本附加项目集”、“体检项目集”等属性集合。
    Dim lcolItem As Collection   'lcolInfo集合中的元素。
    Dim lobj体检项目 As Object   '职业病对象部件.ClsTestItem
    Dim lobjNode As Node
    Dim lstrIdea As String
    Dim lstrItem As String
    Dim lintPos As Long
    Dim i As Long
    
    On Error GoTo errHandler
    
    '清空界面。
'    subClear
    
    clstAllBase.Clear
    
    '为体检表模板对象的体检表名做初始化工作，从而获取该体检表的其他属性。
    With mobj体检表模板
        If .是否已存在 Then
            ctxtName.Text = .体检表名
            ctxtNo.Text = .代号
            ccmbSheet.Text = .体检单名称
            ctxtLetter.Text = .试管字母编号
            cchkAgain.Value = IIf(.是否复查体检表, 1, 0)
            If cchkAgain.Value = 0 Then
                cchkAnnual.Value = IIf(.是否年检表, 1, 0)
            Else
                cchkAnnual.Value = 0
            End If
            
            Frame2.Caption = "修改体检表：" & ctxtName.Text
        Else
            Frame2.Caption = "新增体检表：" & ctxtName.Text
        End If
    End With
    
    
    '通过“体检管理”业务对象获取所有体检附加项目，并初始初始化体检附加项目列表框。
    Set lobjRec = pobj业务对象.所有体检附加项目
    clstAllBase.Clear
    While Not lobjRec.EOF
        clstAllBase.AddItem lobjRec("附加项目").Value
        lobjRec.movenext
    Wend
    
    
    '初始化当前体检结论框
    Set lcolInfo = mobj体检表模板.体检结论集
    ctrwSelectedConclusion.Nodes.Clear
    For Each lcolItem In lcolInfo
    
        '插入一级字节点.
        On Error Resume Next
        ctrwSelectedConclusion.Nodes.Add , , ctrwAllConclusion.Nodes("I" & lcolItem("体检结论ID")).Parent.Key, ctrwAllConclusion.Nodes("I" & lcolItem("体检结论ID")).Parent.Text
        '插入项目。
        ctrwSelectedConclusion.Nodes.Add ctrwAllConclusion.Nodes("I" & lcolItem("体检结论ID")).Parent.Key, tvwChild, "I" & lcolItem("体检结论ID"), lcolItem("名称")
        On Error GoTo errHandler
    Next
        
    '初始化当前体检模板的体检附加项目列表框
    Set lcolInfo = mobj体检表模板.基本附加项目集
    clstSelectedBase.Clear
    For Each lcolItem In lcolInfo
        clstSelectedBase.AddItem lcolItem("附加项目")
        clstSelectedBase.Selected(clstSelectedBase.NewIndex) = lcolItem("是否必录")
        i = 0
        While i <= clstAllBase.ListCount - 1
            If clstAllBase.List(i) = lcolItem("附加项目") Then
                clstAllBase.RemoveItem i
            Else
                i = i + 1
            End If
        Wend
    Next
        
    '初始化当前体检表模板体检项目列表框.
    '修改：2001-11-2（杨春）可选得体检项目显示在一颗树中。
    Set lcolInfo = mobj体检表模板.体检项目集
    ctrwSelectedItem.Nodes.Clear
    For Each lobj体检项目 In lcolInfo
        '插入一级字节点.
        On Error Resume Next
        ctrwSelectedItem.Nodes.Add , , "I" & lobj体检项目.体检大类, ctrwAllItem.Nodes("I" & lobj体检项目.体检大类).Text
        '插入项目。
        ctrwSelectedItem.Nodes.Add "I" & lobj体检项目.体检大类, tvwChild, "I" & lobj体检项目.编码, lobj体检项目.编码 & " " & lobj体检项目.名称
        On Error GoTo errHandler
    Next
    ccmdAddItem.Enabled = False
    ccmdDeleteItem.Enabled = False
    
    '初始化当前体检表诊断处理意见
    lstrIdea = Trim(mobj体检表模板.诊断处理意见)
    If Right(lstrIdea, 1) <> "," Then lstrIdea = lstrIdea & ","
    For i = 0 To clstDisposalIdea.ListCount - 1
        If InStr(1, lstrIdea, clstDisposalIdea.List(i) & ",") > 0 Then
            clstDisposalIdea.Selected(i) = True
        Else
            clstDisposalIdea.Selected(i) = False
        End If
    Next
    
    '体检项目,附加项目,体检结论列表框中。
    '如果所有项目被选中,则添加按钮变为不可用.如果没有一项目被选中,则删除按钮变为不可用.
    If ctrwAllConclusion.SelectedItem Is Nothing Then
        ccmdAddConclusion.Enabled = False
    Else
        ccmdAddConclusion.Enabled = True
    End If
    
    If ctrwSelectedConclusion.SelectedItem Is Nothing Then
        ccmdDeleteConclusion.Enabled = False
    Else
        ccmdDeleteConclusion.Enabled = True
    End If
   
    If clstAllBase.ListCount > 0 Then
        clstAllBase.ListIndex = 0
        ccmdAddBase.Enabled = True
    Else
        ccmdAddBase.Enabled = False
    End If
    
    If clstSelectedBase.ListCount > 0 Then
        clstSelectedBase.ListIndex = 0
        ccmdDeleteBase.Enabled = True
    Else
        ccmdDeleteBase.Enabled = False
    End If
    
    
    '设置收费标准。
    If mobj体检表模板.收费标准 = "" Then
        clstCharge.ListIndex = -1
        ctxtCharge.Text = ""
    Else
        i = gffuncItemIsInListBox(clstCharge, mobj体检表模板.收费标准)
        clstCharge.ListIndex = i
    End If
    Exit Sub
    
errHandler:
    sfsub错误处理 "体检表模板设置", "frmSetMedicalExamTemplate", "subShowTemplate", Err.Number, Err.Description, True
End Sub

