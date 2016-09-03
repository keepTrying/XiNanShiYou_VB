VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "疾病控制管理信息系统"
   ClientHeight    =   10995
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   13710
   ClipControls    =   0   'False
   ForeColor       =   &H00B9F7D3&
   Icon            =   "frmMain1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   13710
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   5040
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F7F2E1&
      Caption         =   "待办工作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   9000
      TabIndex        =   24
      Top             =   7800
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox clblMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F2E1&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   240
         Width           =   4815
      End
   End
   Begin MSComctlLib.StatusBar cstatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   10635
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4817
            MinWidth        =   1060
            Text            =   "用户编号："
            TextSave        =   "用户编号："
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4817
            MinWidth        =   1060
            Text            =   "用户姓名："
            TextSave        =   "用户姓名："
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4815
            MinWidth        =   1058
            Text            =   "工作站名："
            TextSave        =   "工作站名："
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9049
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1800
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7635
      Left            =   -45
      TabIndex        =   7
      Top             =   500
      Width           =   1635
      Begin VB.Image cimgPhoto 
         Height          =   1335
         Left            =   240
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   7
         Left            =   0
         Picture         =   "frmMain1.frx":0CCA
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   6
         Left            =   0
         Picture         =   "frmMain1.frx":111C
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   5
         Left            =   0
         Picture         =   "frmMain1.frx":156E
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   4
         Left            =   0
         Picture         =   "frmMain1.frx":19C0
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   3
         Left            =   0
         Picture         =   "frmMain1.frx":1E12
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   2
         Left            =   0
         Picture         =   "frmMain1.frx":2264
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   1
         Left            =   0
         Picture         =   "frmMain1.frx":26B6
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image Image5 
         Height          =   765
         Left            =   15
         Picture         =   "frmMain1.frx":2B08
         Stretch         =   -1  'True
         Top             =   6705
         Width           =   1725
      End
      Begin VB.Image cimgButton 
         Height          =   330
         Index           =   10
         Left            =   150
         Picture         =   "frmMain1.frx":5B12
         Stretch         =   -1  'True
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   8
         Left            =   300
         Picture         =   "frmMain1.frx":5F64
         Top             =   120
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   0
         Left            =   105
         Picture         =   "frmMain1.frx":63B6
         Top             =   570
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label cmnuItem 
         BackColor       =   &H009EE9C4&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   -240
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "可 用 操 作"
         Height          =   165
         Left            =   615
         TabIndex        =   19
         Top             =   -100
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   255
         Picture         =   "frmMain1.frx":6808
         Stretch         =   -1  'True
         Top             =   1620
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   585
         TabIndex        =   18
         Top             =   195
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   17
         Top             =   195
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   3
         Left            =   525
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   4
         Left            =   570
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   5
         Left            =   555
         TabIndex        =   14
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   6
         Left            =   540
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   7
         Left            =   630
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   8
         Left            =   660
         TabIndex        =   11
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   9
         Left            =   585
         TabIndex        =   10
         Top             =   165
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   10
         Left            =   630
         TabIndex        =   9
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   8
         Top             =   615
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image cpicButton 
         Height          =   360
         Index           =   0
         Left            =   60
         Picture         =   "frmMain1.frx":6EA3
         Top             =   -285
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "自动升级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   7
      Left            =   12840
      TabIndex        =   28
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   7800
      Width           =   4575
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "统计报表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   6720
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image cimgButton 
      Height          =   300
      Index           =   9
      Left            =   45
      Picture         =   "frmMain1.frx":9918
      Top             =   645
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   1575
      Left            =   1600
      Picture         =   "frmMain1.frx":9D6A
      Top             =   500
      Width           =   2535
   End
   Begin VB.Label clblSysName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卫生监督所管理信息系统"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   3300
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帮助"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   11880
      TabIndex        =   20
      Top             =   120
      Width           =   420
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统管理"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   840
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户编号：           用户名称：           工作站名："
      Height          =   180
      Left            =   2700
      TabIndex        =   4
      Top             =   7530
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   6
      Left            =   11280
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "口令修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   10320
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   420
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字典管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   9360
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   3525
      Picture         =   "frmMain1.frx":E268
      Stretch         =   -1  'True
      Top             =   1590
      Width           =   7875
   End
   Begin VB.Image cimgBackground 
      Height          =   585
      Left            =   0
      Picture         =   "frmMain1.frx":1F0A2
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   15225
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj类 As Object                '当前平台的所有类
Private mobj组 As Object                '当前平台的所有操作组
Private mobj操作 As Object              '当前平台的所有操作
Private mobj查询 As Object              '当前平台的所有查询
Private mobj报表 As Object              '当前平台的所有报表
Private mobj查询别名 As Object          '当前平台的所有查询详细信息
Private mobj报表别名 As Object          '当前平台的所有报表详细信息
Private mobjSmartInfos As Object
Private mblnRe As Boolean               '是否确认退出
Private mstr当前类 As String            '当前选中类的名称
Private mblnLoadForm As Boolean         '标识是否在在创建窗体。
Private mstrOper(1 To 10) As String
Private mstrMnu(1 To 15) As String
Private mintMnu As Integer

'修改：新增对象
Private X As Object

Private mblnAutoUpgrade As Boolean      '当前要进行自动升级

'报表查询需要的变量。
Private mobjFrontQueryManager As Object
Private mobjSysAccObj As Object         '数据驱动器.clsSystemAccessObject。

Private mintMinutes As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cimg组_Click(Index As Integer)

End Sub


Private Sub cimgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To clbl通用操作.Count - 1
        clbl通用操作(i).FontUnderline = False
        clbl通用操作(i).ForeColor = vbWhite
    Next
End Sub


Private Sub clbl通用操作_Click(Index As Integer)
    On Error GoTo errHandle
    Dim lobjTemp As Object
    Dim lobj报表 As Object
    Dim lobj查询 As Object
    Dim llngWndProc As Long
    Dim llng窗体句柄 As Long
    
    clbl通用操作(Index).FontUnderline = False
    clbl通用操作(Index).ForeColor = vbWhite
    Select Case clbl通用操作(Index).Caption
        Case "字典管理"
            frm字典列表.subLoad
            If frm字典列表.clbl字典.Count = 1 Then Exit Sub  '如果没有字典设置权限则不响应用户操作
            If frm字典列表.clbl字典.Count = 2 Then            '如果只有一个字典设置的权限则不使用弹出式菜单
                Call sub设置字典(frm字典列表.clbl字典(1).Caption)
            Else                                   '有多个字典设置的权限则使用弹出式菜单
                Unload frm操作列表
                Set frm字典列表.pobjParent = Me
                frm字典列表.Move clbl通用操作(Index).Left, Me.Top + 800
                frm字典列表.Show , Me
            End If
         
         Case "系统管理"
            frm系统管理.Hide

            mobj类.Filter = ""
            
            If um用户编号 = "0000" Then
                mobj类.Filter = "操作组='系统' or 操作组='系统管理'"

            Else
                mobj类.Filter = "所属类名= '业务类' and 操作组='系统'"
            End If
            
            If mobj类.RecordCount > 0 Then
                frm系统管理.subLoad mobj类("操作组"), mobj组, mobj操作
                    
                Set frm系统管理.pobjParent = Me
                frm系统管理.Move clbl通用操作(Index).Left, Me.Top + 800
                frm系统管理.Show , Me
            End If
            mobj类.Filter = "所属类名= '业务类'"
            
        Case "查询"      '选取该平台所有查询
                '启动查询界面。
                Dim lobj通用查询 As Object
                Set lobj通用查询 = CreateObject("通用查询.cls通用查询")

                llng窗体句柄 = lobj通用查询.funcStart("系统管理_通用查询", pstr子系统许可)
                
                '设定打开的窗体为主窗体的子窗体。
                If llng窗体句柄 <> -2 Then
                    '向集合中加入操作名称
                    If Not sffunc判断集合键值是否存在(pcol操作名称, CStr(llng窗体句柄)) Then
                        On Error Resume Next
                        SetParent llng窗体句柄, Me.hWnd
                        llngWndProc = SetWindowLong(llng窗体句柄, GWL_WNDPROC, AddressOf funcClassing)
                        pcolWndProc.Add llngWndProc, CStr(llng窗体句柄)
                        pcol操作名称.Add "查询统计", CStr(llng窗体句柄)
                        pcol子窗体句柄.Add llng窗体句柄, "查询统计"

                        Call MoveWindow(llng窗体句柄, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
                        Err.Clear
                    End If
                End If
        Case "统计报表"       '选取该平台所有报表
            '启动报表查询界面。
            Call sub启动报表查询
                
        Case "口令修改"
            frm密码修改.Show vbModal, Me
        Case "退出"
            Unload Me
        Case "帮助"
            'MsgBox "正在制作中！", vbOKOnly, "帮助"
            ShellExecute Me.hWnd, "Open", App.Path + "\用户手册\manual.chm", "", "", 1
        Case "自动升级"
            If MsgBox("进行自动升级时，必须先退出系统。确认数据已经保存好，可以退出了吗？", vbQuestion + vbYesNo, "系统询问") = vbNo Then Exit Sub
            Shell App.Path & "\autoupgrade.exe '" & pstr用户编号 & "'", vbNormalFocus
            mblnAutoUpgrade = True
            Unload Me
    End Select
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm主界面", "cSSListBar通用_ListItemClick", Err.Number, Err.Description, False)
End Sub


Private Sub clbl通用操作_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    clbl通用操作(Index).FontUnderline = True
    clbl通用操作(Index).ForeColor = &H80FF&
End Sub

Private Sub cmnuItem_Click(Index As Integer)
    Dim ii As Integer, j As Integer
    
    If cmnuSubItem(1).Visible And cmnuSubItem(1).Top = cpicButton(Index).Top + 450 Then '将菜单收起来
        For ii = 1 To 10
            cmnuSubItem(ii).Visible = False
            cimgButton(ii).Visible = False
        Next
        For ii = 1 To mintMnu
            cpicButton(ii).Top = cpicButton(ii - 1).Top + cpicButton(ii - 1).Height + 30
            cmnuItem(ii).Top = cpicButton(ii).Top + 90
        Next
        Exit Sub
    End If
    For ii = 1 To Index
        cpicButton(ii).Top = cpicButton(ii - 1).Top + cpicButton(ii - 1).Height + 30
        cmnuItem(ii).Top = cpicButton(ii).Top + 90
    Next

    cmnuSubItem(0).Top = cpicButton(ii - 1).Top + 120
    sub初始化操作列表 mstrMnu(Index)
    For j = 1 To 10
        If cmnuSubItem(j) = "" Then Exit For
    Next
    If ii <= mintMnu Then
        cpicButton(ii).Top = cmnuSubItem(j - 1).Top + cmnuSubItem(j - 1).Height + 100
        cmnuItem(ii).Top = cpicButton(ii).Top + 90
        For ii = ii + 1 To mintMnu
            cpicButton(ii).Top = cpicButton(ii - 1).Top + cpicButton(ii - 1).Height + 30
            cmnuItem(ii).Top = cpicButton(ii).Top + 90
        Next
    End If
    cmnuItem(Index).FontUnderline = False
    cmnuItem(Index).ForeColor = vbBlack
End Sub

Private Sub cmnuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To cmnuItem.Count - 1
        If i = Index Then
            cmnuItem(Index).FontUnderline = True
            cmnuItem(Index).ForeColor = &H80FF&
        Else
            cmnuItem(i).FontUnderline = False
            cmnuItem(i).ForeColor = vbBlack
        End If
    Next
End Sub


Private Sub cmnuSubItem_Click(Index As Integer)
    If mstrOper(Index) = "统计报表" Then
        Call sub启动报表查询
    Else
        sub创建窗体 mstrOper(Index)
    End If
    cmnuSubItem(Index).FontUnderline = False
    cmnuSubItem(Index).ForeColor = vbBlack
End Sub

Private Sub cmnuSubItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To cmnuSubItem.Count - 1
        If cmnuSubItem(i).Visible Then
            If i = Index Then
                cmnuSubItem(Index).FontUnderline = True
                cmnuSubItem(Index).ForeColor = &H80FF&
            Else
                cmnuSubItem(i).FontUnderline = False
                cmnuSubItem(i).ForeColor = vbBlack
            End If
        End If
    Next
End Sub

Private Sub Form_Activate()
    sub显示代办工作
End Sub

Public Sub sub显示代办工作()
    Dim lobjRec As Object
    On Error Resume Next
    
    '获取待办工作。
    Set lobjRec = dafuncGetData("exec 系统管理_获取待办工作 '" & um用户编号 & "'")
    clblMessage.Text = ""
    If lobjRec.RecordCount > 0 Then
        clblMessage.Text = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    If clblMessage.Text = "" Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '响应Esc键
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer    '循环变量
    Dim ii As Integer   '循环变量
    Dim j As Long
    Dim lobjSys As New FileSystemObject
    
    On Error Resume Next
    Dim lstrServer As String
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    
    clblSysName = um防疫站名 & "管理信息系统"
    
    '获取本机名称。
    Dim lstrLocalName As String
    Dim lobjRec As Object
    
    lstrLocalName = funcGetLocalName()

    cstatusBar.Panels(1).Text = "用户编号：" & um用户编号
    cstatusBar.Panels(2).Text = "用户姓名：" & um用户名
    cstatusBar.Panels(3).Text = "工作站名：" & lstrLocalName
    
    mblnAutoUpgrade = False
    
    If Not pbln注销 Then
        '判断该用户的注册信息
        Dim lobjCheck As New cls用户检查, lstrExpireDate As String
'        lstrExpireDate = lobjCheck.funcGetExpireDate()
'        If lstrExpireDate = "" Then
'            frm身份信息录入.Show 1
'            lstrExpireDate = lobjCheck.funcGetExpireDate()
'            If lstrExpireDate = "" Then
'                MsgBox "您不是系统的正式用户，系统无法运行！", vbCritical, "系统提示"
'                End
'            End If
'        End If
'        If CDate(lstrExpireDate) < Date Then
'            MsgBox "您的系统已经超过了使用期限，请与软件提供商联系！", vbCritical, "系统提示"
'            End
'        End If
'        If CDate(lstrExpireDate) < CDate("2050-01-01") Then
'            If DateDiff("d", Date, CDate(lstrExpireDate)) < 30 Then
'                MsgBox "您的使用期限还剩下" & DateDiff("d", Date, CDate(lstrExpireDate)) & "天！", vbInformation, "系统提示"
'            End If
'        End If
    End If
    '获取许可
    Dim larrLic(1 To 16) As String
    Dim lstrSubSystem As String

    lstrSubSystem = lobjCheck.funcGetSubSystem()
    If lstrSubSystem = "" Then
        MsgBox "您没有使用任何系统功能的权限，请与软件供应商联系！", vbInformation, "系统提示"
        End
    End If

    larrLic(1) = "卫生许可证管理,行政许可"
    larrLic(2) = "监督执法管理,突发事件"
    larrLic(3) = "体检管理"
    larrLic(4) = "健康证管理,健康证"
    larrLic(5) = "卫生监测管理,卫生检测"
    larrLic(6) = "检验管理"
    larrLic(7) = "计划免疫管理"
    larrLic(8) = "后勤管理,行政办公管理"
    larrLic(9) = "收费管理"
    larrLic(10) = "站长查询,领导查询"

    larrLic(11) = "质控管理"
    larrLic(12) = "短信通知"
    larrLic(13) = "医疗机构管理"
    larrLic(14) = "职业卫生管理"
    larrLic(15) = "稽查管理,稽查"
    larrLic(16) = "疫苗管理"

    pstr子系统许可 = "系统管理,单位档案管理,"
    For i = 1 To Len(lstrSubSystem)
        If Mid(lstrSubSystem, i, 1) = "1" Then pstr子系统许可 = pstr子系统许可 + larrLic(i) + ","
    Next
    
'    If Not pbln注销 And Not pbln试用 Then
'        '获取系统配置中的加密狗服务器名。
'        Dim lstrDogServer As String
'        lstrDogServer = sffuncGetSetting("系统管理", "数据库配置", "网络锁服务器名")
'        If lstrDogServer = "" Then lstrDogServer = lstrServer
'
'        Call dafuncGetData("sp_addlinkedserver '" & lstrDogServer & "'")
'        Err.Clear
'        Set lobjRec = dafuncGetData("exec [" & lstrDogServer & "].master.dbo.ryCheck")
'        If Err.Number = 0 Then
'            Select Case lobjRec(0)
'            Case 1, 2, 3, 4, 5, 6
'                MsgBox "网络锁服务安装不正常，导致安全检查失败。" & Chr(13) & Chr(10) & "请重新安装网络锁服务程序。", vbCritical, "系统错误"
'                End
'            Case Else
''                If (lobjRec(0) And (&H8000)) = 32768 Then
'                    '找到了。获取许可。
'                    Dim llngBit As Long
'                    Dim larrLic(1 To 15) As String
'                    larrLic(1) = "卫生许可证管理,行政许可"
'                    larrLic(2) = "监督执法管理,突发事件"
'                    larrLic(3) = "体检管理"
'                    larrLic(4) = "健康证管理,健康证"
'                    larrLic(5) = "卫生监测管理,卫生检测"
'                    larrLic(6) = "检验管理"
'                    larrLic(7) = "计划免疫管理"
'                    larrLic(8) = "后勤管理,行政办公管理"
'                    larrLic(9) = "收费管理"
'                    larrLic(10) = "站长查询,领导查询"
'
'                    larrLic(11) = "质控管理"
'                    larrLic(12) = "短信通知"
'                    larrLic(13) = "医疗机构管理"
'                    larrLic(14) = "职业卫生管理"
'                    larrLic(15) = "稽查管理,稽查"
'                    pstr子系统许可 = ""
'                    llngBit = &H4000
'                    For i = 1 To 15
'                        If (lobjRec(0) And llngBit) = llngBit Then
'                            pstr子系统许可 = pstr子系统许可 & larrLic(i) & ","
'                        End If
'                        llngBit = llngBit / 2
'                    Next
''                Else
''                    MsgBox "网络锁不是正式版的，安全检查失败！系统无法运行。" & Chr(13) & Chr(10) & "请与软件供应商联系。", vbCritical, "系统错误"
''                    End
''                End If
'            End Select
'        Else
'            MsgBox "网络锁服务安装不正常，导致安全检查失败。" & Chr(13) & Chr(10) & "请重新安装网络锁服务程序。安装前确保网络锁服务器上已安装Sql Server2000。", vbCritical, "系统错误"
'            End
'        End If
'        pstr子系统许可 = "系统管理,单位档案管理," & pstr子系统许可
'    ElseIf pbln试用 Then
'        pstr子系统许可 = ""
'        sub检查试用期限
'        Timer1.Enabled = True
'    End If
    
    um子系统许可 = pstr子系统许可
    
    dasubSetQueryTimeout 6000

    On Error GoTo errHandle
    
    Dim llngWndProc As Long
    llngWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf funcClassing)
    pcolWndProc.Add llngWndProc, CStr(Me.hWnd)
    
    pobj平台结构.平台名称 = um用户编号 '给平台结构的用户编号赋值
    Set mobj类 = pobj平台结构.操作组分类  '取该用户的平台结构
    Set mobj组 = pobj平台结构.Operations
    Set mobj操作 = pobj平台结构.Operation
    Set mobj查询 = pobj平台结构.Queries
    Set mobj查询别名 = pobj平台结构.Query
    Set mobj报表 = pobj平台结构.Reports
    Set mobj报表别名 = pobj平台结构.Report
    Set mobjSmartInfos = pobj平台结构.SmartInfos
    
    j = 1
    
    mintMnu = 1
    
    '初始化操作组图片
    If um用户编号 = "0000" Then
        mobj类.Filter = "操作组='系统' or 操作组='系统管理'"
'        clbl通用操作(3).Enabled = True
    Else
        mobj类.Filter = "所属类名= '业务类'"
    End If
    For ii = 1 To mobj类.RecordCount
        Select Case mobj类("操作组")
            Case "系统", "系统管理"
                clbl通用操作(3).Enabled = True

            Case Else
        
                Load cmnuItem(mintMnu)
                Load cpicButton(mintMnu)
                cpicButton(mintMnu).Left = cpicButton(0).Left
                cpicButton(mintMnu).Top = cpicButton(mintMnu - 1).Top + cpicButton(mintMnu - 1).Height + 30
                cpicButton(mintMnu).Visible = True
                cmnuItem(mintMnu) = mobj类("操作组")
                cmnuItem(mintMnu).Left = cmnuItem(0).Left
                cmnuItem(mintMnu).Top = cpicButton(mintMnu).Top + 90
                cmnuItem(mintMnu).Visible = True
                mstrMnu(mintMnu) = mobj类("操作组")
                mintMnu = mintMnu + 1
                
        End Select
        sub初始化字典列表 mobj类("操作组")
        
        mobj类.MoveNext
    Next
    mintMnu = mintMnu - 1
    
    '判断用户是否字典管理员
    Set lobjRec = dafuncGetData("select * from 系统管理_字典_用户管理级别表 where 用户编号='" & um用户编号 & "'")
    If lobjRec.RecordCount = 0 Then clbl通用操作(4).Enabled = False
    
    mobj类.Filter = "所属类名= '业务类'"
    
  
    On Error Resume Next
    If um用户编号 = "0000" Then
        '判断是否第一次运行。
        Dim lobj记忆 As cls用户操作记忆
        Set lobj记忆 = New cls用户操作记忆
        lobj记忆.用户编号 = "0000"
        lobj记忆.业务名 = "系统管理"
        If lobj记忆.记忆项值("第一次运行") <> "否" Then
            Call sub创建窗体("系统管理_运行状态设置")
            '保存已运行果状态。
            lobj记忆.sub覆盖记忆值 "第一次运行", "否"
        End If
    End If
    
    Me.Caption = pstrSysName
    
    On Error Resume Next
    cimgPhoto.Picture = pmfunc获取图片(um用户编号, "员工管理")
    If cimgPhoto.Picture <> 0 Then
        Image5.Visible = False
    End If
    Image1.Picture = LoadPicture(App.Path & "\内页-启动.jpg")
    '无须再判断注册信息，登录系统时已经判断过一次
    'Timer2.Enabled = True
    
    '启动信使服务。
    'Timer3.Enabled = True
    Exit Sub
errHandle:
    If Err.Number = 40003 Or Err.Number = 40002 Then
    Resume Next
    Else
    Call sfsub错误处理("主程序", "frm主界面", "Form_Load", Err.Number, Err.Description, False)
    End If
    Exit Sub
    Resume
End Sub


'功能：初始化报表查询对象。（本窗体启动时调用该方法）。
Private Sub sub初始化报表查询对象()

    On Error Resume Next
    
    Set mobjSysAccObj = CreateObject("数据服务器.clsSystemBaseAccess")
    If Err <> 0 Then
        Err.Clear
        sub注册报表查询部件
        Err.Clear

        Set mobjSysAccObj = CreateObject("数据服务器.clsSystemBaseAccess")
    End If
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "无法创建“数据服务器.dll”的对象，“统计报表”操作不可用。若要使用报表，请退出系统，并重新启动机器后执行系统安装程序。"
    End If
    
    
    '创建报表查询管理对象。
    Set mobjFrontQueryManager = CreateObject("道源报表查询器.clsFrontQueryManager")
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "无法创建“道源报表查询器.dll”的对象，“统计报表”操作不可用。若要使用报表，请退出系统，并重新启动机器后执行系统安装程序。"
    End If
   
    
    '初始化数据访问对象。
    Dim lstrDatabase As String     '数据库名
    lstrDatabase = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    mobjSysAccObj.ODBCConnectString = "DSN=WSFY2001;UID=user26;PWD=welcome;DATABASE=" & lstrDatabase
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "无法根据数据源“WSFY2001”与数据库建立连接。“统计报表”操作不可用。若要使用报表，请退出系统，并重新启动机器后执行系统安装程序。"
    End If
    
    '判断临时路径是否存在。若不存在，创建之。
    If Dir("c:\temp", vbDirectory) = "" Then
        MkDir "c:\temp"
    End If
    Err.Clear
        
    '初始化报表查询对象。
    mobjFrontQueryManager.subFrontQueryInitalize mobjSysAccObj, "", "c:\temp\" ', lobjSpec
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "初始化报表查询器失败。“统计报表”操作不可用。若要使用报表，请退出系统，并重新启动机器后执行系统安装程序。"
    End If
    
    mobjFrontQueryManager.当前用户 = um用户编号
    Exit Sub
errHandler:
    Set mobjFrontQueryManager = Nothing
    Set mobjSysAccObj = Nothing
    Call sfsub错误处理("主程序", "frm主界面", "sub初始化操作列表", Err.Number, Err.Description, True)
End Sub

'功能：启动报表查询界面。（用户按“统计报表”操作菜单时调用）
Private Sub sub启动报表查询()
    Dim lobjRec As Object
    Dim lcolReports As New Collection
    Dim lcolItem As Collection
    Dim lobjValueList As Object
    Dim lstrSql As String
    Dim i As Long
    
    On Error GoTo errHandler
    If mobjFrontQueryManager Is Nothing Then
        '再次初始化。
        sub初始化报表查询对象
    Else

        If mobjFrontQueryManager.ReportDataObject Is Nothing Then
            sub初始化报表查询对象
        End If
    End If


    '启动查询界面。
    Dim llng窗体句柄 As Long
    Dim llngWndProc As Long
    
    llng窗体句柄 = mobjFrontQueryManager.funcStart(pstr子系统许可)
    
    '设定打开的窗体为主窗体的子窗体。
    If llng窗体句柄 <> -2 Then
        '向集合中加入操作名称
        If Not sffunc判断集合键值是否存在(pcol操作名称, CStr(llng窗体句柄)) Then
            On Error Resume Next
            SetParent llng窗体句柄, Me.hWnd
            llngWndProc = SetWindowLong(llng窗体句柄, GWL_WNDPROC, AddressOf funcClassing)
            pcolWndProc.Add llngWndProc, CStr(llng窗体句柄)
            pcol操作名称.Add "报表统计", CStr(llng窗体句柄)
            pcol子窗体句柄.Add llng窗体句柄, "报表统计"
            
            Call MoveWindow(llng窗体句柄, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
            'Call MoveWindow(llng窗体句柄, ScaleX(1700, vbTwips, vbPixels), ScaleX(60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 380, vbTwips, vbPixels), 1)

            Err.Clear
            On Error GoTo errHandler
        End If
    End If
    
    Exit Sub
errHandler:
    Call sfsub错误处理("主程序", "frm主界面", "sub启动报表查询", Err.Number, Err.Description, True)
    Exit Sub
    Resume
End Sub

'功能：创建用户的窗体
Public Sub sub创建窗体(ByVal para业务名称 As String)
    On Error GoTo errHandle
    Dim lobj界面 As Object '启动子窗体的界面对象
    Dim lobj权限 As Object  '当前用户可用权限
    Set lobj权限 = um操作权限
    Dim llng窗体句柄 As Long          '当前活动子窗体
    Dim llngWndProc As Long
    Dim lstr业务名 As String
    If mblnLoadForm Then Exit Sub
    mblnLoadForm = True
   
    lobj权限.Filter = "权限名" & "= '" & para业务名称 & "'"  '比较用户的权限是否能操作该操作
    mobj操作.Filter = 0
    mobj操作.Filter = "操作名称" & "= '" & para业务名称 & "'"
    If lobj权限.RecordCount > 0 Then '有权进行该项操作
        '创建业务对象
        um当前操作子系统名 = mobj操作("业务名")
        If InStr(para业务名称, "健康证管理") > 0 Then
            Dim lobjFrm As New clsManageForm
            If lobjFrm.funcCheck("aaf3") <> "ab!&d3290" Then
                MsgBox "软件版本不正确！", vbCritical, "系统提示"
                Exit Sub
            End If
            llng窗体句柄 = lobjFrm.funcStart(para业务名称)
        Else
            Set lobj界面 = CreateObject(mobj操作("部件名") & "." & mobj操作("类名"))
            
            If para业务名称 = "系统管理_报表查询权限设置" Then
                llng窗体句柄 = lobj界面.funcStart(para业务名称, pstr子系统许可)
            Else
                llng窗体句柄 = lobj界面.funcStart(para业务名称)
            End If
        End If
        If llng窗体句柄 = -1 Then Err.Raise 6666, , "操作名称设定错误！未找到该操作名称所对应的窗体！"
        '设定打开的窗体为主窗体的子窗体。
        If llng窗体句柄 <> -2 Then
            '向集合中加入操作名称
            If Not sffunc判断集合键值是否存在(pcol操作名称, CStr(llng窗体句柄)) Then
                On Error Resume Next
                lstr业务名 = mobj操作("业务名")
                SetParent llng窗体句柄, Me.hWnd
                llngWndProc = SetWindowLong(llng窗体句柄, GWL_WNDPROC, AddressOf funcClassing)
                pcolWndProc.Add llngWndProc, CStr(llng窗体句柄)
                pcol业务名称.Add lstr业务名, CStr(llng窗体句柄)
                pcol子窗体句柄.Add llng窗体句柄, para业务名称
                pcol操作名称.Add para业务名称, CStr(llng窗体句柄)
                
                Call MoveWindow(llng窗体句柄, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)

                Err.Clear
                On Error GoTo errHandle
                Call oesubSave("用户进入" & para业务名称, "进入操作")
            End If
        End If
        Err.Clear
        On Error GoTo errHandle
    Else    '无权进行该项操作
        Call sffuncMsg("无权限进行该项操作", sf警告)
    End If
errHandle:
    mblnLoadForm = False
    Set lobj界面 = Nothing
    Set lobj权限 = Nothing
    If Err.Number = 0 Then Exit Sub
    If Err.Number = 429 Then
        Err.Number = 6666
        Err.Description = "该操作未在本机正确安装或注册！"
    End If
    Call sfsub错误处理("主程序", "frm主界面", "sub创建窗体", Err.Number, Err.Description, False)
End Sub


'功能：将错误信息写进数据中
'输入： 无
'输出： 无
'返回： 无
'注意事项：无
Private Sub WriteErrLog()
    On Error GoTo errHandler
    Dim lstr用户编号 As String        '用户编号
    Dim lstr工作站编号 As String      '工作站编号
    Dim ldat日期  As Date            '错误产生日期
    Dim lstr错误号  As String
    Dim lstr错误描述 As String
    Dim lstr错误产生路径 As String     '错误产生路径
    Dim lstrSql As String
    Dim lstrInput As String
    If Dir("C:\ErrLog") = "" Then Exit Sub  '错误记录为空则退出
    lstr用户编号 = um用户编号              '取用户编号
    lstr工作站编号 = um工作站编号         '取工作站编号
    Open "C:\ErrLog" For Input As #1     '打开错误记录本
        Do While Not EOF(1)
            Line Input #1, lstrInput
            lstr错误号 = Mid(lstrInput, InStr(1, lstrInput, "错误号：") + 4, InStr(1, lstrInput, "错误描述：") - InStr(1, lstrInput, "错误号：") - 5)
            lstr错误描述 = Mid(lstrInput, InStr(1, lstrInput, "错误描述：") + 5, InStr(1, lstrInput, "错误产生路径：") - InStr(1, lstrInput, "错误描述：") - 6)
            Do While InStr(1, lstr错误描述, "'")
                lstr错误描述 = Left(lstr错误描述, InStr(1, lstr错误描述, "'") - 1) & "`" & Right(lstr错误描述, Len(lstr错误描述) - InStr(1, lstr错误描述, "'"))
            Loop
            lstr错误描述 = LeftB(lstr错误描述, 500)
            If lstr错误描述 <> "" Then
                lstr错误描述 = Replace(lstr错误描述, "'", "''")     '将错误描述中出现的"'"转化成"''"
            End If
            lstr错误产生路径 = Mid(lstrInput, InStr(1, lstrInput, "错误产生路径：") + 7, InStr(1, lstrInput, "日期：") - InStr(1, lstrInput, "错误产生路径：") - 8)
            ldat日期 = Format(Mid(lstrInput, InStr(1, lstrInput, "日期：") + 3, Len(lstrInput) - InStr(1, lstrInput, "日期：")), "yyyy/mm/dd hh:mm:ss")
            '写入数据库
            lstrSql = "Insert Into 系统管理_系统错误记录表 Values('" & _
            lstr用户编号 & "' ,'" & _
            lstr工作站编号 & "','" & _
            ldat日期 & "','" & _
            lstr错误号 & "','" & _
            lstr错误描述 & "','" & _
            lstr错误产生路径 & "')"
            dafuncGetData (lstrSql)
        Loop
    Close #1
    Kill "C:\ErrLog"
    Exit Sub
errHandler:
    If Err.Number <> 3000 Then
        Resume Next
    Else
        Close #1
        Kill "C:\ErrLog"
    End If
End Sub


Private Sub sub注册报表查询部件()
    Dim lstrPath As String
    Dim lstrFile As String
    Dim llngRes As Long
    Dim lstrLongPath As String
    Dim lstrShortPath As String
    
    On Error Resume Next
    
    '把长路径转换为短路径。
    lstrLongPath = App.Path & "\公用组件\"
    lstrPath = String$(165, 0)
    
    lstrFile = lstrLongPath & "数据服务器.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
    
    lstrFile = lstrShortPath & "FileToDatabase.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
    
    lstrFile = lstrShortPath & "报表查询部件.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
    
    lstrFile = lstrShortPath & "道源报表查询器.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If (X > 3500 Or X < 1800) And Y > 900 Then
'        cfrm字典.Visible = False
'    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnLoadForm Then Cancel = True: Exit Sub
    On Error Resume Next
    Dim llng窗体句柄 As Long
    Dim lstrTemp As String
'    Me.WindowState = 0
    If mblnAutoUpgrade = False Then     '只有正常进行退出时才弹出该界面，自动升级不需要
        frmExit.Show vbModal, Me
        If pblnCancel Then
            Cancel = True
            Exit Sub
        End If
    Else
        pblnExit = True
    End If
    
    sub退出信使服务
    
    Dim i As Integer
    Dim lstr操作名 As String
    Dim lobj界面 As Object
    Dim lint操作数量 As Integer
    lint操作数量 = pcol操作名称.Count
    For i = 1 To lint操作数量
        lstr操作名 = pcol操作名称(1)
        If lstr操作名 <> "平台设置" Then
            If lstr操作名 = "字典管理" Then
                Set lobj界面 = CreateObject("字典管理.clsalldictionarys")
        
            ElseIf lstr操作名 = "报表统计" Then

                Set lobj界面 = mobjFrontQueryManager
            ElseIf lstr操作名 = "查询统计" Then

                Set lobj界面 = CreateObject("通用查询.cls通用查询")
            Else
                mobj操作.Filter = "操作名称" & "= '" & lstr操作名 & "'"
                '创建业务对象
                Set lobj界面 = CreateObject(mobj操作("部件名") & "." & mobj操作("类名"))
            End If
            On Error GoTo Goon
            lobj界面.funcClose lstr操作名
            If sffunc判断集合键值是否存在(pcol子窗体句柄, lstr操作名) Then
                Cancel = True
                Exit For
            End If
        
        Else
            'Unload frm平台设置
            If sffunc判断集合键值是否存在(pcol子窗体句柄, lstr操作名) Then
                Cancel = True
                Exit For
            End If
        End If
Goon:
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    Next i
    If Cancel = True Then
        mblnRe = False
        Exit Sub
    Else
        mblnRe = True
        Set pobj平台结构 = Nothing
    End If
'    WriteErrLog                                '将错误信息写入数据库
    Call oesubSave("用户退出系统", "退出系统") '记录操作日志
    SetWindowLong Me.hWnd, GWL_WNDPROC, pcolWndProc(CStr(Me.hWnd))
    If Cancel <> True Then
        Me.Hide
        Set pcolWndProc = Nothing
        Set pcol操作名称 = Nothing
        Set pcol业务名称 = Nothing
        Set pcol子窗体句柄 = Nothing
        Set mobjFrontQueryManager = Nothing
        If Not pblnExit And mblnRe Then
            pbln注销 = True
            Call oesubSave("用户注销重新进入系统", "注销")
'            Unload frm短信通知
'            Unload frm后勤管理
            Unload frm系统管理
            Unload frm字典列表
            Call Main
        Else

            subExit
            
        End If
    End If
End Sub
Private Sub subExit()
    On Error Resume Next
    X.subCloseDatabase
'    Unload frm短信通知
'    Unload frm后勤管理
    Unload frm系统管理
    Unload frm字典列表
End Sub
Private Sub Form_Resize()
    On Error Resume Next
'    Image1.Left = 0
'    Image1.Top = 0
'    Image1.Width = Me.ScaleWidth
'    Image1.Height = Me.ScaleHeight
''    clblClose.Left = Me.ScaleWidth - 375
'    Image1.ZOrder 1
'    clblInfo.Top = Me.ScaleHeight - 500
    Frame1.Height = Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 80
    Image5.Top = Frame1.Height - Image5.Height - 100
    cimgBackground.Width = Me.ScaleWidth - cimgBackground.Left
    Frame2.Top = Me.ScaleHeight - cstatusBar.Height - Frame2.Height - 120
    Frame2.Left = Me.ScaleWidth - Frame2.Width - 120
    cimgPhoto.Top = Frame1.Height - cimgPhoto.Height - 100
    
    subResizeChild
End Sub
Private Sub sub初始化字典列表(para组名 As String)
    Dim i As Integer, j As Integer
    Dim lobjRec As Object
    
    On Error GoTo errHandle
    
    mobj组.Filter = ""
    mobj组.Filter = "所属组名" & " ='" & para组名 & "' "
    For i = 1 To mobj组.RecordCount
        '修改：判断当前操作所属业务名是否在加密狗许可范围内。
        mobj操作.Filter = ""
        mobj操作.Filter = "操作名称" & "='" & mobj组.Fields("操作名称") & "'"
        If mobj操作.RecordCount > 0 Then
            If pstr子系统许可 = "" Or InStr(pstr子系统许可, mobj操作.Fields("业务名") & ",") > 0 Then
                If Not sffunc判断集合键值是否存在(pcol字典集, mobj操作.Fields("业务名").Value) Then
                    '判断该业务是否有操作级的字典。
                    Set lobjRec = dafuncGetData("select * from 系统管理_字典_字典表列表 where 业务名='" & mobj操作.Fields("业务名").Value & "' and 级别='操作级'")
                    If lobjRec.RecordCount > 0 Then
                        pcol字典集.Add mobj操作.Fields("业务名").Value, mobj操作.Fields("业务名").Value
                    End If
                End If
            End If
        End If
        mobj组.MoveNext
    Next i

    Exit Sub
errHandle:
    If Err.Number = 40002 Or Err.Number = 40003 Or Err.Number = 40006 Then
        Resume Next
    Else
        Call sfsub错误处理("主程序", "frm主界面", "sub初始化操作列表", Err.Number, Err.Description, False)
    End If
End Sub


Private Sub sub初始化操作列表(para组名 As String)
    Dim i As Integer, j As Integer
    Dim ii As Integer
    Dim lobjRec As Object
    On Error GoTo errHandle
    
    '加入该操作组的操作
    ii = 1
    mobj组.Filter = ""
    mobj组.Filter = "所属组名" & " ='" & para组名 & "' "
    For i = 1 To 10
        cmnuSubItem(i).Visible = False
        cimgButton(i).Visible = False
        cmnuSubItem(i) = ""
    Next

    For i = 1 To mobj组.RecordCount
        '修改：判断当前操作所属业务名是否在加密狗许可范围内。
        mobj操作.Filter = ""
        mobj操作.Filter = "操作名称" & "='" & mobj组.Fields("操作名称") & "'"
        If mobj操作.RecordCount > 0 Then
            If pstr子系统许可 = "" Or InStr(pstr子系统许可, mobj操作.Fields("业务名") & ",") > 0 Then
                cmnuSubItem(ii) = IIf(Len(mobj组("操作别名")) > 6, Left(mobj组("操作别名"), 6) & "...", mobj组("操作别名"))
                cmnuSubItem(ii).Top = cmnuSubItem(ii - 1).Top + cmnuSubItem(ii - 1).Height + 150
                cmnuSubItem(ii).Left = cmnuSubItem(0).Left
                cmnuSubItem(ii).Visible = True
                cimgButton(ii).Top = cmnuSubItem(ii).Top - 30
                cimgButton(ii).Visible = True
                cimgButton(ii).Left = cimgButton(0).Left
                mstrOper(ii) = mobj组.Fields("操作名称")
                ii = ii + 1
            
            End If
        End If
        mobj组.MoveNext
    Next i

    Exit Sub
errHandle:
    If Err.Number = 40002 Or Err.Number = 40003 Or Err.Number = 40006 Then
        Resume Next
    Else
        Call sfsub错误处理("主程序", "frm主界面", "sub初始化操作列表", Err.Number, Err.Description, False)
    End If
End Sub

Public Sub sub设置字典(ByVal para子系统名 As String)
    On Error Resume Next
'    ctxtSmartInfos.Caption = ""   '清除主推信息
    On Error GoTo errHandle
    Dim lobj界面 As Object '启动子窗体的界面对象
    Dim llng窗体句柄 As Long          '当前活动子窗体
    Dim llngWndProc As Long
    '创建业务对象
    
    um当前操作子系统名 = para子系统名 'clbl字典(Index).Caption
    Set lobj界面 = CreateObject("字典管理.clsalldictionarys")
    llng窗体句柄 = lobj界面.funcStart(para子系统名)
    If llng窗体句柄 = -1 Then Err.Raise 6666, , "操作名称设定错误！未找到该操作名称所对应的窗体！"
    '设定打开的窗体为主窗体的子窗体。
    If llng窗体句柄 <> -2 Then
        '向集合中加入操作名称
        If Not sffunc判断集合键值是否存在(pcol操作名称, CStr(llng窗体句柄)) Then
            On Error Resume Next
            SetParent llng窗体句柄, Me.hWnd
            llngWndProc = SetWindowLong(llng窗体句柄, GWL_WNDPROC, AddressOf funcClassing)
            pcolWndProc.Add llngWndProc, CStr(llng窗体句柄)
            pcol操作名称.Add "字典管理", CStr(llng窗体句柄)
            pcol子窗体句柄.Add llng窗体句柄, "字典管理"
            Call MoveWindow(llng窗体句柄, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
            
            Err.Clear
            On Error GoTo errHandle
            Call oesubSave("用户进入字典管理", "进入操作")
        End If
    End If
errHandle:
    Set lobj界面 = Nothing
    If Err.Number = 0 Then Exit Sub
    If Err.Number = 429 Then
        Err.Number = 6666
        Err.Description = "该操作未在本机正确安装或注册！"
    End If
    Call sfsub错误处理("主程序", "frm主界面", "sub设置字典", Err.Number, Err.Description, False)
End Sub


Private Sub sub检查试用期限()
    Dim lstrTime As String
    
    lstrTime = "2008-12-31"
    
    If lstrTime < Format(Now, "yyyy-mm-dd") Then
        MsgBox "对不起，你的使用期限已到，请与软件供应商联系。", vbCritical, "系统提示"
        End
    End If
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 1 To mintMnu
        cmnuItem(i).FontUnderline = False
        cmnuItem(i).ForeColor = vbBlack
    Next
    For i = 1 To 10
        cmnuSubItem(i).FontUnderline = False
        cmnuSubItem(i).ForeColor = vbBlack
    Next
End Sub

Private Sub Timer1_Timer()
    sub检查试用期限
End Sub


Public Sub subResizeChild()
    Dim llngHwnd As Long
    Dim i As Long
    
    For i = 1 To pcol子窗体句柄.Count
        llngHwnd = pcol子窗体句柄(i)

        Call MoveWindow(llngHwnd, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
        'Call MoveWindow(llngHwnd, ScaleX(700, vbTwips, vbPixels), ScaleX(350, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 400, vbTwips, vbPixels), 1)

    Next
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    Dim lforeWin As Long
    '检查用户的有效期限
    Dim lstrDate As String, lstrExpireDate As String
    Dim lobjCheck As New cls用户检查

    lstrDate = (dafuncGetData("select getdate()").Fields(0))
    lstrExpireDate = lobjCheck.funcGetExpireDate()
    If lstrExpireDate = "" Then
        MsgBox "您不是系统的正式用户，系统无法运行！", vbCritical, "系统提示"
        End
    End If
    If lstrExpireDate = "认证信息错误" Then
        MsgBox "系统的认证信息错误，无法继续运行！", vbCritical, "系统提示"
        End
    End If
    If CDate(lstrDate) > CDate(lstrExpireDate) Then
        MsgBox "当前系统已经失效，无法运行！", vbCritical, "系统提示"
        End
    End If
    
'    lforeWin = GetForegroundWindow()
'    If Me.hWnd = lforeWin Then
'
'        sub显示代办工作
'
'    End If
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    On Error Resume Next
    Timer3.Enabled = False
    sub登录信使服务
End Sub
