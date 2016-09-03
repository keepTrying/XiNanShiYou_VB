VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "卫生防疫管理信息系统"
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   15
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00CEE0AF&
      BorderStyle     =   0  'None
      Height          =   7515
      Left            =   0
      TabIndex        =   9
      Top             =   615
      Width           =   1755
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "可 用 操 作"
         Height          =   270
         Left            =   450
         TabIndex        =   22
         Top             =   75
         Width           =   1035
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   -15
         Picture         =   "frmMain.frx":0CCA
         Stretch         =   -1  'True
         Top             =   -15
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         Index           =   7
         Left            =   630
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   11
         Top             =   360
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
         Left            =   405
         TabIndex        =   10
         Top             =   825
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar cstatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   11160
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "用户编号："
            TextSave        =   "用户编号："
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "用户姓名："
            TextSave        =   "用户姓名："
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "工作站名："
            TextSave        =   "工作站名："
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   1800
      Top             =   1200
   End
   Begin VB.Image Image4 
      Height          =   600
      Left            =   30
      Picture         =   "frmMain.frx":1365
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "重庆市九龙坡区卫生局卫生监督所管理信息系统"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   660
      TabIndex        =   24
      Top             =   150
      Width           =   6300
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帮助"
      Height          =   180
      Index           =   0
      Left            =   14445
      TabIndex        =   23
      Top             =   225
      Width           =   360
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统管理"
      Enabled         =   0   'False
      Height          =   180
      Index           =   4
      Left            =   10530
      TabIndex        =   8
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "后勤管理"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   8385
      TabIndex        =   7
      Top             =   225
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   8190
      Picture         =   "frmMain.frx":2783
      Stretch         =   -1  'True
      Top             =   9960
      Visible         =   0   'False
      Width           =   7035
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户编号：           用户名称：           工作站名："
      Height          =   180
      Left            =   2700
      TabIndex        =   5
      Top             =   7530
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "退出"
      Height          =   180
      Index           =   7
      Left            =   13740
      TabIndex        =   4
      Top             =   225
      Width           =   360
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "口令修改"
      Height          =   180
      Index           =   6
      Left            =   12660
      TabIndex        =   3
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "短信通知"
      Enabled         =   0   'False
      Height          =   180
      Index           =   3
      Left            =   9465
      TabIndex        =   2
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查    询"
      Height          =   180
      Index           =   1
      Left            =   7320
      TabIndex        =   1
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clbl通用操作 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字典管理"
      Height          =   180
      Index           =   5
      Left            =   11595
      TabIndex        =   0
      Top             =   225
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   6990
      Left            =   60
      Picture         =   "frmMain.frx":14861
      Stretch         =   -1  'True
      Top             =   1695
      Width           =   15135
   End
   Begin VB.Image cimgBackground 
      Height          =   705
      Left            =   15
      Picture         =   "frmMain.frx":24161
      Stretch         =   -1  'True
      Top             =   -75
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

'修改：2002-2-26（新增对象）。
Private X As Object

'报表查询需要的变量。
'修改：2001-7-12。
Private mobjFrontQueryManager As Object '道源报表查询器.clsFrontQueryManager。
Private mobjSysAccObj As Object         '数据驱动器.clsSystemAccessObject。

Private Sub cimg组_Click(Index As Integer)
'    Dim i As Long
'    Dim lobjSys As New FileSystemObject
'    Dim llngCount As Long
'    Dim llngTop As Long
'
'    If Index = 0 Then Exit Sub
'    On Error Resume Next
'
'    Unload frm字典列表
'
'    If lobjSys.FileExists(App.Path & "\image\" & cimg组(Index).Tag & "2.jpg") Then
'        cimg组(Index).Picture = LoadPicture(App.Path & "\image\" & cimg组(Index).Tag & "2.jpg")
'    Else
'        cimg组(Index).Picture = LoadPicture(App.Path & "\image\" & "hot.jpg")
'    End If
'
'    llngCount = cimg组.Count
'
'    For i = 1 To llngCount - 1
'        If i <> Index Then
'            If lobjSys.FileExists(App.Path & "\image\" & cimg组(i).Tag & "1.jpg") Then
'                cimg组(i).Picture = LoadPicture(App.Path & "\image\" & cimg组(i).Tag & "1.jpg")
'            Else
'                cimg组(i).Picture = LoadPicture(App.Path & "\image\" & "normal.jpg")
'
'            End If
'        End If
'    Next
'
'    sub初始化操作列表 cimg组(Index).Tag
'
'    Set frm操作列表.pfrmParent = Me
'
'    If frm操作列表.clbl操作.Count = 2 Then
'        '只有一个操作，直接启动操作界面。
'        Call sub创建窗体(frm操作列表.clbl操作(1).Tag)
'    Else
'        '显示操作选择列表。
'        frm操作列表.Height = frm操作列表.clbl操作.Count * (frm操作列表.clbl操作(0).Height + 100) + 200
'        llngTop = cimg组(Index).Top
''        If llngTop + frm操作列表.Height > Me.ScaleHeight - 200 Then
''            llngTop = ScaleHeight - frm操作列表.Height - 200
''        End If
''        If llngTop < 720 Then
''            llngTop = 720
''        End If
'        llngTop = llngTop + Me.Top + 200
'        frm操作列表.Move Me.Left + cimg组(0).Width + 300, llngTop '720
'        frm操作列表.Show , Me
'    End If
'
End Sub

'Public Enum Endway                      '定义结束方式
'    CloseAll = 1                        '退出系统
'    Restart = 2                         '重新登录
'End Enum

Private Sub cimgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To 7
        clbl通用操作(i).FontUnderline = False
        clbl通用操作(i).ForeColor = vbBlack
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
    clbl通用操作(Index).ForeColor = vbBlack
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
         Case "后勤管理"
                frm后勤管理.Hide
                Set frm后勤管理.pobjParent = Me
                frm后勤管理.Move clbl通用操作(Index).Left, Me.Top + 800
                frm后勤管理.Show , Me
         Case "短信通知"
                frm短信通知.Hide
                Set frm短信通知.pobjParent = Me
                frm短信通知.Move clbl通用操作(Index).Left, Me.Top + 800
                frm短信通知.Show , Me
         Case "系统管理"
                frm系统管理.Hide
                Set frm系统管理.pobjParent = Me
                frm系统管理.Move clbl通用操作(Index).Left, Me.Top + 800
                frm系统管理.Show , Me
           
        Case "查    询"      '选取该平台所有查询
                '启动查询界面。
                Dim lobj通用查询 As Object
                Set lobj通用查询 = CreateObject("通用查询.cls通用查询")
                '修改：2003-7-22（杨春）增加子系统许可参数。
                llng窗体句柄 = lobj通用查询.funcStart("系统管理_通用查询", pstr子系统许可)
                
                '设定打开的窗体为主窗体的子窗体。
                If llng窗体句柄 <> -2 Then
                    '向集合中加入操作名称
                    If Not sffunc判断集合键值是否存在(pcol操作名称, CStr(llng窗体句柄)) Then
                        On Error Resume Next
                        SetParent llng窗体句柄, Me.hWnd
                        llngWndProc = SetWindowLong(llng窗体句柄, GWL_WNDPROC, AddressOf funcClassing)
                        pcolWndProc.add llngWndProc, CStr(llng窗体句柄)
                        pcol操作名称.add "查询统计", CStr(llng窗体句柄)
                        pcol子窗体句柄.add llng窗体句柄, "查询统计"
'                        Call MoveWindow(llng窗体句柄, ScaleX(700, vbTwips, vbPixels), ScaleX(350, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 350, vbTwips, vbPixels), 1)
                        Call MoveWindow(llng窗体句柄, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
                        
                        '********************临时测试代码********************
                        'sfsubSaveSetting "系统管理", "平台问题测试", "查询统计初始化", "hWnd：" & llng窗体句柄 & " WndProc：" & llngWndProc & " 时间" & Format(Now, "yyyy年mm月dd日hh时mm分ss秒")
                        '****************************************************
                        Err.Clear
                    End If
                End If
        Case "统计报表"       '选取该平台所有报表
'                mobj报表.Filter = ""
'                If mobj报表.RecordCount = 0 Then Exit Sub
'                mobj报表别名.Filter = "操作名称" & "='" & mobj报表.Fields("操作名称") & "'"
                
                '启动报表查询界面。
                '修改：2001-7-13（杨春）
                Call sub启动报表查询
                
        Case "口令修改"
            frm密码修改.Show vbModal, Me
        Case "退出"
            Unload Me
        Case "帮助"
            MsgBox "正在制作中！", vbOKOnly, "帮助"
        
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
    
    For ii = 1 To Index
        cmnuItem(ii).Top = cmnuItem(ii - 1).Top + cmnuItem(ii - 1).Height + 150
    Next
    cmnuSubItem(0).Top = cmnuItem(ii - 1).Top '+ cmnuItem(ii - 1).Height + 150
    sub初始化操作列表 mstrMnu(Index)
    For j = 1 To 10
        If cmnuSubItem(j) = "" Then Exit For
    Next
    If ii < mintMnu Then
        cmnuItem(ii).Top = cmnuSubItem(j - 1).Top + cmnuSubItem(j - 1).Height + 150
        For ii = ii + 1 To mintMnu
            cmnuItem(ii).Top = cmnuItem(ii - 1).Top + cmnuItem(ii - 1).Height + 150
        Next
    End If
    cmnuItem(Index).FontUnderline = False
    cmnuItem(Index).ForeColor = vbBlack
End Sub

Private Sub cmnuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmnuItem(Index).FontUnderline = True
    cmnuItem(Index).ForeColor = &H80FF&
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
    cmnuSubItem(Index).FontUnderline = True
    cmnuSubItem(Index).ForeColor = &H80FF&
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
    
    '修改：2001-11-16（杨春）为了在笔记本上可以运行，只对网络运行的情况加密，单机不加密。
    On Error Resume Next
    Dim lstrServer As String
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    
    Me.Caption = um防疫站名 & "管理信息系统"
    
    '获取本机名称。
    Dim lstrLocalName As String
    lstrLocalName = funcGetLocalName()
    'clblInfo.Caption = "用户编号：" & um用户编号 & "       用户名称：" & um用户名 & "      工作站名：" & lstrLocalName
    cstatusBar.Panels(1).Text = "用户编号：" & um用户编号
    cstatusBar.Panels(2).Text = "用户姓名：" & um用户名
    cstatusBar.Panels(3).Text = "工作站名：" & lstrLocalName
    'clblSys.Caption = IIf(pstrSysName = "", "卫生防疫管理信息系统", pstrSysName)
    
    '修改：2002-3-8（若重新注销后，不需要再检查机密狗）。
    If Not pbln注销 Then
    '修改：2003-4-17（杨春）检查加密狗改为调用存储过程。
'    If Not pbln注销 And UCase(Trim(lstrServer)) <> UCase(Trim(lstrLocalName)) Then
'        '获取系统配置中的加密狗服务器名。
'        Dim lstrDogServer As String
'        lstrDogServer = sffuncGetSetting("系统管理", "数据库配置", "网络锁服务器名")
'        If lstrDogServer = "" Then lstrDogServer = lstrServer
'
'        Dim lobjRec As Object
'        Call dafuncGetData("sp_addlinkedserver '" & lstrDogServer & "'")
'        Err.Clear
'        Set lobjRec = dafuncGetData("exec " & lstrDogServer & ".master.dbo.ryCheck")
'        If Err.Number = 0 Then
'            Select Case lobjRec(0)
'            Case 1, 2, 3, 4, 5, 6
'                MsgBox "网络锁服务安装不正常，导致安全检查失败。" & Chr(13) & Chr(10) & "请重新安装网络锁服务程序。", vbCritical, "系统错误"
'                End
'            Case Else
'                If (lobjRec(0) And (&H8000)) = 32768 Then
'                    '找到了。获取许可。
'                    Dim llngBit As Long
'                    Dim larrLic(1 To 10) As String
'                    larrLic(1) = "卫生许可证管理"
'                    larrLic(2) = "监督执法管理"
'                    larrLic(3) = "体检管理"
'                    larrLic(4) = "健康证管理,健康证"
'                    larrLic(5) = "卫生监测管理"
'                    larrLic(6) = "检验管理"
'                    larrLic(7) = "计划免疫管理"
'                    larrLic(8) = "后勤管理"
'                    larrLic(9) = "收费管理"
'                    larrLic(10) = "站长查询"
'                    pstr子系统许可 = ""
'                    llngBit = &H4000
'                    For i = 1 To 10
'                        If (lobjRec(0) And llngBit) = llngBit Then
'                            pstr子系统许可 = pstr子系统许可 & larrLic(i) & ","
'                        End If
'                        llngBit = llngBit / 2
'                    Next
'                Else
'                    MsgBox "网络锁不是正式版的，安全检查失败！系统无法运行。" & Chr(13) & Chr(10) & "请与软件供应商联系。", vbCritical, "系统错误"
'                    End
'                End If
'            End Select
'        Else
'            MsgBox "网络锁服务安装不正常，导致安全检查失败。" & Chr(13) & Chr(10) & "请重新安装网络锁服务程序。安装前确保网络锁服务器上已安装Sql Server2000。", vbCritical, "系统错误"
'            End
'        End If
'
'        pstr子系统许可 = "系统管理,单位档案管理," & pstr子系统许可
        sub检查试用期限
        
    End If
    
    'pstr子系统许可 = "系统管理,单位档案管理,体检管理,卫生许可证管理,健康证管理,健康证,"
    
    '修改：2003-7-9（杨春）在通用对象的全局变量中保存子系统徐可。
    um子系统许可 = pstr子系统许可
    
    dasubSetQueryTimeout 6000

    On Error GoTo errHandle
    
    Dim llngWndProc As Long
    llngWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf funcClassing)
    pcolWndProc.add llngWndProc, CStr(Me.hWnd)
    
    pobj平台结构.平台名称 = um用户编号 '给平台结构的用户编号赋值
    Set mobj类 = pobj平台结构.操作组分类  '取该用户的平台结构
    Set mobj组 = pobj平台结构.Operations
    Set mobj操作 = pobj平台结构.Operation
    Set mobj查询 = pobj平台结构.Queries
    Set mobj查询别名 = pobj平台结构.Query
    Set mobj报表 = pobj平台结构.Reports
    Set mobj报表别名 = pobj平台结构.Report
    Set mobjSmartInfos = pobj平台结构.SmartInfos
    
    '初始化操作组图片
'    mobj类.Filter = "所属类名= '业务类'"
    j = 1
    '加载图标库
    Dim lstrFileName As String, lobjPic As StdPicture
    
'    lstrFileName = Dir(App.Path & "\image\*.ico")
'    Do While lstrFileName <> ""
'        Set lobjPic = LoadPicture(App.Path & "\image\" & lstrFileName)
'        cbarOper.IconsLarge.add , Left(lstrFileName, InStr(lstrFileName, ".") - 1), lobjPic
'        lstrFileName = Dir
'    Loop
    
    Dim lcolOper As New Collection, lstrOper(1 To 14) As String, lstrAlias(1 To 14) As String
    
    lcolOper.add 8, "单位档案"
    lcolOper.add 1, "卫生许可证"
    lcolOper.add 3, "监督执法"
    lcolOper.add 4, "稽查"
    lcolOper.add 6, "食卫管理员"
    lcolOper.add 5, "医疗机构"
    lcolOper.add 2, "健康证"
    lcolOper.add 7, "收费"
    'lcolOper.add 9, "后勤"
    'lcolOper.add 10, "短信通知"
    lcolOper.add 9, "文件"
    lcolOper.add 10, "领导查询"
    'lcolOper.add 13, "系统"
    
    For ii = 1 To mobj类.RecordCount
        '判断该组是否有操作。
        'cbarOper.Groups.add ii + 1, , mobj类("操作组")
        Select Case mobj类("操作组")
            Case "查询"
                'lstrOper(lcolOper(mobj类("操作组"))) = "后勤管理"
                clbl通用操作(1).Enabled = True
            Case "后勤"
                'lstrOper(lcolOper(mobj类("操作组"))) = "后勤管理"
                clbl通用操作(2).Enabled = True
                frm后勤管理.subLoad mobj类("操作组"), mobj组, mobj操作
            Case "短信通知"
                'lstrOper(lcolOper(mobj类("操作组"))) = "后勤管理"
                clbl通用操作(3).Enabled = True
                frm短信通知.subLoad mobj类("操作组"), mobj组, mobj操作
            Case "系统"
                'lstrOper(lcolOper(mobj类("操作组"))) = "系统管理"
                clbl通用操作(4).Enabled = True
                frm系统管理.subLoad mobj类("操作组"), mobj组, mobj操作
                'sub初始化操作列表 frm后勤管理, mobj类("操作组")
            Case "字典"
                'lstrOper(lcolOper(mobj类("操作组"))) = "后勤管理"
                clbl通用操作(5).Enabled = True
            Case "稽查"
                lstrOper(lcolOper(mobj类("操作组"))) = "稽查管理"
                lstrAlias(lcolOper(mobj类("操作组"))) = mobj类("操作组")
            Case "健康证"
                lstrOper(lcolOper(mobj类("操作组"))) = "健 康 证"
                lstrAlias(lcolOper(mobj类("操作组"))) = mobj类("操作组")
            Case "收费"
                lstrOper(lcolOper(mobj类("操作组"))) = "收费管理"
                lstrAlias(lcolOper(mobj类("操作组"))) = mobj类("操作组")
            Case "文件"
                lstrOper(lcolOper(mobj类("操作组"))) = "文件管理"
                lstrAlias(lcolOper(mobj类("操作组"))) = mobj类("操作组")
            Case Else
                lstrOper(lcolOper(mobj类("操作组"))) = mobj类("操作组")
                lstrAlias(lcolOper(mobj类("操作组"))) = mobj类("操作组")
        End Select
        mobj类.MoveNext
    Next
    mintMnu = 0
    For ii = 1 To lcolOper.Count
        If lstrOper(ii) <> "" Then
            'cbarOper.Groups.Add ii + 1, , "　· " + lstrOper(ii)
            Load cmnuItem(ii)
            cmnuItem(ii) = lstrOper(ii)
            cmnuItem(ii).Left = cmnuItem(0).Left
            cmnuItem(ii).Top = cmnuItem(ii - 1).Top + cmnuItem(ii - 1).Height + 150
            cmnuItem(ii).Visible = True
            mstrMnu(ii) = lstrAlias(ii)
            mintMnu = mintMnu + 1
            sub初始化字典列表 lstrAlias(ii)
        End If
    Next
    
    'cbarOper.Groups.Remove cbarOper.Groups(1)
    
'    frm字典列表.subLoad

    '如果系统不提供查询，则 查询统计 按钮不可用
    Dim lobjRec As Object
'    Set lobjRec = dafuncGetData("select * from 系统管理_查询信息表")
'    If lobjRec.RecordCount = 0 Then
'        clbl通用操作(2).Visible = False
'        clbl通用操作(5).Left = clbl通用操作(4).Left
'        clbl通用操作(4).Left = clbl通用操作(3).Left
'        clbl通用操作(3).Left = clbl通用操作(2).Left
''    Else
''        cbarOper.Groups.add
''        cbarOper.Groups(cbarOper.Groups.Count).Caption = "查    询"
''        cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
''        cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "查询"
''        cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).IconLarge = "查询"
'   End If
    '如果系统不提供报表，则报表统计 按钮不可用
    Set lobjRec = dafuncGetData("select * from 报表基本信息表")
    If lobjRec.RecordCount = 0 Then
'        clbl通用操作(3).Visible = False
'        clbl通用操作(5).Left = clbl通用操作(4).Left
'        clbl通用操作(4).Left = clbl通用操作(3).Left
    Else
'        '修改：2001-7-13（杨春）。
'        cbarOper.Groups.add
'        cbarOper.Groups(cbarOper.Groups.Count).Caption = "　· 统计报表"
'        cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'        cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "统计报表"
        mintMnu = mintMnu + 1
        Load cmnuItem(mintMnu)
        cmnuItem(mintMnu) = "统计报表"
        cmnuItem(mintMnu).Left = cmnuItem(0).Left
        cmnuItem(mintMnu).Top = cmnuItem(mintMnu - 1).Top + cmnuItem(mintMnu - 1).Height + 150
        cmnuItem(mintMnu).Visible = True
        mstrMnu(mintMnu) = "统计报表"
'        sub初始化报表查询对象
    End If
    
   
'    If pcol字典集.Count Then
'        cbarOper.Groups.add
'        cbarOper.Groups(cbarOper.Groups.Count).Caption = "字典设置"
'        For ii = 1 To pcol字典集.Count
'            cbarOper.Groups(cbarOper.Groups.Count).ListItems.add ii, "字典" & pcol字典集(ii)
'            cbarOper.Groups(cbarOper.Groups.Count).ListItems(ii).Text = pcol字典集(ii)
'            cbarOper.Groups(cbarOper.Groups.Count).ListItems(ii).IconLarge = pcol字典集(ii)
'       Next
'    End If
'    cbarOper.Groups.add
'    cbarOper.Groups(cbarOper.Groups.Count).Caption = "修改口令"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "修改口令"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).IconLarge = "修改口令"
'    cbarOper.Groups.add
'    cbarOper.Groups(cbarOper.Groups.Count).Caption = "退出/注销"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "退出"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).IconLarge = "退出"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(2).Text = "注销"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(2).IconLarge = "注销"
   
'    For ii = 1 To cbarOper.Groups.Count
'        For j = 1 To cbarOper.Groups(ii).ListItems.Count
'            cbarOper.Groups(ii).ListItems(j).ForeColorSource = ssUseListItem
'            cbarOper.Groups(ii).ListItems(j).ForeColor = &H0
'        Next
'    Next
    
    '修改：2002-1-10(系统管理员第一次运行，直接启动“运行状态设置”）。
    On Error Resume Next
    If um用户编号 = "0000" Then
        '判断是否第一次运行。
        Dim lobj记忆 As cls用户操作记忆
        Set lobj记忆 = New cls用户操作记忆
        lobj记忆.用户编号 = "0000"
        lobj记忆.业务名 = "系统管理"
        If lobj记忆.记忆项值("第一次运行") <> "否" Then
'            ctxtSmartInfos.Caption = ""   '清除主推信息
            Call sub创建窗体("系统管理_运行状态设置")
            '保存已运行果状态。
            lobj记忆.sub覆盖记忆值 "第一次运行", "否"
        End If
    End If
    
    'Me.Caption = pstrSysName ' & "（试用版）"
    
    On Error Resume Next
'    Image2.Picture = LoadPicture(App.Path & "\image\background(试用" & pstr版本代号 & ").jpg")
    pstrSysName = Me.Caption
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
'修改：2001-7-13（杨春）。
Private Sub sub初始化报表查询对象()

    On Error Resume Next
    
    Set mobjSysAccObj = CreateObject("数据服务器.clsSystemBaseAccess")
    If Err <> 0 Then
        '修改：2002-3-4（杨春）重新注册。
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
'修改：2001-7-13（杨春）。
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
        'Err.Raise 6666, , "报表查询初始化失败，不能启动报表查询界面。请退出系统，重新安装系统。"
    Else
    '修改：2001-8-15（杨春）。
        If mobjFrontQueryManager.ReportDataObject Is Nothing Then
            sub初始化报表查询对象
        End If
    End If
    
    '启动查询界面。
    Dim llng窗体句柄 As Long
    Dim llngWndProc As Long
    '修改：2003-7-22（杨春）增加子系统许可参数。
    llng窗体句柄 = mobjFrontQueryManager.funcStart(pstr子系统许可)
    
    '设定打开的窗体为主窗体的子窗体。
    If llng窗体句柄 <> -2 Then
        '向集合中加入操作名称
        If Not sffunc判断集合键值是否存在(pcol操作名称, CStr(llng窗体句柄)) Then
            On Error Resume Next
            SetParent llng窗体句柄, Me.hWnd
            llngWndProc = SetWindowLong(llng窗体句柄, GWL_WNDPROC, AddressOf funcClassing)
            pcolWndProc.add llngWndProc, CStr(llng窗体句柄)
            pcol操作名称.add "报表统计", CStr(llng窗体句柄)
            pcol子窗体句柄.add llng窗体句柄, "报表统计"
            
            Call MoveWindow(llng窗体句柄, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
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
'输入：操作名称
'输出：无
'返回：无
'注意事项：如果该窗体已创建则显示，未创建则调用各子系统的通用方法funcStart创建业务窗体
'作者：王晓华
'创建时间：2001-3-8
' 修改说明：由于窗体对象在不同进程间传递存在问题，故改为只返回窗体句柄。
' 修改人：  罗庆
' 修改时间：2001-3-20
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
        Set lobj界面 = CreateObject(mobj操作("部件名") & "." & mobj操作("类名"))
        
        If para业务名称 = "系统管理_报表查询权限设置" Then
            llng窗体句柄 = lobj界面.funcStart(para业务名称, pstr子系统许可)
        Else
            llng窗体句柄 = lobj界面.funcStart(para业务名称)
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
                pcolWndProc.add llngWndProc, CStr(llng窗体句柄)
                pcol业务名称.add lstr业务名, CStr(llng窗体句柄)
                pcol子窗体句柄.add llng窗体句柄, para业务名称
                pcol操作名称.add para业务名称, CStr(llng窗体句柄)
                
                Call MoveWindow(llng窗体句柄, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)

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
'作者：王晓华
'创建时间：2001-03-21
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
    frmExit.Show vbModal, Me
    If pblnCancel Then
        Cancel = True
        Exit Sub
    End If
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
                '修改：2001-7-13（杨春）
                Set lobj界面 = mobjFrontQueryManager
            ElseIf lstr操作名 = "查询统计" Then
                '修改：2001-7-24（罗庆）
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
            Unload frm平台设置
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
    WriteErrLog                                '将错误信息写入数据库
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
            Unload frm短信通知
            Unload frm后勤管理
            Unload frm系统管理
            Unload frm字典列表
            Call Main
        Else
            '修改：2002-2-26（退出）。
            subExit
        End If
    End If
End Sub
Private Sub subExit()
    On Error Resume Next
    X.subCloseDatabase
    Unload frm短信通知
    Unload frm后勤管理
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
    Frame1.Height = Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100
    cimgBackground.Width = Me.ScaleWidth - cimgBackground.Left
    subResizeChild
End Sub
Private Sub sub初始化字典列表(para组名 As String)
    Dim i As Integer, j As Integer
    Dim lobjRec As Object
    
    On Error GoTo errHandle
    
    mobj组.Filter = ""
    mobj组.Filter = "所属组名" & " ='" & para组名 & "' "
    For i = 1 To mobj组.RecordCount
        '修改：2003-7-9（杨春）判断当前操作所属业务名是否在加密狗许可范围内。
        mobj操作.Filter = ""
        mobj操作.Filter = "操作名称" & "='" & mobj组.Fields("操作名称") & "'"
        If mobj操作.RecordCount > 0 Then
            If pstr子系统许可 = "" Or InStr(pstr子系统许可, mobj操作.Fields("业务名") & ",") > 0 Then
                If Not sffunc判断集合键值是否存在(pcol字典集, mobj操作.Fields("业务名").Value) Then
                    '判断该业务是否有操作级的字典。
                    Set lobjRec = dafuncGetData("select * from 系统管理_字典_字典表列表 where 业务名='" & mobj操作.Fields("业务名").Value & "' and 级别='操作级'")
                    If lobjRec.RecordCount > 0 Then
                        pcol字典集.add mobj操作.Fields("业务名").Value, mobj操作.Fields("业务名").Value
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
    
    'frm操作列表.subClear

    '加入该操作组的操作
    'frm操作列表.clblTitle.Caption = para组名
    ii = 1
    mobj组.Filter = ""
    mobj组.Filter = "所属组名" & " ='" & para组名 & "' "
    For i = 1 To 10
        cmnuSubItem(i).Visible = False
        cmnuSubItem(i) = ""
    Next
    'cbarOper.Groups(cbarOper.Groups.Count).ListItems.Clear
    If para组名 = "统计报表" Then
        cmnuSubItem(ii) = "统计报表"
        cmnuSubItem(ii).Top = cmnuSubItem(ii - 1).Top + cmnuSubItem(ii - 1).Height + 150
        cmnuSubItem(ii).Left = cmnuSubItem(0).Left
        cmnuSubItem(ii).Visible = True
        mstrOper(ii) = "统计报表"
        Exit Sub
    End If
    For i = 1 To mobj组.RecordCount
        '修改：2003-7-9（杨春）判断当前操作所属业务名是否在加密狗许可范围内。
        mobj操作.Filter = ""
        mobj操作.Filter = "操作名称" & "='" & mobj组.Fields("操作名称") & "'"
        If mobj操作.RecordCount > 0 Then
            If pstr子系统许可 = "" Or InStr(pstr子系统许可, mobj操作.Fields("业务名") & ",") > 0 Then
                cmnuSubItem(ii) = mobj组("操作别名")
                cmnuSubItem(ii).Top = cmnuSubItem(ii - 1).Top + cmnuSubItem(ii - 1).Height + 150
                cmnuSubItem(ii).Left = cmnuSubItem(0).Left
                cmnuSubItem(ii).Visible = True
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
            pcolWndProc.add llngWndProc, CStr(llng窗体句柄)
            pcol操作名称.add "字典管理", CStr(llng窗体句柄)
            pcol子窗体句柄.add llng窗体句柄, "字典管理"
            Call MoveWindow(llng窗体句柄, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
            
            '********************临时测试代码********************
            'sfsubSaveSetting "系统管理", "平台问题测试", "字典管理初始化", "hWnd：" & llng窗体句柄 & " WndProc：" & llngWndProc & " 时间" & Format(Now, "yyyy年mm月dd日hh时mm分ss秒")
            '****************************************************
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
    
    lstrTime = "2005-12-31"
    
    If lstrTime < Format(Now, "yyyy-mm-dd") Then
        MsgBox "对不起，你的试用期限已到，请与软件供应商联系。", vbCritical, "系统提示"
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

        Call MoveWindow(llngHwnd, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
        'Call MoveWindow(llngHwnd, ScaleX(700, vbTwips, vbPixels), ScaleX(350, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 400, vbTwips, vbPixels), 1)

    Next
End Sub

