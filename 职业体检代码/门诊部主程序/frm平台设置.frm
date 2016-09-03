VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm平台设置 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "平台设置"
   ClientHeight    =   7455
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm平台设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleMode       =   0  'User
   ScaleWidth      =   10337.84
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TreeView ctrw权限 
      Height          =   5175
      Left            =   6000
      TabIndex        =   14
      Top             =   1680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   804
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   465
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   10500
      Begin VB.OptionButton copt选项 
         BackColor       =   &H8000000B&
         Caption         =   "主推信息"
         Height          =   270
         Index           =   3
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton copt选项 
         BackColor       =   &H8000000B&
         Caption         =   "操作"
         Height          =   270
         Index           =   0
         Left            =   255
         TabIndex        =   7
         Top             =   120
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton copt选项 
         BackColor       =   &H8000000B&
         Caption         =   "报表"
         Height          =   270
         Index           =   1
         Left            =   4200
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton copt选项 
         BackColor       =   &H8000000B&
         Caption         =   "查询"
         Height          =   270
         Index           =   2
         Left            =   6480
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame fram操作 
      BackColor       =   &H8000000B&
      Height          =   5130
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   1155
      Begin VB.CommandButton ccmd移动 
         BackColor       =   &H8000000B&
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   120
         MaskColor       =   &H00FFF1EC&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   960
      End
      Begin VB.CommandButton ccmd移动 
         BackColor       =   &H8000000B&
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4080
         Width           =   960
      End
      Begin VB.Label clblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "你必须选中左边平台树第三或第四级时，才可以去掉操作！"
         ForeColor       =   &H8000000D&
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label clblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "你必须选中左边平台树第三或第四级时，并选中右边权限时才可以添加操作！"
         ForeColor       =   &H8000000D&
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CheckBox ccheck缺省 
      BackColor       =   &H8000000B&
      Caption         =   "使用缺省别名"
      Height          =   345
      Left            =   4440
      TabIndex        =   0
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5010
      Top             =   3855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm平台设置.frx":030A
            Key             =   "Second"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm平台设置.frx":039E
            Key             =   "First"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   8220
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView ctv平台树 
      Height          =   5235
      Left            =   90
      TabIndex        =   9
      Top             =   1680
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   9234
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar c工具栏 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar c状态栏 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   7080
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label clbl标签 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "平台树(双击操作名可以去掉操作)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   2700
   End
   Begin VB.Label clbl标签 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前用户可用的操作权限(双击操作名可以添加)"
      Height          =   180
      Index           =   1
      Left            =   6120
      TabIndex        =   11
      Top             =   1440
      Width           =   3780
   End
   Begin VB.Menu 平台 
      Caption         =   "平台"
      Visible         =   0   'False
      Begin VB.Menu add 
         Caption         =   "添加(&A)"
         Index           =   1
      End
      Begin VB.Menu delete 
         Caption         =   "删除(&D)"
         Index           =   2
      End
      Begin VB.Menu modify 
         Caption         =   "修改(&U)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frm平台设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj类 As Collection               '当前平台的操作组分类
Private mobjSmartInfos As Collection        '当前用户的主推信息
Private mobj操作 As Collection              '当前平台的所有操作
Private mobj可用操作 As Object
Private mobj查询 As Collection              '当前平台的所有查询
Private mobj报表 As Collection               '当前平台的所有报表
Private mobj权限 As Object                  '当前用户可用操作权限
Private mint当前选项 As Integer            ' 当前用户的选择的操作
Private WithEvents mobjGUI As cls界面通用对象 '界面上引用的界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mblnSave  As Boolean                 '判断是否已经保存

Private Enum 移动方式
    左移 = 0
    右移 = 1
End Enum

Private Enum 当前选项
    操作 = 0
    报表 = 1
    查询 = 2
    主推信息 = 3
End Enum

Public pblnInUse As Boolean



'功能：响应用户选择的平台操作
'输入：移动方式
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
Private Sub ccmd移动_Click(Index As Integer)
    On Error GoTo errHandle
    Dim lnodeTemp As Node
    Dim lobj操作 As Collection
    Dim lstrChild As String
    Dim lstr别名 As String
    Dim lintCount As Integer
    Dim lstr操作名 As String
    Dim i As Integer
    Dim ii As Integer
    Dim llngChildren As Long '子节点数。
    Dim llngIndex As Long
    
    mblnSave = False
    
    Select Case Index
        Case 左移                         '左移
            If ctv平台树.SelectedItem Is Nothing Then Exit Sub
            
            '判断能否左移。
            If func节点位置(ctv平台树.SelectedItem.FullPath) = 3 Or func节点位置(ctv平台树.SelectedItem.FullPath) = 4 Then
                If ctrw权限.SelectedItem.Parent Is Nothing Then
                    '一类操作的添加。
                    llngIndex = ctrw权限.SelectedItem.Child.Index
                    llngChildren = ctrw权限.SelectedItem.Children
                    For i = llngIndex To llngIndex + llngChildren - 1
                        lstr操作名 = ctrw权限.Nodes(i).Key
                        If Not func拥有(lstr操作名) Then
                            lstr别名 = lstr操作名                         '使用缺省别名
                            '修改：2001-11-2（去掉业务名前缀）。
                            If InStr(lstr别名, "_") > 0 Then
                                lstr别名 = Right(lstr别名, Len(lstr别名) - InStr(lstr别名, "_"))
                            End If
                            Call sub增加(lstr操作名, lstr别名)                           '调用移动过程
                        End If
                    Next
                Else
                    '单个操作添加。
                    lstr操作名 = ctrw权限.SelectedItem.Key
                    '判断该操作用户是否已拥有。
                    If func拥有(lstr操作名) Then
                        Call sffuncMsg("你已拥有" & lstr操作名 & "操作！", sf警告)
                    Else
                        If ccheck缺省 = 0 Then                            '不使用缺省别名
                            lstr别名 = InputBox("请输入该操作的别名", "系统提示", lstr操作名)
                            lstr别名 = Trim(Replace(lstr别名, "'", ""))
                            If lstr别名 = "" Then Exit Sub
                            If Len(lstr别名) > 15 Then Call sffuncMsg("操作别名不能超过十五个字符!", sf警告): Exit Sub
                        Else
                            lstr别名 = lstr操作名                         '使用缺省别名
                            '修改：2001-11-2（去掉业务名前缀）。
                            If InStr(lstr别名, "_") > 0 Then
                                lstr别名 = Right(lstr别名, Len(lstr别名) - InStr(lstr别名, "_"))
                            End If
                        End If
                        Call sub增加(lstr操作名, lstr别名)                           '调用移动过程
                    End If
                End If
            End If
        Case 右移
            If func节点位置(ctv平台树.SelectedItem.FullPath) = 3 Then
                '全右移。
                lintCount = ctv平台树.SelectedItem.Children               '该组下所有的操作数
                For ii = 1 To lintCount
                   lstrChild = ctv平台树.SelectedItem.Child.Key
                    If func节点位置(ctv平台树.SelectedItem.Child.FullPath) = 4 Then
                        Call sub删除(lstrChild, "全")
                    End If
                Next ii
            ElseIf func节点位置(ctv平台树.SelectedItem.FullPath) = 4 Then    '判断能否右移
                '单个右移。
                lstr操作名 = ctv平台树.SelectedItem.Key   '当前选中的平台树中的操作名称
                Call sub删除(lstr操作名, "")
            End If
    End Select
    ctv平台树.Refresh
    
    Exit Sub
errHandle:
    Call sfsub错误处理("平台设置", "frm平台设置", "ccmd移动_Click", Err.Number, Err.Description, False)
End Sub



'功能：根据用户的选择增加用户的操作类或操作组
'输入：无
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-15
Private Sub subAdd()
    On Error GoTo errHandle
    Dim lstr名称 As String      '地增加的类或组的名称
    Dim lstrTemp As String      '确定增加的类或组
    Dim lobj类 As New Collection '增加的集合
    Dim lnodeTemp As Node
    Select Case func节点位置(ctv平台树.SelectedItem.FullPath)  '根据节点的位置确定是增加类还是增加组
        Case 1
            lstrTemp = "类的名称"
        Case 2
            lstrTemp = "组的名称"
        Case Else
            sffuncMsg "请先选定需要增加的类型", sf警告
            Exit Sub
    End Select
    lstr名称 = InputBox("请输入你要增加" & lstrTemp, "系统提示")
    lstr名称 = Trim(Replace(lstr名称, "'", ""))
    If lstr名称 = "" Then Exit Sub '用户取消
    If Len(lstr名称) > 6 Then Call sffuncMsg("类或组的名称不能超过六个字符!", sf警告): Exit Sub
    If IsNumeric(lstr名称) Then Call sffuncMsg("名称不能全部是数字!", sf警告): Exit Sub
    If IsDate(lstr名称) Then Call sffuncMsg("名称不能是日期形式!", sf警告): Exit Sub
    If funcInOperation(lstr名称) Then sffuncMsg "新增的名称不能是系统已有的操作名称! ", sf警告: Exit Sub
    Set lnodeTemp = ctv平台树.Nodes.Add(ctv平台树.SelectedItem.Key, 4, lstr名称, lstr名称, "First")   'ctv平台树增加节点
    If lstrTemp = "组的名称" Then          '增加操作组
        lobj类.Add ctv平台树.SelectedItem.Text, "所属类名"
        lobj类.Add lstr名称, "操作组"
    Else                                 '增加操作类
        lobj类.Add lstr名称, "所属类名"
        lobj类.Add CStr(Now), "操作组"
    End If
    mobj类.Add lobj类
    mblnSave = False
    Exit Sub
errHandle:
    If Err.Number = 35602 Then
        Err.Number = 6666
        Err.Description = "你新增的名称已经存在，请换个名称！"
    End If
    Call sfsub错误处理("主程序", "frm平台设置", "form_Load", Err.Number, Err.Description, False)
End Sub



'功能：根据用户的选择删除用户的操作类或操作组
'输入：无
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-15
Private Sub subDelete()
    On Error GoTo errHandle
    Dim lstrTemp As String
    Dim i As Integer
    Select Case func节点位置(ctv平台树.SelectedItem.FullPath)  '判断该节点能否删除
        Case 1
            Call sffuncMsg("根结点不允许删除！", sf警告)         '根结点不允许删除
        Case 2                      '删除 操作类
            If ctv平台树.SelectedItem.Children > 0 Then          '该结点下有子结点不允许删除
                Call sffuncMsg("请先删除该结点的子结点！", sf警告)
            Else
                For i = 1 To mobj类.Count
                    If mobj类(i)("所属类名") = ctv平台树.SelectedItem.Text Then
                        mobj类.Remove (i)
                        Exit For
                    End If
                Next i
                ctv平台树.Nodes.Remove (ctv平台树.SelectedItem.Key)
            End If
               mblnSave = False
        Case 3        '删除 操作组
            lstrTemp = func能否删除(ctv平台树.SelectedItem.Key)
            If Not lstrTemp = "" Then
                Call sffuncMsg("请先删除该结点的" & lstrTemp & "中的子结点！", sf警告)
            Else
                For i = 1 To mobj类.Count
                    If mobj类(i)("操作组") = ctv平台树.SelectedItem.Text Then
                        mobj类.Remove (i)
                        ctv平台树.Nodes.Remove (ctv平台树.SelectedItem.Key)
                        Exit For
                    End If
                Next i
            End If
             mblnSave = False
        Case 4
            mblnSave = False
            Call sub删除(ctv平台树.SelectedItem.Key, "")         '删除权限
    End Select
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "subDelete", Err.Number, Err.Description, False)
End Sub

Private Sub copt选项_Click(Index As Integer)
    On Error Resume Next
    mint当前选项 = Index
    subLoad平台树 (Mid(copt选项.Item(Index).Caption, 1, 2))
    subLoad权限树 (Mid(copt选项.Item(Index).Caption, 1, 2))
    Set ctv平台树.SelectedItem = ctv平台树.Nodes(1)
End Sub

Private Sub ctrw权限_Click()
    ctv平台树_Click
End Sub

'响应用户的双击操作
Private Sub ctrw权限_DblClick()
    On Error Resume Next
    If ccmd移动(左移).Enabled = True And Not ctrw权限.SelectedItem Is Nothing Then
        '只有单项可以通过双击左移。
        If Not ctrw权限.SelectedItem.Parent Is Nothing Then
            Call ccmd移动_Click(左移)
        End If
    End If
End Sub

'确定移动的按纽何时可用
Private Sub ctv平台树_Click()
    On Error Resume Next
    ccmd移动(右移).Enabled = False
    ccmd移动(左移).Enabled = False
    Select Case func节点位置(ctv平台树.SelectedItem.FullPath)
        Case 3, 4
            ccmd移动(右移).Enabled = True
            If Not ctrw权限.SelectedItem Is Nothing Then
                ccmd移动(左移).Enabled = True
            End If
    End Select
'    clblInfo(0).Visible = Not ccmd移动(左移).Enabled
'    clblInfo(1).Visible = Not ccmd移动(右移).Enabled
End Sub

Private Sub ctv平台树_DblClick()
    '双击可以去掉权限。
    On Error Resume Next
    If ccmd移动(右移).Enabled = True And Not ctv平台树.SelectedItem Is Nothing Then
        '只有单项可以通过双击左移。
        If func节点位置(ctv平台树.SelectedItem.FullPath) = 4 Then
            Call ccmd移动_Click(右移)
        End If
    End If
    
End Sub

Private Sub ctv平台树_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandle
    If Button = 2 Then
        PopupMenu 平台, vbPopupMenuCenterAlign
    End If
    Exit Sub
errHandle:
    Call sfsub错误处理("平台设置", "frm平台设置", "ctv平台树_MouseUp", Err.Number, Err.Description, False)
End Sub


Private Sub Form_Load()
    On Error GoTo errHandle
    'dasubInitialize ("driver=sql server;server=wangxiaohua;database=防疫26系统管理数据库;uid=user26;pwd=welcome")
    If pblnInUse Then Exit Sub
    pblnInUse = True
    Dim llng As Long
    llng = GetWindowLong(Me.hWnd, GWL_STYLE)
    
    '隐藏标题栏。
'    If (llng And WS_BORDER) = WS_BORDER Then
'        llng = llng - WS_BORDER
'    End If
    SetWindowLong Me.hWnd, GWL_STYLE, llng
    SetWindowPos Me.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
    
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.c工具栏 = c工具栏
    Set mobjGUI.c状态栏 = c状态栏
    
    lcol工具栏按钮.Add "添加"
    lcol工具栏按钮.Add "删除"
    lcol工具栏按钮.Add "修改"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "保存"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    
    mobjGUI.subInitialize lcol工具栏按钮, ""
    pobj平台结构.平台名称 = um用户编号
    
    Set mobj类 = funcConvertColl(pobj平台结构.操作组分类)
    Set mobj操作 = funcConvertColl(pobj平台结构.Operations)
    Set mobj查询 = funcConvertColl(pobj平台结构.Queries)
    Set mobj报表 = funcConvertColl(pobj平台结构.Reports)
    Set mobj可用操作 = pobj平台结构.Operation
    
    If um用户编号 = "0000" And um操作权限.RecordCount = 0 Then
        On Error Resume Next
        '若系统管理员的权限为空，自动添加权限。
        dafuncGetData "insert into 系统管理_用户操作权限表 values('0000','系统管理_用户及权限设置')"
        On Error GoTo errHandle
    End If
    
    If um操作权限.RecordCount = 0 Then sffuncMsg "请通知系统管理员先进入“用户权限设置”操作，给你增加操作权限！", sf警告
    
    Set mobj权限 = funcConvertColl(um操作权限)
    Set mobjSmartInfos = funcConvertColl(pobj平台结构.SmartInfos)
    subLoad平台树 ("")       '初始化界面上的平台信息
    subLoad权限树 ("操作")   '初始化界面上的可用操作信息
    mint当前选项 = 0
    Call copt选项_Click(mint当前选项)
    
    '修改：2001-8-24（杨春：若是系统管理员，不能设置主推信息）。
    If um用户编号 = "0000" Then
        copt选项(3).Visible = False
    Else
        copt选项(3).Visible = True
    End If
    
    mblnSave = True
    c状态栏.Panels(1).Text = "请注意在增加类名，组名时不要与已有操作同名。推荐类名组名以类字和组字结尾。"
    Exit Sub
errHandle:
    Call sfsub错误处理("平台设置", "frm平台设置", "form_Load", Err.Number, Err.Description, False)
End Sub


'功能：根据用户选择的类型，显示不同的用户平台信息
'输入：操作类型
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
Private Sub subLoad平台树(ByVal para操作 As String)
    On Error GoTo errHandle
    Dim lobj组 As Object
    Dim lnodeTemp As Node
    Dim lstrTemp As String
    Dim i As Integer
    
    ctv平台树.Nodes.Clear        '清空原有的平台信息
    ctv平台树.Nodes.Add , , "平台结构", "当前用户的平台结构", "First" '加入初始值
    If mobj类.Count > 0 Then      '加入操作分类
        For i = 1 To mobj类.Count
            Set lnodeTemp = ctv平台树.Nodes.Add("平台结构", 4, mobj类(i)("所属类名"), mobj类(i)("所属类名"), "First")
        Next i
        '加入操作分组
        For i = 1 To mobj类.Count
            lstrTemp = mobj类(i)("所属类名")
            If Not IsDate(mobj类(i)("操作组")) Then
                Set lnodeTemp = ctv平台树.Nodes.Add(lstrTemp, 4, mobj类(i)("操作组"), mobj类(i)("操作组"), "First")
            End If
        Next i
    
        Select Case para操作   '根据不同的需求，显示不同的平台信息
            Case "报表"
                Set lobj组 = mobj报表       '报表信息
            Case "查询"
                Set lobj组 = mobj查询      '查询信息
            Case "主推"
                Set lobj组 = mobjSmartInfos '主推信息
            Case Else
                Set lobj组 = mobj操作       '操作信息
        End Select
        If lobj组.Count > 0 Then    '将信息填入TreeView
            For i = 1 To lobj组.Count
                lstrTemp = lobj组(i)("所属组名")
                '修改：2003-7-9（杨春）判断当前操作所属业务名是否在加密狗许可范围内。
                mobj可用操作.Filter = ""
                mobj可用操作.Filter = "操作名称" & "='" & lobj组(i)("操作名称") & "'"
                If mobj可用操作.RecordCount > 0 Then
                    If pstr子系统许可 = "" Or InStr(pstr子系统许可, mobj可用操作.Fields("业务名") & ",") > 0 Then
                        Set lnodeTemp = ctv平台树.Nodes.Add(lstrTemp, 4, lobj组(i)("操作名称"), lobj组(i)("操作名称") & "(别名：" & lobj组(i)("操作别名") & ")", "Second")
                    End If
                End If
            Next i
        End If
    End If
    Exit Sub
errHandle:
    If Err.Number = 35602 Then
        Resume Next
    Else
        Call sfsub错误处理("主程序", "frm平台设置", "subload平台树", Err.Number, Err.Description, False)
    End If
End Sub


'功能：根据用户选择的类型，显示不同的用户可见的信息
'输入：操作类型
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
'修改：2001-11-2（杨春）把权限树改为树状。
Private Sub subLoad权限树(ByVal para操作 As String)
    On Error GoTo errHandle
    Dim lobj权限 As Object
    Dim lobj一级操作 As Object
    Dim i As Integer
    Dim lobjList As ListItem
    Dim lstr业务名 As String
    
    ctrw权限.Nodes.Clear
    Select Case para操作
        Case "报表"    '可见报表
            
        Case "查询"
        Case "主推"   '可见的主推信息
            Set lobj权限 = dafuncGetData("select 操作名,业务名 from 系统管理_业务主推信息表")
            If lobj权限.RecordCount > 0 Then
                lobj权限.MoveFirst
                For i = 1 To lobj权限.RecordCount
                    On Error Resume Next
                    '加入业务名根节点。
                    ctrw权限.Nodes.Add , , lobj权限("业务名"), lobj权限("业务名")
                    '加入操作子节点。
                    ctrw权限.Nodes.Add lobj权限("业务名").Value, tvwChild, lobj权限("操作名"), lobj权限("操作名")
                    On Error GoTo errHandle
                    lobj权限.MoveNext
                Next i
            End If
        Case Else   '可用的操作权限
            Set lobj一级操作 = pobj平台结构.一级操作
            If mobj权限.Count > 0 And lobj一级操作.RecordCount > 0 Then
                For i = 1 To mobj权限.Count
                    '判断是否是一级操作权限。
                    lobj一级操作.Filter = "操作名" & "= '" & mobj权限(i)("权限名") & "'"
                    If lobj一级操作.RecordCount > 0 Then
                        '修改：2003-7-9（杨春）判断当前操作所属业务名是否在加密狗许可范围内。
                        If pstr子系统许可 = "" Or InStr(pstr子系统许可, lobj一级操作("业务名") & ",") > 0 Then
                            On Error Resume Next
                            '加入业务名根节点。
                            ctrw权限.Nodes.Add , , lobj一级操作("业务名"), lobj一级操作("业务名")
                            '加入操作子节点。
                            ctrw权限.Nodes.Add lobj一级操作("业务名").Value, tvwChild, mobj权限(i)("权限名"), mobj权限(i)("权限名")
                            On Error GoTo errHandle
                        End If
                   End If
                   lobj一级操作.Filter = ""
                Next i
           Else
           End If
    End Select
Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "subload权限树", Err.Number, Err.Description, False)
End Sub


'功能：根据用户选择的操作，判断是否拥有
'输入：操作
'输出：无
'返回：True ：拥有，False：没有
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
Private Function func拥有(ByVal para操作 As String) As Boolean
    On Error GoTo errHandle
    Dim i As Integer
    func拥有 = False
    For i = 1 To ctv平台树.Nodes.Count    '遍历平台树
       ' MsgBox ctv平台树.Nodes.Item(i).Key
    If para操作 = ctv平台树.Nodes.Item(i).Key Then
            func拥有 = True              '拥有
            Exit For
        End If
    Next i
    Exit Function
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "func拥有", Err.Number, Err.Description, False)
End Function


'功能：根据用户选择的操作，判断该操作属于平台树的第几层
'输入：名称的完整路径
'输出：无
'返回：１：顶层　２：类层　３：组层 ４：操作层
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
Private Function func节点位置(ByVal paraPath As String) As Integer
    On Error GoTo errHandle
    Dim lint As Integer
    func节点位置 = 1
    lint = InStr(1, paraPath, "\", vbTextCompare) '判断路径中"\"的数量
    Do While lint > 0
        lint = InStr(lint + 1, paraPath, "\", vbTextCompare)
        func节点位置 = func节点位置 + 1
    Loop
Exit Function
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "func节点位置", Err.Number, Err.Description, False)
End Function


'功能：根据用户选择的操作，将该操作移到平台树里
'输入：操作别名
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
'修改：2001-11-2（杨春）添加参数“para操作名”。
Private Sub sub增加(ByVal para操作名 As String, ByVal para别名 As String)
    On Error GoTo errHandle
    Dim i As Integer
    Dim lobj操作 As New Collection
    Dim lstrTemp As String
    Dim lnodeTemp As Node
    
    If func节点位置(ctv平台树.SelectedItem.FullPath) = 3 Then
        Set lnodeTemp = ctv平台树.SelectedItem
    Else
        Set lnodeTemp = ctv平台树.SelectedItem.Parent
    End If
    
    lstrTemp = lnodeTemp.Key
    lobj操作.Add lstrTemp, "所属组名"
    lobj操作.Add para操作名, "操作名称"
    lobj操作.Add para别名, "操作别名"
    Select Case mint当前选项
        Case 操作
            mobj操作.Add lobj操作
        Case 查询
            mobj查询.Add lobj操作
        Case 报表
            mobj报表.Add lobj操作
        Case 主推信息
            mobjSmartInfos.Add lobj操作
    End Select
    Set lnodeTemp = ctv平台树.Nodes.Add(lstrTemp, 4, para操作名, para操作名 & "(别名：" & para别名 & ")", "Second")
    
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "sub增加", Err.Number, Err.Description, False)
End Sub

'功能：根据用户选择的操作，将该操作从平台树里删除
'输入：操作名称
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
'修改：2001-11-2（杨春）去掉更改权限树中的拥有字段。
Private Sub sub删除(ByVal paraName As String, ByVal para移除方式 As String)
    On Error GoTo errHandle
    Dim i As Integer
    Dim j As Integer
    mblnSave = False
    If para移除方式 = "全" Then
        ctv平台树.Nodes.Remove (ctv平台树.SelectedItem.Child.Index)   '移除
    Else
        ctv平台树.Nodes.Remove (ctv平台树.SelectedItem.Index)   '移除
    End If
    Select Case mint当前选项
    Case 操作
        For j = 1 To mobj操作.Count
            If mobj操作(j)("操作名称") = paraName Then
                mobj操作.Remove (j)
                Exit For
            End If
        Next j
    Case 查询
        For j = 1 To mobj查询.Count
            If mobj查询(j)("操作名称") = paraName Then
                mobj查询.Remove (j)
                Exit For
            End If
        Next j
    Case 报表
        For j = 1 To mobj报表.Count
            If mobj报表(j)("操作名称") = paraName Then
                mobj报表.Remove (j)
                Exit For
            End If
        Next j
    Case 主推信息
        For j = 1 To mobjSmartInfos.Count
            If mobjSmartInfos(j)("操作名称") = paraName Then
                mobjSmartInfos.Remove (j)
                Exit For
            End If
        Next j
End Select
    ctv平台树.Refresh                                      '刷新平台树
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "sub删除", Err.Number, Err.Description, False)
End Sub

'功能：将记录集转换成集合
'输入：记录集
'输出：无
'返回：集合
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
Private Function funcConvertColl(ByVal paraRes As Object) As Object
    On Error GoTo errHandle
    Dim i As Integer  '循环变量
    Dim lintTemp As Integer
    Dim lobjTemp As Collection  '集合
    Dim lstr项目 As String      '项目名称
    Dim lstr字段名 As String    '字段名称
    Set funcConvertColl = New Collection
    If paraRes.RecordCount > 0 Then      '记录集的记录数必须大于零
        paraRes.MoveFirst
        For i = 1 To paraRes.RecordCount
            Set lobjTemp = New Collection
            lintTemp = paraRes.Fields.Count
            Do While lintTemp > 0
                lstr项目 = paraRes.Fields(lintTemp - 1)
                lstr字段名 = paraRes.Fields(lintTemp - 1).Name
                lobjTemp.Add lstr项目, lstr字段名
                'lobjTemp.Add paraRes.Fields("操作组"), "操作组"
                lintTemp = lintTemp - 1
            Loop
            paraRes.MoveNext
            funcConvertColl.Add lobjTemp, CStr(i)
        Next i
    End If
    Exit Function
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "funcConvertColl", Err.Number, Err.Description, False)
End Function


'功能：保存用户对平台的修改
'输入：无
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
Private Sub subSave()
    On Error GoTo errHandle
    '将平台的类组保存
    pobj平台结构.操作组分类 = mobj类   '同时删除平台以前的操作组设置。
    pobj平台结构.Operations = mobj操作 '给操作组赋值
    pobj平台结构.Queries = mobj查询    '给查询赋值
    pobj平台结构.Reports = mobj报表    '给报表赋值
    pobj平台结构.SmartInfos = mobjSmartInfos '给主推信息赋值
    pobj平台结构.funcSaveSetupOP       '保存操作设置
    pobj平台结构.funcSaveSetupRP       '保存报表设置
    pobj平台结构.funcSaveSetupQE       '保存查询设置
    pobj平台结构.funcSaveSetupSI       '保存主推信息设置
    sffuncMsg "要使当前设置有效，请注销重新登录系统！", sf警告
    mblnSave = True
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "subSave", Err.Number, Err.Description, False)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Set mobjGUI.Form = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj权限 = Nothing
    Set mobjGUI = Nothing
    pblnInUse = False
End Sub

'工具栏操作
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandle
    Select Case Operate
        Case "保存"
            subSave

            Cancel = True
        Case "删除"
            subDelete
            Cancel = True
        Case "添加"
            subAdd
            Cancel = True
        Case "修改"
            subModify
            Cancel = True
        Case "退出"
            If mblnSave = False Then
                If sffuncMsg("当前所做的改动是否保存？", sf询问) Then
                    subSave
                End If
            End If
    End Select
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False)
End Sub


'功能：判断该组能否删除
'输入：组名
'输出：无
'返回：如果可以删除则返回空串，否则返回相应的提示信息
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
Private Function func能否删除(ByVal para组名 As String) As String
    On Error GoTo errHandle
    Dim i As Integer
    For i = 1 To mobj报表.Count        '判断报表下是否有该组
        If mobj操作(i)("所属组名") = para组名 Then
            func能否删除 = "报表信息"
            Exit Function
        End If
    Next i
    For i = 1 To mobj查询.Count       '判断查询下是否有该组
        If mobj查询(i)("所属组名") = para组名 Then
            func能否删除 = "查询信息"
            Exit Function
        End If
    Next i
    For i = 1 To mobjSmartInfos.Count     '判断主推信息下是否有该组
        If mobjSmartInfos(i)("所属组名") = para组名 Then
            func能否删除 = "主推信息"
            Exit Function
        End If
    Next i
    For i = 1 To mobj操作.Count         ' '判断操作下是否有该组
        If mobj操作(i)("所属组名") = para组名 Then
            func能否删除 = "操作信息"
            Exit Function
        End If
    Next i
    func能否删除 = ""
    Exit Function
errHandle:
Call sfsub错误处理("主程序", "frm平台设置", "func能否删除", Err.Number, Err.Description, True)
End Function



'功能：判断新类名或组名是否和已有操作同名
'输入：类名或组名
'输出：无
'返回：True：存在同名 False：不存在同名
'注意事项：
'作者：王晓华
'创建时间：2001-3-12
'修改：2001-11-2（杨春）。
Private Function funcInOperation(ByVal para名称 As String) As Boolean
    Dim lobjNode As Node
    
    On Error GoTo errHandle
    
    funcInOperation = True
    For Each lobjNode In ctrw权限.Nodes
        If Not lobjNode.Parent Is Nothing Then
            If para名称 = lobjNode.Key Then
                funcInOperation = True
                Exit Function
            End If
        End If
    Next
    funcInOperation = False
    Exit Function
errHandle:
Call sfsub错误处理("主程序", "frm平台设置", "funcinoperation", Err.Number, Err.Description, True)
End Function



'功能：修改类名、组名或操作别名
'输入：无
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-4-16
Private Sub subModify()
    On Error GoTo errHandle
    Dim lstr旧名称 As String      '修改的类或组的名称
    Dim lstr名称 As String      '修改的类或组的名称
    Dim lstrTemp As String      '确定修改的类或组
    Dim lobj类 As New Collection '修改的集合
    Dim i As Integer
    Dim lnodeTemp As Node
    Select Case func节点位置(ctv平台树.SelectedItem.FullPath)  '根据节点的位置确定是增加类还是增加组
        Case 2
            lstrTemp = "类名"
            lstr旧名称 = ctv平台树.SelectedItem.Text
            
        Case 3
            lstrTemp = "组名"
            lstr旧名称 = ctv平台树.SelectedItem.Text
            
        Case 4
            lstrTemp = "别名"
            lstr旧名称 = ctv平台树.SelectedItem.Text
            lstr旧名称 = Mid(lstr旧名称, Len(ctv平台树.SelectedItem.Key) + 5)
            lstr旧名称 = Left(lstr旧名称, Len(lstr旧名称) - 1)     '将操作别名取出
        Case Else
            sffuncMsg "请先选定需要修改的类、组或操作别名", sf警告
            Exit Sub
    End Select
    lstr名称 = InputBox("请输入" & lstr旧名称 & "新的名称", "系统提示", lstr旧名称)
    lstr名称 = Trim(Replace(lstr名称, "'", ""))
    If IsDate(lstr名称) Then Call sffuncMsg("名称不能是日期形式!", sf警告): Exit Sub
    If lstr名称 = "" Or lstr旧名称 = lstr名称 Then Exit Sub '用户取消
    If IsNumeric(lstr名称) Then Call sffuncMsg("名称不能全部是数字!", sf警告): Exit Sub
    If lstrTemp <> "别名" Then
        If Len(lstr名称) > 6 Then Call sffuncMsg("类或组的名称不能超过六个字符!", sf警告): Exit Sub
        If funcInOperation(lstr名称) Then sffuncMsg "新增的名称不能是系统已有的操作名称!", sf警告: Exit Sub
        For i = 1 To mobj类.Count
            If lstr名称 = mobj类.Item(i)("所属类名") Then  '名称存在
                Err.Raise 6666, , "你改的新名称已经存在，请换个名称！"
                Exit For
            End If
        Next i
        For i = 1 To mobj类.Count
            If lstr名称 = mobj类.Item(i)("操作组") Then
                Err.Raise 6666, , "你改的新名称已经存在，请换个名称！"
                Exit For
            End If
        Next i
    Else
        If Len(lstr名称) > 15 Then Call sffuncMsg("操作别名不能超过十五个字符!", sf警告): Exit Sub
    End If
    If func修改名称(lstr旧名称, lstr名称, lstrTemp) Then
        If lstrTemp <> "别名" Then
            ctv平台树.SelectedItem.Text = lstr名称  '界面换名
            ctv平台树.SelectedItem.Key = lstr名称
        Else
            ctv平台树.SelectedItem.Text = ctv平台树.SelectedItem.Key & "(别名：" & lstr名称 & ")"
        End If
    End If
    mblnSave = False
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "subModify", Err.Number, Err.Description, False)
End Sub

'功能：修改类名、部分组名
'输入：旧名称，新名称，修改类型
'输出：无
'返回：True：修改成功　False：修改失败
'注意事项：
'作者：王晓华
'创建时间：2001-4-16
Private Function func修改名称(ByVal para旧名称 As String, ByVal para新名称 As String, ByVal para类型 As String) As Boolean
    On Error GoTo errHandle
    Dim i As Integer
    Dim lstr部件名 As String
    Dim lstr类名 As String
    Dim lstr操作组 As String
    Dim lstr别名 As String
    Dim lstr类 As String
    Dim lstr操作 As String
    Dim lstrTemp As String
    Dim lobjTemp As Collection
    Dim lobj组 As Collection
    Select Case para类型    '判断更改的是类名、组名、别名
        Case "类名"     '类名
           For i = 1 To mobj类.Count
                If mobj类.Item(i)("所属类名") = para旧名称 Then
                    lstr操作组 = mobj类.Item(i)("操作组")
                    Set lobjTemp = New Collection
                    mobj类.Remove (i)         '删除原有的
                    lobjTemp.Add para新名称, "所属类名"
                    lobjTemp.Add lstr操作组, "操作组"
                    mobj类.Add lobjTemp
                    i = 0
                End If
            Next i
            func修改名称 = True
           Exit Function
        Case "组名"     '组名先在类中删除该组
            lstrTemp = ctv平台树.SelectedItem.Key
            For i = 1 To mobj类.Count
                If mobj类.Item(i)("操作组") = para旧名称 Then
                    lstr类 = mobj类.Item(i)("所属类名")
                    Set lobjTemp = New Collection
                    mobj类.Remove (i)
                    lobjTemp.Add lstr类, "所属类名"
                    lobjTemp.Add para新名称, "操作组"
                    mobj类.Add lobjTemp
                    Exit For
                End If
            Next i
'            Set lobj组 = mobj操作
        Case "别名"
            lstrTemp = ctv平台树.SelectedItem.Key
    End Select
    For i = 1 To 4
        func针对具体的类型修改 para旧名称, para新名称, para类型, i
    Next i
    func修改名称 = True
    Exit Function
errHandle:
    func修改名称 = False
    Call sfsub错误处理("主程序", "frm平台设置", "func修改名称", Err.Number, Err.Description, True)
End Function

'功能：根据传过来的数依次修改操作中的组，报表中的组，主推信息中的组
'输入：旧名称，新名称，修改的类型，计数
'输出：无
'返回：无
'注意事项：
'作者：王晓华
'创建时间：2001-4-16
Private Sub func针对具体的类型修改(ByVal para旧名称 As String, ByVal para新名称 As String, ByVal para类型 As String, ByVal paraInt数 As Integer)
    On Error GoTo errHandle
    Dim lobj组 As Object
    Dim lobjTemp As Object
    Dim lstr操作组 As String
    Dim lstr操作 As String
    Dim lstr别名 As String
    Dim lstrTemp As String
    Dim i As Integer
    Select Case paraInt数
        Case 1
        Set lobj组 = mobj操作
        Case 2
        Set lobj组 = mobj报表
        Case 3
        Set lobj组 = mobj查询
        Case 4
        Set lobj组 = mobjSmartInfos
    End Select
    If para类型 = "组名" Then '更改组名
        For i = 1 To lobj组.Count
            If lobj组.Item(i)("所属组名") = para旧名称 Then
                lstr操作 = lobj组.Item(i)("操作名称")
                lstr别名 = lobj组.Item(i)("操作别名")
                Set lobjTemp = New Collection
                lobj组.Remove (i)
                lobjTemp.Add para新名称, "所属组名"
                lobjTemp.Add lstr操作, "操作名称"
                lobjTemp.Add lstr别名, "操作别名"
                lobj组.Add lobjTemp
                i = 0
            End If
        Next i
    Else         '更改别名
        lstrTemp = ctv平台树.SelectedItem.Key
       For i = 1 To lobj组.Count
          If lobj组.Item(i)("操作名称") = lstrTemp Then
             lstr操作组 = lobj组.Item(i)("所属组名")
            Set lobjTemp = New Collection
            lobj组.Remove (i)
            lobjTemp.Add lstr操作组, "所属组名"
            lobjTemp.Add lstrTemp, "操作名称"
            lobjTemp.Add para新名称, "操作别名"
            lobj组.Add lobjTemp
            Exit For
        End If
        Next i
    End If
    Select Case paraInt数
        Case 1
        Set mobj操作 = lobj组
        Case 2
        Set mobj报表 = lobj组
        Case 3
        Set mobj查询 = lobj组
        Case 4
        Set mobjSmartInfos = lobj组
    End Select
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm平台设置", "func修改名称", Err.Number, Err.Description, True)
End Sub

Private Sub add_Click(Index As Integer)
    subAdd
End Sub

Private Sub delete_Click(Index As Integer)
    subDelete
End Sub

Private Sub modify_Click(Index As Integer)
    subModify
End Sub
