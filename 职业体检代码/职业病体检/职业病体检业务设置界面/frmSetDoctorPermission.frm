VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetDoctorPermission 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "医师权限管理"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退出(&X)"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.TreeView ctvPerm 
      Height          =   6015
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10610
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView ctvDept 
      Height          =   6015
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10610
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdDoctor 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
      _cx             =   3413
      _cy             =   10610
      Appearance      =   0
      BorderStyle     =   1
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
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "医师编号|医师姓名"
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label3 
      Caption         =   "操作权限："
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "可操作科室："
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "医师列表："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSetDoctorPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-03-15 于登淼 添加整个窗体和所有权限控制功能
'包括医师列表，选中的某一个医师的所属科室(可多选)，和所选科室下属的可用操作。
'该窗体的权限，由系统设置→用户及权限设置来管理。
'这部分的医师，只有在系统设置→员工管理中被设为06职业病体检科的医师才能放进来。
'医师科室或可用操作打钩或取消时，数据库内立刻完成增加或删除操作，并即时显示。
'对应除去字典操作的其它业务操作，均在新建的"clsPermissionConfigure"中

'-------加载和修改科室和可用操作代码繁琐，以下为简单说明-------
'1、form_load时，加载所有医师和所有科室，不加载可用操作
'2、每次选择医师时，医师已有科室打钩，加载该医师所属科室的所有操作，已选操作打钩，所有操作节点展开(方便查看)
'3、2中流程：LoadDocAllDept → LoadDocAllOperate → LoadOneDeptNowOperate(LoadOneDeptAllOperate) → ExpandAllNodes
'4、科室增加时，科室内所有可用操作选中，并展开；科室取消时，科室内所有可用操作删除
'5、可用操作添加时，包括当前节点在内，所有子节点和对应所有父节点均被添加
'6、可用操作删除时，包括当前节点在内，所有子节点被删除
'7、子节点操作为递归，父节点操作为线性操作
'----------------------------说明完毕--------------------------

Option Explicit
Private mblnInUse As Boolean    '表明当前窗体是否已加载。
Private mobjCls As Object       '定义form内全局变量，专门调用clsPermissionConfigure
Private DoctorNo As String      '定义form内全局变量，记录医师编号（即，用户编号）
Private DoctorDept As Object    '定义form内全局变量，记录医师所属科室
Private DoctorPerm As Object    '定义form内全局变量，记录医师操作权限

Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub ccmdExit_Click()
    Unload Me
    Set mobjCls = Nothing
    Set DoctorDept = Nothing
    Set DoctorPerm = Nothing
    Set frmSetDoctorPermission = Nothing
End Sub

Private Sub cgrdDoctor_Click()
    If cgrdDoctor.MouseRow < 0 Or cgrdDoctor.MouseCol < 0 Then
        Exit Sub
    Else
        ctvDept.Enabled = True
        ctvPerm.Enabled = True
        DoctorNo = cgrdDoctor.TextMatrix(cgrdDoctor.SelectedRow(0), 0)
        LoadDocAllDept DoctorNo
    End If
End Sub

Private Sub ctvDept_NodeCheck(ByVal Node As MSComctlLib.Node)
    If DoctorNo = "" Then Exit Sub
    
    On Error GoTo errHandler
    Dim rootNode As Object
    
    'Node.Checked=false表示取消权限;true表示增加权限
    Call mobjCls.func修改职业病体检医师单个科室(DoctorNo, Node.Checked, Right(Node.Key, Len(Node.Key) - 1))
    If Node.Checked = True Then
        LoadDocAllDept (DoctorNo)
        Set rootNode = func获得科室操作的根节点(Node)
        rootNode.Checked = Node.Checked
        Call mobjCls.func修改医师单个科室系统权限(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        Call mobjCls.func修改职业病体检医师单个可用操作(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        If rootNode.Children > 0 Then Call DownwardModify(rootNode.Child, Node.Checked)
    Else
        Set rootNode = func获得科室操作的根节点(Node)
        rootNode.Checked = Node.Checked
        Call mobjCls.func修改医师单个科室系统权限(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        Call mobjCls.func修改职业病体检医师单个可用操作(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        If rootNode.Children > 0 Then Call DownwardModify(rootNode.Child, Node.Checked)
        LoadDocAllDept (DoctorNo)
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctorPermission", "ctvDept_NodeCheck", 6666, lstrError, False
End Sub

Private Sub ctvPerm_NodeCheck(ByVal Node As MSComctlLib.Node)
    If DoctorNo = "" Then Exit Sub
    
    On Error GoTo errHandler
    
    'Node.Checked=false表示取消权限;true表示增加权限
    Call mobjCls.func修改职业病体检医师单个可用操作(DoctorNo, Node.Checked, Right(Node.Key, Len(Node.Key) - 1))
    If Node.Children > 0 Then Call DownwardModify(Node.Child, Node.Checked)
    If Node.Checked = True Then Call UpwardModify(Node, Node.Checked)
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctorPermission", "ctvPerm_NodeCheck", 6666, lstrError, False
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    cgrdDoctor.SelectionMode = flexSelectionListBox
    cgrdDoctor.AllowSelection = False
    
    Set mobjCls = CreateObject("职业病设置.clsPermissionConfigure")
    DoctorNo = ""
    ctvDept.Enabled = False
    ctvPerm.Enabled = False
    LoadDoctor
    LoadAllDept
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctorPermission", "Form_Load", 6666, lstrError, False
End Sub

Sub LoadDoctor()
    Dim lobjRec As Object
    Dim i As Integer
    On Error GoTo errHandler
    
    Set lobjRec = pobjDict.FetchEx("员工字典")
    
    If lobjRec.recordcount = 0 Then Exit Sub
    lobjRec.movefirst
'    lobjRec.Filter = "科室=06"
    For i = 1 To lobjRec.recordcount
        If lobjRec("编号") <> "0000" And lobjRec("编号") <> "gues" Then
            cgrdDoctor.AddItem lobjRec("编号") & vbTab & lobjRec("姓名"), cgrdDoctor.Rows
        End If
        lobjRec.movenext
    Next
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctorPermission", "LoadDoctor", 6666, lstrError, False
End Sub

Sub LoadAllDept()
    Dim lobjDeptSet As Object
    Set lobjDeptSet = pobjDict.Fetch("职业病体检科室字典")
    If lobjDeptSet.recordcount > 0 Then lobjDeptSet.movefirst
    Do While Not lobjDeptSet.EOF
        ctvDept.Nodes.Add , , "R" & lobjDeptSet("编号"), lobjDeptSet("编号") & "  " & lobjDeptSet("名称")
        lobjDeptSet.movenext
    Loop
End Sub

'加载当前医师的所有科室
Sub LoadDocAllDept(ByVal paraDoctorNo As String)
    On Error GoTo errHandler
    
    Dim lobjRec As Object
    Set lobjRec = pobjDict.Fetch("职业病体检科室字典")

    '如果之前加载过科室，在显示时，须将前面的钩钩去掉(方法有些奇怪)
    ctvDept.Checkboxes = False
    ctvDept.Checkboxes = True
    
    '重新获得医生科室信息，选中科室打钩
    '科室信息加载完后，相应的科室下，医师现有权限打钩
    ctvPerm.Nodes.Clear
    Set DoctorDept = mobjCls.func获取职业病体检单个医师科室(paraDoctorNo)
    If DoctorDept.recordcount > 0 Then DoctorDept.movefirst
    Do While Not DoctorDept.EOF
        ctvDept.Nodes.Item("R" & DoctorDept("科室编号").Value).Checked = True
        DoctorDept.movenext
    Loop

    '加载科室之后，加载所有的现有操作(包括打钩)
    LoadDocAllOperate paraDoctorNo

    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctorPermission", "LoadDocAllDept", 6666, lstrError, False
End Sub

'刷新当前医师的所有操作(包括打钩)
Sub LoadDocAllOperate(ByVal paraDoctorNo As String)
    On Error GoTo errHandler
    
    Dim i As Integer
    Set DoctorPerm = mobjCls.func获取职业病体检单个医师操作权限(paraDoctorNo) '[所有操作权限]
    For i = 1 To ctvDept.Nodes.Count
        If ctvDept.Nodes.Item(i).Checked = True Then
            Dim L As Integer
            Dim lstrDeptName As String
            L = Len(ctvDept.Nodes.Item(i).Text) - 4  'key的首位是加了字母“R”的
            lstrDeptName = Right(ctvDept.Nodes.Item(i).Text, L - 0)
            Call LoadOneDeptNowOperate(paraDoctorNo, lstrDeptName, DoctorPerm)
        End If
    Next

    ExpandAllNodes

    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctorPermission", "LoadDocAllOperate", 6666, lstrError, False
End Sub

'加载当前医师的单个科室的现有操作(包括打钩)
Sub LoadOneDeptNowOperate(ByVal paraDoctorNo As String, ByVal paraDeptName As String, ByVal paraDoctorPerm As Object)
    Call LoadOneDeptAllOperate(paraDeptName)
    
    If paraDoctorPerm.recordcount = 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To ctvPerm.Nodes.Count
        paraDoctorPerm.movefirst
        Do While Not paraDoctorPerm.EOF
            If ctvPerm.Nodes.Item(i).Text = paraDoctorPerm("权限名") Then
                ctvPerm.Nodes.Item(i).Checked = True
            End If
            paraDoctorPerm.movenext
        Loop
    Next
End Sub

'加载单个科室的所有操作（不包括打钩）
Sub LoadOneDeptAllOperate(ByVal paraDeptName As String)
    Dim lobjPerm As Object
    Set lobjPerm = mobjCls.func获取职业病体检单个科室所有操作权限(paraDeptName)
    
    If lobjPerm.recordcount = 0 Then Exit Sub
    lobjPerm.movefirst
    Do While Not lobjPerm.EOF
        If IsNull(lobjPerm("上级操作名")) = True Then
            ctvPerm.Nodes.Add , , "R" & lobjPerm("操作名"), lobjPerm("操作名")
        Else
            ctvPerm.Nodes.Add "R" & lobjPerm("上级操作名"), tvwChild, "R" & lobjPerm("操作名"), lobjPerm("操作名")
        End If
        lobjPerm.movenext
    Loop
End Sub

'从treeview上某个特定节点开始，对上面的子节点进行添加或删除操作(线性)
Sub UpwardModify(ByVal paraNode As Object, ByVal paraCheck As Boolean)
    Dim rootNode As Object
    Set rootNode = func获得科室操作的根节点(paraNode)
    Do While paraNode <> rootNode
        paraNode.Checked = paraCheck
        Call mobjCls.func修改职业病体检医师单个可用操作(DoctorNo, paraNode.Checked, Right(paraNode.Key, Len(paraNode.Key) - 1))
        Set paraNode = paraNode.Parent
    Loop
    paraNode.Checked = paraCheck
    Call mobjCls.func修改职业病体检医师单个可用操作(DoctorNo, paraNode.Checked, Right(paraNode.Key, Len(paraNode.Key) - 1))
End Sub

'从treeview上某个特定节点开始，对下面的子节点进行添加或删除操作(递归)
Sub DownwardModify(ByVal paraNode As Object, ByVal paraCheck As Boolean)
    paraNode.Checked = paraCheck
    Call mobjCls.func修改职业病体检医师单个可用操作(DoctorNo, paraNode.Checked, Right(paraNode.Key, Len(paraNode.Key) - 1))
    
    If paraNode.Children > 0 Then Call DownwardModify(paraNode.Child, paraCheck)
    If paraNode <> paraNode.LastSibling Then Call DownwardModify(paraNode.Next, paraCheck)
End Sub

'展开可用操作的所有节点。
Sub ExpandAllNodes()
    Dim i As Integer
    For i = 1 To ctvPerm.Nodes.Count
        ctvPerm.Nodes(i).Expanded = True
    Next
End Sub

'获得科室对应可用操作的根节点。多处地方用到的，比较重要的函数。
'1、对数据库中可用操作的命名有明确的要求：
'(1)每一科室的可用操作总节点，必须命名为“职业病体检_XXX”格式。如“职业病体检_五官科结果录入”
'(2)总结点下面的子节点，命名时必须以“职业病体检_XXX_”为前缀，后面加上相应操作名称。如“职业病体检_五官科结果录入_保存”
Private Function func获得科室操作的根节点(ByVal paraNodeDept As Object) As Object
    On Error GoTo errHandler
    
    Dim i, idx As Integer
    Dim returnNode As Object
    Dim strArray
    strArray = Split(paraNodeDept.Text, "_", -1, vbBinaryCompare)
    If UBound(strArray) = 0 Then strArray = Split(paraNodeDept.Text, "  ", -1, vbBinaryCompare)
    For i = 1 To ctvPerm.Nodes.Count
        idx = InStr(ctvPerm.Nodes.Item(i).Text, strArray(1))
        If idx <> 0 Then Set returnNode = ctvPerm.Nodes.Item(i): Exit For
    Next
    Set func获得科室操作的根节点 = returnNode
    
    Exit Function
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctorPermission", "func获得科室操作的根节点", 6666, lstrError, False
End Function

Private Sub Form_Resize()
    On Error Resume Next
    cgrdDoctor.Height = Me.ScaleHeight - cgrdDoctor.Top - 20
    ctvDept.Height = cgrdDoctor.Height
    ctvPerm.Height = cgrdDoctor.Height
End Sub
