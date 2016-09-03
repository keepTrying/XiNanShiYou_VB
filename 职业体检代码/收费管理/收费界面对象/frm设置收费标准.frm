VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm设置收费标准 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "收费标准设置"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10335
   ClipControls    =   0   'False
   Icon            =   "frm设置收费标准.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "请先选择子系统和使用范围"
      ForeColor       =   &H00C00000&
      Height          =   6375
      Left            =   3240
      TabIndex        =   7
      Top             =   960
      Width           =   7095
      Begin VB.TextBox ctxt旧标准名称 
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton ccmdDel 
         Caption         =   "-->"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "<--"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   495
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   3735
         _cx             =   23861692
         _cy             =   23864232
         _ConvInfo       =   1
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
         BackColorBkg    =   16777215
         BackColorAlternate=   15791081
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "收费项目             |单价     |数量 "
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
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
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
      End
      Begin VB.TextBox ctxt标准名称 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin MSComctlLib.TreeView ctvwItem 
         Height          =   5175
         Left            =   4560
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   9128
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标准名称："
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb设置 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView ctvwStandard 
      Height          =   6105
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   10769
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "子系统、使用范围、标准"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frm设置收费标准"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

Private pint科目级数 As Integer

Private Sub ccmdAdd_Click()
    Dim i As Long
    On Error GoTo errhandler
    If ctvwItem.SelectedItem Is Nothing Then
        Err.Raise 6666, , "请选择要添加的收费项目！"
    ElseIf ctvwItem.SelectedItem.Key = "s" Then
        Err.Raise 6666, , "请选择要添加的收费项目！"
    End If
    sub添加当前项 ctvwItem.SelectedItem
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费标准", "ccmdAdd_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub sub添加当前项(ByVal paraNode As Node)
    Dim lobjChild As Node
    Dim i As Long
    On Error GoTo errhandler
    If Len(paraNode.Key) = pint科目级数 * 3 + 1 Then
        '选择的是末级项目
        sub添加指定项目 Right(paraNode.Key, Len(paraNode.Key) - 1)
    Else
        '添加所有下级项目。
        If paraNode.Children > 0 Then
            Set lobjChild = paraNode.Child
            For i = 1 To paraNode.Children
                sub添加当前项 lobjChild
                Set lobjChild = lobjChild.Next
            Next
        End If
    End If
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费标准", "sub添加当前项", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
Private Sub sub添加指定项目(ByVal para编号 As String)
    On Error GoTo errhandler
    '检查该项目是否已添加。
    Dim i As Long
    For i = 1 To cgrdMain.Rows - 1
        If cgrdMain.TextMatrix(i, 3) = para编号 Then
            Exit Sub
        End If
    Next
    
    '需要添加。
    Dim lobjItem As Object
    Set lobjItem = CreateObject("收费对象部件.cls收费项目")
    lobjItem.收费项目编号 = para编号
    cgrdMain.Rows = cgrdMain.Rows + 1
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 0) = lobjItem.收费项目名称
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 1) = lobjItem.单价
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 2) = 1
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 3) = para编号
    
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费标准", "sub添加指定项目", Err.Number, Err.Description, True
    Exit Sub
    Resume
    
End Sub

Private Sub ccmdDel_Click()
    On Error GoTo errhandler
    If cgrdMain.Rows = 1 Then
        MsgBox "没有收费项目可去掉！", vbOKOnly + vbExclamation, "系统提示"
        Exit Sub
    ElseIf cgrdMain.Row < 1 Then
        MsgBox "请在网格中选择要去掉的收费项目！", vbOKOnly + vbExclamation, "系统提示"
        Exit Sub
    End If
    cgrdMain.RemoveItem cgrdMain.Row
    
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费标准", "ccmdDel_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '只能修改单价、数量。
    If Col = 0 Then Cancel = True
End Sub


Private Sub ctvwItem_DblClick()
    ccmdAdd_Click
    
End Sub

Private Sub ctvwStandard_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errhandler
    Dim i As Long
    
    If Left(Node.Key, 1) = "S" Then
        ctxt标准名称 = Node.Text
        ctxt旧标准名称 = Node.Text
        Frame1.Caption = "标准：" & Node.Text
        Frame1.Enabled = True
        
        '显示标准信息。
        Dim lobj标准 As Object
        Dim lcol项目 As Collection
        Set lobj标准 = CreateObject("收费对象部件.cls收费标准")
        lobj标准.收费标准名称 = Node.Text
        Set lcol项目 = lobj标准.收费项目
        cgrdMain.Rows = lcol项目.Count + 1
        For i = 1 To lcol项目.Count
            cgrdMain.TextMatrix(i, 0) = lcol项目(i)("收费项目名称")
            cgrdMain.TextMatrix(i, 1) = lcol项目(i)("单价")
            cgrdMain.TextMatrix(i, 2) = lcol项目(i)("数量")
            cgrdMain.TextMatrix(i, 3) = lcol项目(i)("收费项目编号")
        Next
    Else
        ctxt旧标准名称 = ""
        ctxt标准名称 = ""
        If Node.Parent Is Nothing Then
            Frame1.Caption = "请选择下面的使用范围后添加"
            Frame1.Enabled = False
        Else
            Frame1.Caption = "添加" & Node.Parent.Text & "-" & Node.Text & "的标准"
            Frame1.Enabled = True
        End If
        cgrdMain.Rows = 1
    End If
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费标准", "ctvwStandard_NodeClick", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True

    '初始化工具栏
    Dim lcol工具栏按钮 As Collection
    Set lcol工具栏按钮 = New Collection
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.c工具栏 = ctlb设置

    lcol工具栏按钮.Add "添加"
    lcol工具栏按钮.Add "删除"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "保存"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    '获取所有的收费标准。
    Dim lobjBase As Object
    Dim lstrSysName As String
    Set lobjBase = dafuncGetData("select * from 收费管理_子系统过滤条件表 order by 系统名,条件名")
    Do While Not lobjBase.EOF
        ctvwStandard.Nodes.Add , , lobjBase!系统名, lobjBase!系统名
        lstrSysName = lobjBase!系统名
        Do While lstrSysName = lobjBase!系统名
            ctvwStandard.Nodes.Add lstrSysName, tvwChild, "F" & lobjBase!编号, lobjBase!条件名
            
            '获取该子系统该条件的所有标准。
            Set lobjRec = dafuncGetData("select distinct 收费标准名称 from 收费管理_收费标准信息表 where 过滤条件=" & lobjBase!编号)
            Do While Not lobjRec.EOF
                ctvwStandard.Nodes.Add "F" & lobjBase!编号, tvwChild, "S" & lobjRec!收费标准名称, lobjRec!收费标准名称
                lobjRec.MoveNext
            Loop
            
            ctvwStandard.Nodes("F" & lobjBase!编号).Expanded = True
            
            lobjBase.MoveNext
            If lobjBase.EOF Then Exit Do
            
        Loop
    Loop
    If ctvwStandard.Nodes.Count > 1 Then
        ctvwStandard.Nodes(2).Selected = True
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
    
    '初始化收费项目树。
    Dim lint级数 As Long
    Dim lstrKey As String
    
    pint科目级数 = Val(pobj收费管理.业务设置("科目级数"))
    If pint科目级数 = 0 Then pint科目级数 = 2
    
    ctvwItem.Nodes.Add , , "s", "收费项目"
    For lint级数 = 1 To pint科目级数
        Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where Len(收费项目编号) =" & lint级数 * 3 & " order by 收费项目编号")
        Do While (Not lobjRec.EOF)
            lstrKey = "s" & lobjRec("收费项目编号").Value
            ctvwItem.Nodes.Add "s" & Mid(lstrKey, 2, ((lint级数 - 1) * 3)), tvwChild, lstrKey, lobjRec("收费项目名称").Value
            lobjRec.MoveNext
        Loop
    Next
    ctvwItem.Nodes(1).Expanded = True
    
    cgrdMain.Cols = 4
    cgrdMain.ColHidden(3) = True '保存收费项目编号。
    cgrdMain.Editable = True
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费标准", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    Dim lobj标准 As Object
    On Error GoTo errhandler
    Select Case Operate
    Case "添加"
        Cancel = True
        If ctvwStandard.SelectedItem Is Nothing Then
            Err.Raise 6666, , "请选择使用范围！"
        ElseIf Left(ctvwStandard.SelectedItem.Key, 1) <> "S" And Left(ctvwStandard.SelectedItem.Key, 1) <> "F" Then
            Err.Raise 6666, , "请选择使用范围！"
        End If
        
        ctxt旧标准名称 = ""
        ctxt标准名称.Text = ""
        cgrdMain.Rows = 1
        ctxt标准名称.SetFocus
        Frame1.Caption = "添加标准："
        
    Case "保存"
        Cancel = True
        If ctxt标准名称.Text = "" Then
            ctxt标准名称.SetFocus
            Err.Raise 6666, , "请输入收费标准名称！"
        End If
        If cgrdMain.Rows = 1 Then
            Err.Raise 6666, , "请添加收费项目！"
        End If
        Set lobj标准 = CreateObject("收费对象部件.cls收费标准")
        If ctxt旧标准名称.Text <> "" Then
            lobj标准.收费标准名称 = ctxt旧标准名称
        End If
        If Left(ctvwStandard.SelectedItem.Key, 1) = "S" Then
            lobj标准.过滤条件 = Right(ctvwStandard.SelectedItem.Parent.Key, Len(ctvwStandard.SelectedItem.Parent.Key) - 1)
        Else
            lobj标准.过滤条件 = Right(ctvwStandard.SelectedItem.Key, Len(ctvwStandard.SelectedItem.Key) - 1)
        End If
        
        For i = 1 To cgrdMain.Rows - 1
            lobj标准.sub添加项目 cgrdMain.TextMatrix(i, 3), cgrdMain.TextMatrix(i, 0), cgrdMain.TextMatrix(i, 1), cgrdMain.TextMatrix(i, 2)
        Next
        lobj标准.sub保存 (ctxt标准名称.Text)
        If ctxt旧标准名称.Text <> "" Then
            ctvwStandard.SelectedItem.Text = ctxt标准名称.Text
            ctvwStandard.SelectedItem.Key = "S" & ctxt标准名称.Text
        Else
            '添加节点。
            Dim lstrParent As String
            If Left(ctvwStandard.SelectedItem.Key, 1) = "S" Then
                lstrParent = ctvwStandard.SelectedItem.Parent.Key
            Else
                lstrParent = ctvwStandard.SelectedItem.Key
            End If
            ctvwStandard.Nodes.Add lstrParent, tvwChild, "S" & ctxt标准名称.Text, ctxt标准名称.Text
            ctvwStandard.Nodes("S" & ctxt标准名称.Text).Selected = True
            ctxt旧标准名称.Text = ctxt标准名称.Text
        End If
        
    Case "删除"
        Cancel = True
        
        Set lobj标准 = CreateObject("收费对象部件.cls收费标准")
        If ctxt旧标准名称.Text = "" Then
            Err.Raise 6666, , "请选择要删除的标准！你现在选择的是试用范围。"
        End If
        lobj标准.收费标准名称 = ctxt旧标准名称
        lobj标准.sub删除标准
        
        ctvwStandard.Nodes.Remove ctvwStandard.SelectedItem.Key
        
    End Select

    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费标准", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub
