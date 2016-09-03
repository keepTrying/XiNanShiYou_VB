VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectItem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择复查项目、收费项目"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   3975
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrd收费项目 
      Height          =   3615
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      _cx             =   59120272
      _cy             =   59119848
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
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
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
      FormatString    =   "项目编号     |项目名称            |单价"
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
      Begin VB.Label clblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "“业务设置”中设置了不收费，所以不需要选择收费项目！"
         ForeColor       =   &H00800000&
         Height          =   2340
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3420
      End
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.TreeView ctrwItem 
      Height          =   3645
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6429
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
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
   Begin VB.Label clblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "说明：上面列出的是“”上的复查项目，若你发现项目不够，请进入“体检表设置”操作，把你需要的项目加入该体检表。"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(元)"
      Height          =   180
      Index           =   3
      Left            =   6720
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label clblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5400
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费总额："
      Height          =   180
      Index           =   2
      Left            =   4440
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择收费项目："
      Height          =   180
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择复查项目："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frmSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'名称：体检项目选定
'功能：选择体检项目
'函数：
'作者：刘云乐
'时间：2012-03
Option Explicit

'2012-08-22 于登淼 ↓
'暂停X毫秒。用于体检项选择时，防止点击速度过快，导致控件反应不过来。
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'2012-08-22 于登淼 ↑

Public pstr体检表名称 As String
Public pcol复查项目 As Collection  '返回结果[编码，名称],key:编码。
Public pcol收费项目 As Collection  '返回结果[收费项目编号,单价]key:编号。
Public pblnOk As Boolean           '是否确定返回。

Private Sub ccmdCancel_Click()
    On Error Resume Next
    pblnOk = False
    Set pcol收费项目 = New Collection
    Unload Me
End Sub

Public Sub ccmdOk_Click()
    Dim lobjNode As Node
    Dim lcolInfo As Collection
    Dim lcolItem As Collection
    Dim lstrItem As String
    
    Dim i As Long
    
    On Error GoTo errHandler
    '获取选择的项目。
    Set lcolInfo = New Collection
    For Each lobjNode In ctrwItem.Nodes
        If lobjNode.Checked And Not lobjNode.Parent Is Nothing Then
            lstrItem = Right(lobjNode.Key, Len(lobjNode.Key) - 1)
            Set lcolItem = New Collection
            lcolItem.Add lstrItem, "编码"
            lcolItem.Add Right(lobjNode.Text, Len(lobjNode.Text) - InStr(lobjNode.Text, " ")), "名称"
            lcolInfo.Add lcolItem, lstrItem
        End If
    Next
    Set lobjNode = Nothing
'    If lcolInfo.Count = 0 Then
'        sffuncMsg "必须选择体检项目！", sf警告
'        Set lcolInfo = Nothing
'        Exit Sub
'    End If
    
'    '复查表    2016-6-15 by 牟俊
'    Dim par复查系统编号 As String
'    Dim obj As Object
'    par复查系统编号 = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.Row, frmFinalConclusion.cgrdInfo.ColIndex("系统编号")) & "F"
'    Set obj = dafuncGetData("select * from 职业病体检_复查项目表 where 系统编号='" & par复查系统编号 & "'")
'        If obj.RecordCount > 0 Then
'            dafuncGetData ("delete from 职业病体检_复查项目表 where 系统编号='" & par复查系统编号 & "'")
'        End If
'    For Each lobjNode In ctrwItem.Nodes
'        If lobjNode.Checked And Not lobjNode.Parent Is Nothing Then
'            lstrItem = Right(lobjNode.Key, Len(lobjNode.Key) - 1)
'            dafuncGetData ("insert into 职业病体检_复查项目表 values ('" & par复查系统编号 & "','" & lstrItem & "')")
'        End If
'    Next
    
    
    '获取收费项目。
    'Set pcol收费项目 = New Collection
    'If pobj业务对象.业务设置("是否收费") = "是" Then
    '    For i = 1 To cgrd收费项目.Rows - 1
    '        If cgrd收费项目.Cell(flexcpChecked, i, 0, i, 0) = flexChecked Then
    '            Set lcolItem = New Collection
    '            lcolItem.Add cgrd收费项目.TextMatrix(i, 0), "收费项目编号"
    '            lcolItem.Add cgrd收费项目.ValueMatrix(i, 2), "单价"
    '            pcol收费项目.Add lcolItem, lcolItem("收费项目编号")
    '        End If
    '    Next
    '    If pcol收费项目.Count = 0 Then
    '        If MsgBox("业务设置了要收费，你确认该人员不收费吗？" & Chr(13) & Chr(10) & "若无法选择收费项目，请进入“业务设置”操作，设置不收费；或按界面上的蓝色提示信息进行处理！", vbYesNo + vbQuestion + vbDefaultButton2, "系统询问") = vbNo Then
    '            Exit Sub
    '        End If
    '    End If
    'End If
    
    Set pcol复查项目 = lcolInfo
    pblnOk = True
    Unload Me

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmSelectItem", "ccmdOk_Click", 6666, lstrError, False

End Sub

Private Sub cgrd收费项目_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim dblTotal As Double
    
    On Error Resume Next
    '计算总金额
    For i = 1 To cgrd收费项目.rows - 1
        If cgrd收费项目.Cell(flexcpChecked, i, 0) = flexChecked Then
            dblTotal = Format(dblTotal + cgrd收费项目.ValueMatrix(i, 2), "0.00")
        End If
    Next
    
    clblTotal.Caption = dblTotal
End Sub

Private Sub cgrd收费项目_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Row > 0 And Col = 0 Then
    Else
        Cancel = True
    End If
End Sub
'功能：若对父节点进行操作，自动对子节点进行操作，反之。
Private Sub ctrwItem_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Dim lobjNode As Node
    Dim i As Long
    
    If Node.Parent Is Nothing Then
        '当前操作的是父节点。
        '自动选中(或不选中)所有子节点。
        If Node.Children > 0 Then
            For i = 1 To Node.Children
                Set lobjNode = ctrwItem.Nodes(Node.Index + i)
                lobjNode.Checked = Node.Checked
            Next
        End If
    Else
        '当前操作的是子节点。
        If Node.Checked Then
            Node.Parent.Checked = True
        Else
            For i = 1 To Node.Parent.Children
                If ctrwItem.Nodes(Node.Parent.Index + i).Checked Then
                    Exit For
                End If
            Next
            If i > Node.Parent.Children Then
                Node.Parent.Checked = False
            End If
        End If
    End If
    
    '2012-08-22 翁乔 ↑
    '控制控件相应速度，防止速度过快，造成项目保存错误。本机上350ms一般情况下没有问题，点击非常快的话，仍然会有选错的可能。
    ctrwItem.Visible = False
    Sleep (500)
    ctrwItem.Visible = True
    '2012-08-22 翁乔 ↑
End Sub

Private Sub Form_Load()
    Dim lobj体检模板  As Object
    Dim lobjDict As Object
    Dim lobj体检项目集 As Object
    Dim lobjRec As Object
    Dim lobjItem As Object
    Dim lobjNode As Node
    Dim lcolInfo As Collection
    Dim ldblTotal As Double
    Dim i As Long
    Dim lbln全部选中 As Boolean
    
    On Error GoTo errHandler
    '确定体检表名称  2016-6-13
    pstr体检表名称 = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.Row, frmFinalConclusion.cgrdInfo.ColIndex("体检表编号"))
    clblMsg.Caption = "说明：上面列出的是“" & pstr体检表名称 & "”上的体检项目，若你发现项目不够，请进入“体检表设置”操作，把你需要的项目加入该体检表。"
    
    '创建体检表模板对象。
    Set lobj体检模板 = CreateObject("职业病对象.clsMedicalExamTemplate")
    lobj体检模板.体检表名 = pstr体检表名称
    
    '获取复查体检表的体检项目。
    Set lcolInfo = lobj体检模板.体检项目集
    
    '创建字典对象。
    Set lobjDict = CreateObject("字典管理.clsDictionary")
        
    lbln全部选中 = True
    
    '创建体检项目集对象。
    Set lobj体检项目集 = CreateObject("职业病对象.clsTestItemSet")
    
    '获取所有体检大类、体检项目。
    Set lobjRec = lobjDict.Fetch("职业病体检科室字典")
    '显示体检大类在ctrvItem中（其中节点的key=体检大类id）；
    ctrwItem.Nodes.Clear
    Do While Not lobjRec.EOF
        '通过"lobj体检项目集"依次获取各大类的体检项目。
        lobj体检项目集.体检大类 = lobjRec("InnerID")
        Set lobjItem = lobj体检项目集.体检项目
        If Not lobjItem.EOF Then
            ctrwItem.Nodes.Add , , "I" & lobjRec("InnerID"), lobjRec("编号") & " " & Trim(lobjRec("名称"))
        End If
        
        '显示体检项目在ctrvItem中（其中节点的key=编码,parent=体检大类）。
        Do While Not lobjItem.EOF
            If sffunc判断集合键值是否存在(lcolInfo, lobjItem("编码")) Then
                Set lobjNode = ctrwItem.Nodes.Add("I" & lobjRec("InnerID"), tvwChild, "I" & lobjItem("编码"), lobjItem("编码") & " " & Trim(lobjItem("名称")))
                
                '选中已有的复查项目。
                If Not pcol复查项目 Is Nothing Then
                    If sffunc判断集合键值是否存在(pcol复查项目, lobjItem("编码")) Then
                        lobjNode.Checked = True
                        lobjNode.Parent.Checked = True
                    Else
                        lbln全部选中 = False
                    End If
                    
                End If
            End If
            
            lobjItem.MoveNext
        Loop

        If lobjItem.RecordCount > 0 Then
            If ctrwItem.Nodes("I" & lobjRec("InnerID")).Children = 0 Then
                ctrwItem.Nodes.Remove "I" & lobjRec("InnerID")
            End If
        End If

        lobjRec.MoveNext
    Loop
    
    '判断是否收费，若要，可以设置收费项目。
    clblInfo.Visible = False
    '修改：2002-10-28（杨春）为了嘉定定制需求，不传递内部收费信息时，仍可以选择收费项目。
    'If pobj业务对象.业务设置("是否收费") = "是" Then
    '    If lobj体检模板.收费标准 = "" Then
    '        clblInfo.Caption = "没有设置体检表的收费标准，无法选择收费项目！请先进入“体检表设置”操作，进行收费标准设置！"
    '        clblInfo.Visible = True
    '    Else
    '        '获取该收费标准的收费项目。
    '        Set lobjRec = pobj业务对象.收费标准的项目(lobj体检模板.收费标准)
    '        cgrd收费项目.Rows = lobjRec.RecordCount + 1
    '        i = 1
    '        ldblTotal = 0
    '        Do While Not lobjRec.EOF
    '            cgrd收费项目.TextMatrix(i, 0) = lobjRec("收费项目编号")
    '            cgrd收费项目.TextMatrix(i, 1) = IIf(IsNull(lobjRec("收费项目名称")), "", lobjRec("收费项目名称"))
    '            cgrd收费项目.TextMatrix(i, 2) = IIf(IsNull(lobjRec("单价")), 0, lobjRec("单价"))
    '
    '            If sffunc判断集合键值是否存在(pcol收费项目, lobjRec("收费项目编号")) Or lbln全部选中 Then
    '                cgrd收费项目.Cell(flexcpChecked, i, 0, i, 0) = flexChecked
    '                ldblTotal = Format(ldblTotal + IIf(IsNull(lobjRec("单价")), 0, lobjRec("单价")), "0.00")
    '            Else
    '                cgrd收费项目.Cell(flexcpChecked, i, 0, i, 0) = flexUnchecked
    '            End If
    '
    '            lobjRec.MoveNext
    '            i = i + 1
    '        Loop
    '        If cgrd收费项目.Rows > 1 Then
    '            cgrd收费项目.Editable = True
    '            clblTotal.Caption = ldblTotal
    '        Else
    '            clblInfo.Caption = "当前体检表的收费标准没有任何收费项目，无法选择收费项目。请收费科室先进行收费标准设置。"
    '            clblInfo.Visible = True
    '        End If
    '    End If
    'Else
    '    clblInfo.Caption = "“业务设置”中设置了不收费，所以不需要选择收费项目！"
    '    clblInfo.Visible = True
    'End If
    
    Set lobjDict = Nothing

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frmSelectItem", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

