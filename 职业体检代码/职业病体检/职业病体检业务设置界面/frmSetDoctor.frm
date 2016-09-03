VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSetDoctor 
   Caption         =   "医师体检项目设置"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11295
   Icon            =   "frmSetDoctor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   11295
   Begin VB.CommandButton ccmdExit 
      Caption         =   "返回(&X)"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   1275
   End
   Begin MSComctlLib.TreeView ctrwAllItem 
      Height          =   5925
      Left            =   6480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10451
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
   Begin VB.CommandButton ccmdDel 
      Caption         =   ">> 去掉"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton ccmdAdd 
      Caption         =   "<< 添加"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   855
   End
   Begin MSComctlLib.TreeView ctrwDoctor 
      Height          =   5895
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   10398
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
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
   Begin MSComctlLib.TreeView ctrwSelectedItem 
      Height          =   5895
      Left            =   2280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10398
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
      Caption         =   "说明：设置各医师可作的项目后，医师在填写常规或化验结果时，只能看见并填写其可作的项目。"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   6840
      Width           =   7740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "所有可选体检项目(双击单项可以添加)："
      Height          =   180
      Left            =   6480
      TabIndex        =   6
      Top             =   240
      Width           =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "可作体检项目列表："
      Height          =   180
      Left            =   2280
      TabIndex        =   2
      Top             =   195
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检医师列表："
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   195
      Width           =   1260
   End
End
Attribute VB_Name = "frmSetDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春

Private mobj体检医师  As Object 'ClsMedicalExamer,负责增加、删除体检医师可作的体检项目。

Private Sub ctrwSelectedItem_DblClick()
    On Error Resume Next
    '叶节点可以双击删除。
    If ccmdDel.Enabled = True Then
        If Not ctrwSelectedItem.Parent Is Nothing Then
            If Not ctrwSelectedItem.Parent.Parent Is Nothing Then
                ccmdDel_Click
            End If
        End If
    End If
End Sub

Private Sub ctrwSelectedItem_NodeClick(ByVal Node As MSComctlLib.Node)
    ccmdDel.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If (Shift & vbAltMask) = vbAltMask And KeyCode = vbKeyX Then
        Unload Me
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub
Private Sub Form_Load()
    Dim lobj体检项目集 As Object 'clsTestItemSet。
    Dim lobjRec As Object        '获取的所有用户记录集，体检大类字典项目集。
    Dim lobjItem As Object       '某大类的体检项目记录集。
    
    On Error GoTo errHandler
    '调用“用户管理.umfunc获取所有用户”获取所有用户列表。
    Set lobjRec = pobjDict.FetchEx("员工字典")
    
    '显示在ctrvDoctor中（期中节点的key=用户编号）。
    '显示用户在医师树中。
    ctrwDoctor.Nodes.Add , , "R", "体检医师"
    Do While Not lobjRec.EOF
        '修改：2001-11-7（杨春）不显示0000、gues。
        If lobjRec("编号") <> "0000" And lobjRec("编号") <> "gues" Then
            ctrwDoctor.Nodes.Add "R", tvwChild, "I" & lobjRec("编号"), lobjRec("编号") & " " & lobjRec("姓名")
        End If
        lobjRec.movenext
    Loop
    If ctrwDoctor.Nodes.Count > 0 Then
        ctrwDoctor.Nodes(1).Expanded = True
    End If
    
    '创建体检项目集对象。
    Set lobj体检项目集 = CreateObject("职业病对象.clsTestItemSet")
    
    '获取所有体检大类、体检项目。
    Set lobjRec = pobjDict.Fetch("职业病体检科室字典")
    
    '显示体检大类在ctrvItem中（其中节点的key=体检大类id）；
    ctrwAllItem.Nodes.Add , , "R", "体检大类"
    Do While Not lobjRec.EOF
        
        '通过"lobj体检项目集"依次获取各大类的体检项目。
        lobj体检项目集.体检大类 = lobjRec("InnerID")
        Set lobjItem = lobj体检项目集.体检项目
        If Not lobjItem.EOF Then
            ctrwAllItem.Nodes.Add "R", tvwChild, "I" & lobjRec("InnerID"), lobjRec("编号") & " " & lobjRec("名称")
        End If
        '显示体检项目在ctrvItem中（其中节点的key=编码,parent=体检大类）。
        Do While Not lobjItem.EOF
            ctrwAllItem.Nodes.Add "I" & lobjRec("InnerID"), tvwChild, "I" & lobjItem("编码"), lobjItem("编码") & " " & lobjItem("名称")
            lobjItem.movenext
        Loop
        
        lobjRec.movenext
    Loop
    
    '创建对象"mobj体检医师"。
    Set mobj体检医师 = CreateObject("职业病对象.clsMedicalExaminer")
    
    On Error Resume Next
    If ctrwAllItem.Nodes.Count > 0 Then
        ctrwAllItem.Nodes(1).Expanded = True
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctor", "Form_load", 6666, lstrError, False
End Sub

Private Sub ctrwAllItem_DblClick()
    On Error Resume Next
    '叶节点可以双击删除。
    If ccmdAdd.Enabled = True Then
        If Not ctrwAllItem.Parent Is Nothing Then
            If Not ctrwAllItem.Parent.Parent Is Nothing Then
                ccmdAdd_Click
            End If
        End If
    End If
End Sub


Private Sub ccmdAdd_Click()
    Dim lstr编码 As String   '体检项目编码。
    Dim lobjNode As Node
    Dim i As Long
    
    On Error GoTo errHandler
    
    MousePointer = 11
    Me.Enabled = False
    
    If ctrwAllItem.SelectedItem.Children = 0 Then
        '单个项目。
        lstr编码 = ctrwAllItem.SelectedItem.Key
        lstr编码 = Right(lstr编码, Len(lstr编码) - 1)
        If lstr编码 = "" Then Exit Sub
        
        '判断该项目是否已在。
        If Not mobj体检医师.func是否可作项目(lstr编码) Then
            '从库中添加该项目设置。
            mobj体检医师.Sub添加体检项目 lstr编码
        
            '把大类加入。
            On Error Resume Next
            If ctrwSelectedItem.Nodes.Count = 0 Then
                ctrwSelectedItem.Nodes.Add , , "R", "体检大类"
            End If
            ctrwSelectedItem.Nodes.Add "R", tvwChild, ctrwAllItem.SelectedItem.Parent.Key, ctrwAllItem.SelectedItem.Parent.Text
            '加入体检项目（其中节点的key=编码,parent=体检大类）。
            ctrwSelectedItem.Nodes.Add ctrwAllItem.SelectedItem.Parent.Key, tvwChild, ctrwAllItem.SelectedItem.Key, ctrwAllItem.SelectedItem.Text
            
        End If
        
    Else '整个项目大类或整个树。
        
        If ctrwAllItem.SelectedItem.Parent Is Nothing Then
            '整个树加入。
            If ctrwAllItem.SelectedItem.Children > 0 Then
                '加入根节点。
                On Error Resume Next
                ctrwSelectedItem.Nodes.Add , , "R", "体检大类"
                
                On Error GoTo errHandler
                '依次加入大类。
                Set lobjNode = ctrwAllItem.SelectedItem.Child
                For i = 1 To ctrwAllItem.SelectedItem.Children
                    sub添加大类 lobjNode
                    Set lobjNode = lobjNode.Next
                Next
            End If
        Else
            '某个大类加入。
            sub添加大类 ctrwAllItem.SelectedItem
        End If
        
    End If
        
    ccmdAdd.Enabled = False
    MousePointer = 0
    Me.Enabled = True
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    MousePointer = 0
    Me.Enabled = True
    sfsub错误处理 "职业病设置界面", "frmSetDoctor", "ccmdAdd_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub
Private Sub sub添加大类(ByVal paraAllItemNode As Node)
    Dim i As Long
    Dim lstr编码 As String
    Dim lobjNode As Node
    
    If paraAllItemNode.Children = 0 Then Exit Sub
    
    '把大类加入。
    On Error Resume Next
    If ctrwSelectedItem.Nodes.Count = 0 Then
        ctrwSelectedItem.Nodes.Add , , "R", "体检大类"
    End If
    
    '把大类加入。
    ctrwSelectedItem.Nodes.Add "R", tvwChild, paraAllItemNode.Key, paraAllItemNode.Text
    
    On Error GoTo errHandler
    
    '依次把该大类的项目叶节点加入。
    For i = 1 To paraAllItemNode.Children
        Set lobjNode = ctrwAllItem.Nodes(paraAllItemNode.Index + i)
        
        lstr编码 = lobjNode.Key
        lstr编码 = Right(lstr编码, Len(lstr编码) - 1)
        
        If lstr编码 <> "" Then
            If Not mobj体检医师.func是否可作项目(lstr编码) Then
                '从库中添加该项目设置。
                mobj体检医师.Sub添加体检项目 lstr编码
            
                '加入体检项目（其中节点的key=编码,parent=体检大类）。
                On Error Resume Next
                ctrwSelectedItem.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
                On Error GoTo errHandler
            End If
        End If
    Next
    Exit Sub
errHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

'功能：删除单个项目、一个大类、所有项目。
Private Sub ccmdDel_Click()
    Dim llngIndex As Long
    Dim lobjNode As Node
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwSelectedItem.SelectedItem Is Nothing Then Exit Sub
    
    MousePointer = 11
    Me.Enabled = False
    If ctrwSelectedItem.SelectedItem.Parent Is Nothing Then
        '删除整个树。
        
        '从库中删除当前医师可作的所有体检项目设置。
        mobj体检医师.Sub删除所有体检项目
    
        ctrwSelectedItem.Nodes.Clear
        
    Else
        If Not ctrwSelectedItem.SelectedItem.Parent.Parent Is Nothing Then
            '删除单个项目。
            Set lobjNode = ctrwSelectedItem.SelectedItem
            
            '从库中删除当前医师可作的该项目设置。
            mobj体检医师.Sub删除体检项目 Right(lobjNode.Key, Len(lobjNode.Key) - 1)
            
            '若父节点下没有子节点，删除父节点。
            If lobjNode.Parent.Children = 0 Then
                ctrwSelectedItem.Nodes.Remove lobjNode.Parent.Key
            Else
                '删除当前选中节点。
                ctrwSelectedItem.Nodes.Remove lobjNode.Key
            End If
        Else
            '删除整个大类。
            For i = 1 To ctrwSelectedItem.SelectedItem.Children
                Set lobjNode = ctrwSelectedItem.Nodes(ctrwSelectedItem.SelectedItem.Index + i)
                '从库中删除当前医师可作的该项目设置。
                mobj体检医师.Sub删除体检项目 Right(lobjNode.Key, Len(lobjNode.Key) - 1)
            Next
            
            '若父节点下没有子节点，删除父节点。
            If ctrwSelectedItem.SelectedItem.Parent.Children = 0 Then
                ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Parent.Key
            Else
                '删除当前选中节点。
                ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Key
            End If
        End If
        
        
        '若整个树空了，删除根节点。
        If ctrwSelectedItem.Nodes.Count = 1 Then
            ctrwSelectedItem.Nodes.Clear
        End If
        
    End If
    
    ccmdDel.Enabled = False
    MousePointer = 0
    Me.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    MousePointer = 0
    Me.Enabled = True
    sfsub错误处理 "职业病设置界面", "frmSetDoctor", "ccmdDel_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub ctrwDoctor_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lcolItemSet As Collection '当前医师可作项目的集合。
    Dim lcolItem As Variant       'lcolItemSet中的某个元素[编码，名称]。
    
    On Error GoTo errHandler
    
    ctrwSelectedItem.Nodes.Clear
    ccmdDel.Enabled = False
    ccmdAdd.Enabled = False
    
    If Node.Parent Is Nothing Then
        Label2 = "可作体检项目列表："
        mobj体检医师.编号 = ""
        Exit Sub
    End If
    
    '设置mobj体检医师.编号=当前节点的key。
    mobj体检医师.编号 = Right(Node.Key, Len(Node.Key) - 1)
    Label2 = Right(Node.Text, Len(Node.Text) - InStr(Node.Text, " ")) & "可作体检项目列表："
    
    '获取当前医师可作的所有体检项目
    Set lcolItemSet = mobj体检医师.可作体检项目
    
    If lcolItemSet.Count > 0 Then
        ctrwSelectedItem.Nodes.Add , , "R", "体检大类"
    End If
    
    '显示项目在clstItem中(编码+''+名称)。
    On Error Resume Next
    Dim lobjNode As Node
    For Each lcolItem In lcolItemSet
    
        '获取该项目在ctrwAllItem树中的节点。
        Set lobjNode = Nothing
        Set lobjNode = ctrwAllItem.Nodes("I" & lcolItem("编码"))
        
        If Not lobjNode Is Nothing Then
            '加入大类节点。
            ctrwSelectedItem.Nodes.Add "R", tvwChild, lobjNode.Parent.Key, lobjNode.Parent.Text
            '加入体检项目（其中节点的key=编码,parent=体检大类）。
            ctrwSelectedItem.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
        End If
        
    Next
    
    On Error Resume Next
    If ctrwSelectedItem.Nodes.Count > 0 Then
        ctrwSelectedItem.Nodes(1).Expanded = True
    End If
    
    ccmdDel.Enabled = False
    If ctrwAllItem.SelectedItem Is Nothing Then
        ccmdAdd.Enabled = False
    Else
        ccmdAdd.Enabled = True
    End If
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetDoctor", "ctrwDoctor_NodeClick", 6666, lstrError, False
End Sub

Private Sub ccmdExit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub clstItem_Click()
    On Error Resume Next
    ccmdDel.Enabled = True
End Sub

Private Sub ctrwAllItem_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errHandler
    
    If mobj体检医师.编号 = "" Then
        ccmdAdd.Enabled = False
    Else
        ccmdAdd.Enabled = True
    End If
    
    Exit Sub
errHandler:
    'sfsub错误处理 "职业病设置界面", "frmSetDoctor", "ctrwAllItem_NodeClick", Err.Number, Err.Description, False
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj体检医师 = Nothing
End Sub

