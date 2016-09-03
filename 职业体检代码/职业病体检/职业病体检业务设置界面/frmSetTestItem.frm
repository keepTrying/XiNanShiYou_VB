VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{DA03AE6C-F6DB-11D1-9E6E-0040053F8E31}#3.3#0"; "DYCOMMINPUT.OCX"
Begin VB.Form frmSetTestItem 
   Caption         =   "体检项目设置"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13200
   ClipControls    =   0   'False
   Icon            =   "frmSetTestItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   13200
   Begin 道源通用录入控件.DyInputGrid cgrdInput 
      Height          =   6975
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   12303
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ErrorColor      =   12648447
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "返回(&X)"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   2355
   End
   Begin MSComctlLib.TreeView ctrwType 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   12303
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
      Caption         =   "说明：选中右边网格中的某行，按《Del》键可以删除一行,直接修改单元格中的单价能够直接修改单价"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   8100
   End
End
Attribute VB_Name = "frmSetTestItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：设置体检项目  界面
'时间：2012-04
'作者：何啸天

Option Explicit
'修改人：张令2012.11.29
'bug号：0000036
'说明：添加变量   ↓↓
Public mstrNode As String
'修改             ↑↑

Private mobj体检项目集 As Object  'ClsTestItemSet,负责获取指定体检大类的体检项目。

'german
'功能：切换是否批量删除标志
'作者：何啸天
Private Sub any_delete_Click()
    If (flag_delete_any = False) Then
        flag_delete_any = True
    Else
        flag_delete_any = False
    End If
End Sub

'german
'功能：调试
'作者：何啸天
Private Sub cgrdInput_Click()
'    Dim loop_a As Integer
'    Dim string_debug As String
'    string_debug = ""
'    'MsgBox cgrdInput.Rows, , "消息"
'    'MsgBox cgrdInput.Cols, , "消息"
'    For loop_a = 0 To cgrdInput.Rows Step 1
'        If (cgrdInput.Value(loop_a, "删除标志") = "1") Then
'        'string_debug = cgrdInput.Value(loop_a, "删除标志") & " " & string_debug
'            'MsgBox "okay", , "okay"
'        End If
'    Next
End Sub

'作者：张令 2012.11.30
'说明：用ascii码控制不能输入特殊符号   ↓↓
'bug号：0000044
Private Sub cgrdInput_ItemKeyPress(ByVal paraItem As String, KeyAscii As Integer)
    Dim a As Integer
    Dim b As Integer
    a = 1
    b = 1
    If KeyAscii >= 48 And KeyAscii <= 57 Or ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123)) Or KeyAscii = 44 Then
    Else
        a = 0
    End If
        
    If KeyAscii >= -20319 And KeyAscii <= -3652 Or KeyAscii = 8 Then
    Else
        b = 0
    End If
    
    If a = 0 And b = 0 Then
    KeyAscii = 0
    End If
End Sub
'                            ↑↑

Private Sub Form_Load()
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    
    '通过字典对象获取"体检大类"的所有id、编号和名称，显示在ctrvType中（key=id）；
    Set lobjRec = pobjDict.Fetch("职业病体检科室字典")
    
    ctrwType.Nodes.Add , , "R", "体检大类"
    Do While Not lobjRec.EOF
        Dim str As String
        str = lobjRec!名称
        If Right(str, 1) = "科" Then
            ctrwType.Nodes.Add "R", tvwChild, "I" & lobjRec!InnerID, lobjRec!编号 & " " & lobjRec!名称
        End If
        lobjRec.movenext
    Loop
    
    '初始化录入网格:编码，名称，缺省值，枚举来源(可选值)，属性。
    'subAddItem,第3个参数true，背景色为浅黄
    With cgrdInput.InputTemplate
        .subAddItem "编码", 0, True, True, 5
        .subAddItem "名称", 0, True, True, 30
        .subAddItem "缺省值", 0, False, True, 300
        .subAddItem "枚举来源", 0, False, True, 500
        .subAddItem "属性", 4, True, True, 10, , "常规,化验" 'dyInputSingleselecttext
        .subAddItem "比较方式", 4, False, True, 20, , "=,≤,≥,≠,属于,不属于,范围"
        .subAddItem "标准值", 0, False, True, 300
        .subAddItem "单位", 0, False, True, 50
        .subAddItem "单价", 0, True, True, 10 'german
    End With
    cgrdInput.subDraw
    cgrdInput.Enabled = False
    
    '创建对象mobj体检项目集。
    Set mobj体检项目集 = CreateObject("职业病对象.clsTestItemSet")
    
    On Error Resume Next
    If ctrwType.Nodes.Count > 0 Then
        ctrwType.Nodes(1).Expanded = True
    End If
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetTestItem", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    insert_danjia '修改单价 'german
    Set mobj体检项目集 = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub
Private Sub cgrdInput_BeforeAddNew(paraCancel As Boolean)
    On Error GoTo errHandler
    cgrdInput.ItemEnabled("编码") = True
    sub填充数据
    Exit Sub
errHandler:
    paraCancel = True
End Sub

'修改者：何啸天
'功能增加：新增条目的时候，加入数据校验，即不能输入无关的其他数据
'3.1
'german
Private Sub cgrdInput_AddNew(paraValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim lstrEnum As String
    Dim tmp_string_1 As String 'german
    Dim loop_a As Integer 'german
    
    On Error GoTo errHandler
    '创建体检项目对象。
    Set lobjItem = CreateObject("职业病对象.clsTestItem")
    
    tmp_string_1 = paraValue("编码")("值") 'german
    'MsgBox tmp_string_1, , "消息"
    '数据校验
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'german
'    For loop_a = 1 To Len(tmp_string_1) Step 1
'        If (Asc(Mid(tmp_string_1, loop_a, 1)) < 48 Or Asc(Mid(tmp_string_1, loop_a, 1)) > 57) Then
'            MsgBox "你的编码值必须为数字，数据存贮请求已经被拒绝", 16, "消息"
'            Exit Sub
'        End If
'    Next
    '取消for循环，直接判断
    '修改：张令
    'bug号：0000044
    '2012.11.29    ↓↓
    If IsNumeric(tmp_string_1) = False Then
        MsgBox "你的编码值必须为数字，数据存贮请求已经被拒绝", 16, "消息"
        Exit Sub
    End If
    '2012.11.29    ↑↑
    If (paraValue("名称")("值") = "") Then
        MsgBox "你的名称不能为空，数据存贮请求已经被拒绝", 16, "消息"
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'german
        '校验单价的数据是否全部为数字，或者小于0
        If (IsNull(paraValue("单价")("值"))) Then
            MsgBox "您填写的单价不能为空", 16, "消息"
            cgrdInput.subRefresh
            Exit Sub
        End If
        
        tmp_string_1 = paraValue("单价")("值") 'german
        
        For loop_a = 1 To Len(tmp_string_1) Step 1
            If (Asc(Mid(tmp_string_1, loop_a, 1)) < 48 Or Asc(Mid(tmp_string_1, loop_a, 1)) > 57) Then
                MsgBox "你填写的单价必须为数字，数据存贮请求已经被拒绝", 16, "消息"
                Exit Sub
            End If
        Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '根据当前录入行信息设置lobjItem的属性。
    With lobjItem
        .编码 = paraValue("编码")("值")
        If .是否存在 Then
            Err.Raise 6666, , "你设置的体检项目已存在，体检项目编码不允许重复。"
        End If
        
        .名称 = paraValue("名称")("值")
        
        .缺省值 = IIf(IsNull(paraValue("缺省值")("值")), "", paraValue("缺省值")("值"))
        lstrEnum = IIf(IsNull(paraValue("枚举来源")("值")), "", paraValue("枚举来源")("值"))
        
        '把中文逗号换成英文逗号。
        lstrEnum = gffuncStrReplace(lstrEnum, "，", ",")
        
        .枚举来源 = lstrEnum
        .属性 = paraValue("属性")("值")
        .体检大类 = Right(ctrwType.SelectedItem.Key, Len(ctrwType.SelectedItem.Key) - 1)
        
        '检查缺省值是否在枚举来源中。
        If .缺省值 <> "" And lstrEnum <> "" Then
            If Right(lstrEnum, 1) <> "," Then lstrEnum = lstrEnum & ","
            If InStr(1, lstrEnum, .缺省值 & ",") = 0 Then
                Err.Raise 6666, , "缺省值必须在枚举来源中。"
            End If
        End If
        
        .比较方式 = IIf(IsNull(paraValue("比较方式")("值")), "", paraValue("比较方式")("值"))
        .标准值 = IIf(IsNull(paraValue("标准值")("值")), "", paraValue("标准值")("值"))
        .单位 = IIf(IsNull(paraValue("单位")("值")), "", paraValue("单位")("值"))
        
        .单价 = paraValue("单价")("值")
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End With
    
    '保存当前行信息到数据库。
    lobjItem.subSave
    
    Exit Sub
errHandler:
    Cancel = True
    ErrorInfo = func错误处理(Err.Number, Err.Description)
End Sub

Private Sub cgrdInput_BeforeRowChange(ByVal paraRow As Long, paraCancel As Boolean)
    On Error GoTo errHandler
    cgrdInput.ItemEnabled("编码") = False
    
    Exit Sub
errHandler:
    paraCancel = True
    
End Sub

Private Sub cgrdInput_ItemChange(ByVal paraItem As String)
    Dim lobjItem As Object
    
    On Error GoTo errHandler
    If cgrdInput.ItemValue(paraItem) <> "" Then
        Select Case paraItem
        Case "编码"
            '创建体检项目对象。
            Set lobjItem = CreateObject("职业病对象.clsTestItem")
            
            '若录入的是"编码"，检查编码是否唯一，若已存在，提示错误。
            lobjItem.编码 = cgrdInput.ItemValue(paraItem)
            If lobjItem.是否存在 Then
                Err.Raise 6666, , "编码“" & cgrdInput.ItemValue(paraItem) & "”已存在。编码不允许重复，请重新输入编码。", sf警告
            End If
        End Select
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    If Err.Number <> 60844 Then
        lstrError = func错误处理(Err.Number, Err.Description)
        sffuncMsg lstrError, sf警告
        'cgrdInput.ItemValue(paraItem) = ""
        cgrdInput.ItemSetfocus paraItem
    End If
    Exit Sub
    Resume
End Sub

Private Sub cgrdInput_RowChange(ByVal paraRow As Long, paraNewValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim lstrEnum As String
    
    On Error GoTo errHandler
    '创建体检项目对象。
    Set lobjItem = CreateObject("职业病对象.clsTestItem")
    'MsgBox "german", , "german"
    '根据当前录入行信息设置lobjItem的属性。
    With lobjItem
        .编码 = cgrdInput.Value(paraRow, "编码")
        .名称 = paraNewValue("名称")("值")
        
        lstrEnum = IIf(IsNull(paraNewValue("枚举来源")("值")), "", paraNewValue("枚举来源")("值"))
        
        '把中文逗号换成英文逗号。
        If lstrEnum <> "" Then
            lstrEnum = gffuncStrReplace(lstrEnum, "，", ",")
        End If
        
        .枚举来源 = lstrEnum
        .缺省值 = IIf(IsNull(paraNewValue("缺省值")("值")), "", paraNewValue("缺省值")("值"))
        .属性 = paraNewValue("属性")("值")
        .体检大类 = Right(ctrwType.SelectedItem.Key, Len(ctrwType.SelectedItem.Key) - 1)
        
        '检查缺省值是否在枚举来源中。
        If .缺省值 <> "" And lstrEnum <> "" Then
            If Right(lstrEnum, 1) <> "," Then lstrEnum = lstrEnum & ","
            If InStr(1, lstrEnum, .缺省值 & ",") = 0 Then
                Err.Raise 6666, , "缺省值必须在枚举来源中。"
            End If
        End If
        .比较方式 = IIf(IsNull(paraNewValue("比较方式")("值")), "", paraNewValue("比较方式")("值"))
        .标准值 = IIf(IsNull(paraNewValue("标准值")("值")), "", paraNewValue("标准值")("值"))
        .单位 = IIf(IsNull(paraNewValue("单位")("值")), "", paraNewValue("单位")("值"))
        
    End With
    
    '保存当前行信息到数据库。
    lobjItem.subSave

    Set lobjItem = Nothing
    Exit Sub
errHandler:
    ErrorInfo = func错误处理(Err.Number, Err.Description)
    Cancel = True
End Sub

'修改者：何啸天
'german
'修改内容 添加彻底删除功能
Private Sub cgrdInput_Delete(ByVal paraRow As Long, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim loop_a As Integer
    
    On Error GoTo errHandler
    '创建体检项目对象。
    Set lobjItem = CreateObject("职业病对象.clsTestItem")
    
    '根据当前录入行信息设置lobjItem的属性。
    lobjItem.编码 = cgrdInput.Value(paraRow, "编码") '缓存用户选择列的编码值，在删除此条数据的时候用于定位
    
    'MsgBox CStr(paraRow), , "german" 'german
    
    lobjItem.subDelete '从数据库中将此条记录删除
    '删除库中的该项目。
'    If (flag_delete_any = False) Then 'german 是否批量删除数据，如果不是那么就按照单条数据删除流程走
'        lobjItem.subDelete (flag_database_delete) 'german 参数1，为真则在按下DEL键的同时，将数据库中的这项内容删除 假则反之
'    Else '批量删除数据,在这里循环扫描用户指定要删除的数据，然后依次删除
'        For loop_a = 0 To cgrdInput.Rows Step 1
'            If (cgrdInput.Value(loop_a, "删除标志") = "1") Then
'            'string_debug = cgrdInput.Value(loop_a, "删除标志") & " " & string_debug
'                'MsgBox "okay", , "okay"
'                lobjItem.编码 = cgrdInput.Value(loop_a, "编码")
'                lobjItem.subDelete_any
'            End If
'        Next
        
    'End If
        
    Set lobjItem = Nothing
    Exit Sub
errHandler:
    ErrorInfo = func错误处理(Err.Number, Err.Description)
    Cancel = True
End Sub

Private Sub ctrwType_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lobjRec As Object          '当前选中大类的所有体检项目。
    Dim lcolInfo As Collection     '加入到录入网格中一行的内容。
    
    On Error GoTo errHandler
    insert_danjia '修改单价 'german
    cgrdInput.subClear
    
    If Node.Parent Is Nothing Then
        '选中大类型节点，不允许录入。
        cgrdInput.Enabled = False
    Else
        cgrdInput.Enabled = True
        
        '修改人：张令2012.11.29
        'bug号：0000036
        '修改说明：mstrNode获取点击行数据，通过“sub填充数据”获取数据   ↓↓
'        '设置mobj体检项目集.体检大类=当前节点的key；
'        mobj体检项目集.体检大类 = Right(Node.Key, Len(Node.Key) - 1) 'ClsTestItemSet,负责获取指定体检大类的体检项目。
'
'        'see
'        '获取指定体检大类的体检项目：编码，名称，缺省值，枚举来源，属性，体检大类。
'        Set lobjRec = mobj体检项目集.体检项目 '返回一个数据库对象
'
'        '把lobjRec中所有所有记录显示在cgrdInput中。
'        gfsubLoadDyGridFromRec cgrdInput, lobjRec '传入树形列表以及数据库对象
'        cgrdInput.subExpand
'        lobjRec.Close
        mstrNode = Node.Key
        sub填充数据
        '2012.11.29     ↑↑
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmSetTestItem", "ctrwType_NodeClick", Err.Number, Err.Description, False
End Sub

'修改人：张令2012.11.29
'bug号：0000036
'修改说明：刷新函数
Private Sub sub填充数据()
    Dim lobjRec As Object
    '设置mobj体检项目集.体检大类=当前节点的key；
    mobj体检项目集.体检大类 = Right(mstrNode, Len(mstrNode) - 1) 'ClsTestItemSet,负责获取指定体检大类的体检项目。
    
    'see
    '获取指定体检大类的体检项目：编码，名称，缺省值，枚举来源，属性，体检大类。
    Set lobjRec = mobj体检项目集.体检项目 '返回一个数据库对象
    
    '把lobjRec中所有所有记录显示在cgrdInput中。
    gfsubLoadDyGridFromRec cgrdInput, lobjRec '传入树形列表以及数据库对象
    cgrdInput.subExpand
    lobjRec.Close
End Sub

Private Sub cgrdInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
    
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmSetTestItem", "cgrdInput_KeyDown", Err.Number, Err.Description, False
End Sub
Private Sub ccmdExit_Click()
    ccmdExit.Caption = "正在保存数据...."
    insert_danjia '执行单价修改过程
    ccmdExit.Caption = "返回(&X)"
    Unload Me
End Sub

'作者：何啸天
'功能：切换是否强制删除数据功能标志
'german
'Private Sub is_delete_Click()
'    If flag_database_delete = False Then
'        flag_database_delete = True
'        'MsgBox "变换为真", , "消息"
'    Else
'        flag_database_delete = False
'        'MsgBox "变换为假", , "消息"
'    End If
'End Sub

'作者：何啸天
'功能：批量单价修改功能
'日期:3.13
'german
Private Sub insert_danjia()
    Dim loop_a As Integer
    Dim lobjItem
    
    On Error Resume Next
    Set lobjItem = CreateObject("职业病对象.clsTestItem")
    For loop_a = 1 To cgrdInput.Rows - 1 Step 1
        lobjItem.编码 = CStr(cgrdInput.Value(loop_a, "编码"))
        lobjItem.单价 = cgrdInput.Value(loop_a, "单价")
        lobjItem.SubSaveUnitprice
    Next
End Sub

