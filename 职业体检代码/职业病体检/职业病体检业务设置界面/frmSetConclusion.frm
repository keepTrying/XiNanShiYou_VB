VERSION 5.00
Object = "{DA03AE6C-F6DB-11D1-9E6E-0040053F8E31}#3.3#0"; "DYCOMMINPUT.OCX"
Begin VB.Form frmSetConclusion 
   Caption         =   "最终结论模板设置"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10830
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmd返回 
      Caption         =   "返回(&X)"
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "当前信息"
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.Label LblDate 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label LblNo 
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label LblName 
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "当前时间："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "医师编号："
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "医师姓名："
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "结论列表"
      Height          =   5895
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin 道源通用录入控件.DyInputGrid cgrdInput 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9128
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明：选中右边网格中的某行，按《Del》键可以删除一行,直接修改单元格中的结论能够直接修改结论"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   5520
         Width           =   8100
      End
   End
End
Attribute VB_Name = "frmSetConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ccmd返回_Click()
    Unload Me
End Sub

Private Sub cgrdInput_BeforeAddNew(paraCancel As Boolean)
    On Error GoTo errHandler
    '张令 2012.12.28  ↓↓
    '说明：在addnew之后查询一次数据填充到表格，使其表格自适应数据。
    Set lobj结论模板 = CreateObject("职业病对象.clsConclusionSet")
    Set lobjRec = lobj结论模板.func获取最终结论模板
    
    gfsubLoadDyGridFromRec cgrdInput, lobjRec '传入树形列表以及数据库对象
    cgrdInput.subExpand
    lobjRec.Close
    '张令 2012.12.28  ↑↑
    Exit Sub
errHandler:
    paraCancel = True
End Sub

'作者：翁乔
'功能：新增条目的时候，加入数据校验，即不能输入无关的其他数据
'时间：2012-05-29
Private Sub cgrdInput_AddNew(paraValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobj结论模板 As Object
    Dim lobjRec As Object
    Dim flag As Boolean
    Dim flag1 As Boolean
    
    On Error GoTo errHandler

    '校验结论模板的数据是否为空
    If (IsNull(paraValue("结论模板")("值"))) Or (paraValue("结论模板")("值") = "") Then
        MsgBox "您填写的结论模板不能为空", 16, "消息"
        cgrdInput.subRefresh
        Exit Sub
    End If
    
    Set lobj结论模板 = CreateObject("职业病对象.clsConclusionSet")
    
    flag = lobj结论模板.func是否存在(paraValue("结论模板")("值"), paraValue("结论标准")("值"))
    If flag = False Then
        flag1 = lobj结论模板.func添加最终结论模板("16", "最终结论录入", paraValue("结论模板")("值"), um用户名, paraValue("结论标准")("值"))
    
        If flag1 = False Then
            MsgBox "填写有误、请输入有信息，不包含特殊字符！", 16, "消息"
            cgrdInput.subRefresh
            Exit Sub
        End If
    End If
    Set lobj结论模板 = Nothing
    Exit Sub
errHandler:
    Cancel = True
    ErrorInfo = func错误处理(Err.Number, Err.Description)
End Sub

Private Sub cgrdInput_BeforeRowChange(ByVal paraRow As Long, paraCancel As Boolean)
    On Error GoTo errHandler
    
    cgrdInput.ItemEnabled("结论日期") = False
    cgrdInput.ItemEnabled("结论医师") = False
    
    Exit Sub
errHandler:
    paraCancel = True
    
End Sub

Private Sub cgrdInput_ItemChange(ByVal paraItem As String)
    
    On Error GoTo errHandler
'    If cgrdInput.ItemValue(paraItem) <> "" Then
'        Select Case paraItem
'        Case "结论模板"
''            '创建体检项目对象。
''            Set lobjItem = CreateObject("职业病对象.clsTestItem")
''
''            '若录入的是"编码"，检查编码是否唯一，若已存在，提示错误。
''            lobjItem.编码 = cgrdInput.ItemValue(paraItem)
''            If lobjItem.是否存在 Then
''                Err.Raise 6666, , "编码“" & cgrdInput.ItemValue(paraItem) & "”已存在。编码不允许重复，请重新输入编码。", sf警告
''            End If
'        End Select
'    End If
    
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

'作者：张令 2012.11.30
'说明：用ascii码控制不能输入特殊符号   ↓↓
'bug号：0000044
Private Sub cgrdInput_ItemKeyPress(ByVal paraItem As String, KeyAscii As Integer)
    Dim a As Integer
    Dim b As Integer
    a = 1
    b = 1
    If KeyAscii >= 48 And KeyAscii <= 57 Or ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123)) Then
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

Private Sub cgrdInput_RowChange(ByVal paraRow As Long, paraNewValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim lstrEnum As String
    
    On Error GoTo errHandler
    '创建体检项目对象。
'    Set lobjItem = CreateObject("职业病对象.clsTestItem")
'    'MsgBox "german", , "german"
'    '根据当前录入行信息设置lobjItem的属性。
'    With lobjItem
'        .编码 = cgrdInput.Value(paraRow, "编码")
'        .名称 = paraNewValue("名称")("值")
'
'        lstrEnum = IIf(IsNull(paraNewValue("枚举来源")("值")), "", paraNewValue("枚举来源")("值"))
'
'        '把中文逗号换成英文逗号。
'        If lstrEnum <> "" Then
'            lstrEnum = gffuncStrReplace(lstrEnum, "，", ",")
'        End If
'
'        .枚举来源 = lstrEnum
'        .缺省值 = IIf(IsNull(paraNewValue("缺省值")("值")), "", paraNewValue("缺省值")("值"))
'        .属性 = paraNewValue("属性")("值")
'        .体检大类 = Right(ctrwType.SelectedItem.Key, Len(ctrwType.SelectedItem.Key) - 1)
'
'        '检查缺省值是否在枚举来源中。
'        If .缺省值 <> "" And lstrEnum <> "" Then
'            If Right(lstrEnum, 1) <> "," Then lstrEnum = lstrEnum & ","
'            If InStr(1, lstrEnum, .缺省值 & ",") = 0 Then
'                Err.Raise 6666, , "缺省值必须在枚举来源中。"
'            End If
'        End If
'        .比较方式 = IIf(IsNull(paraNewValue("比较方式")("值")), "", paraNewValue("比较方式")("值"))
'        .标准值 = IIf(IsNull(paraNewValue("标准值")("值")), "", paraNewValue("标准值")("值"))
'        .单位 = IIf(IsNull(paraNewValue("单位")("值")), "", paraNewValue("单位")("值"))
'
'    End With
'
'    '保存当前行信息到数据库。
'    lobjItem.subSave
'
'    Set lobjItem = Nothing
    Exit Sub
errHandler:
    ErrorInfo = func错误处理(Err.Number, Err.Description)
    Cancel = True
End Sub

'作者：翁乔
'功能：添加彻底删除功能
'时间：2012-05-29
Private Sub cgrdInput_Delete(ByVal paraRow As Long, Cancel As Boolean, ErrorInfo As String)
    Dim lobj结论模板 As Object
    Dim flag As Boolean
    
    On Error GoTo errHandler
    
    '创建体检项目对象。
    Set lobj结论模板 = CreateObject("职业病对象.clsConclusionSet")
    flag = lobj结论模板.func删除最终结论模板(cgrdInput.Value(paraRow, "结论模板"), cgrdInput.Value(paraRow, "结论标准"))
    
    If flag = False Then
         MsgBox "删除失败，请退出窗口重新进入！", 16, "消息"
            cgrdInput.subRefresh
            Exit Sub
    End If
    
    'cgrdInput.Value(paraRow, "编码")
    Exit Sub
errHandler:
    ErrorInfo = func错误处理(Err.Number, Err.Description)
    Cancel = True
End Sub

Private Sub Form_Load()
    Dim lobj结论模板 As Object
    Dim lobjRec As Object
    
    On Error GoTo errHandler

    LblName.Caption = um用户名
    LblNo.Caption = um用户编号
    LblDate.Caption = Date
    
    
    '初始化录入网格:编码，名称，缺省值，枚举来源(可选值)，属性。
    'subAddItem,第3个参数true，背景色为浅黄
    cgrdInput.Enabled = True
    With cgrdInput.InputTemplate
        .subAddItem "结论模板", 0, True, True, 1200
        .subAddItem "结论日期", 0, False, False, 100, , , , , LblDate.Caption
        .subAddItem "结论标准", 4, False, True, 50, , "合格,不合格", , , "合格"
    End With
    cgrdInput.subDraw
    
    Set lobj结论模板 = CreateObject("职业病对象.clsConclusionSet")
    Set lobjRec = lobj结论模板.func获取最终结论模板
    
    gfsubLoadDyGridFromRec cgrdInput, lobjRec '传入树形列表以及数据库对象
    cgrdInput.subExpand
    lobjRec.Close
Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetTestItem", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub
