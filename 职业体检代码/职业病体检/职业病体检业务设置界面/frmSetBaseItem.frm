VERSION 5.00
Object = "{DA03AE6C-F6DB-11D1-9E6E-0040053F8E31}#3.3#0"; "DYCOMMINPUT.OCX"
Begin VB.Form frmSetBaseItem 
   Caption         =   "设置体检附加基本项目"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12300
   Icon            =   "frmSetBaseItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12300
   Begin 道源通用录入控件.DyInputGrid cgrdInput 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   11456
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "      (2) 选中网格中的某行，按《Del》键可以删除一行"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   4590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "说明：(1) 所谓附加项目就是除“姓名、性别、年龄、单位名称”四项各类体检人员共有的           属性外的其他属性。"
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   7260
   End
End
Attribute VB_Name = "frmSetBaseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春

'创造者:何啸天
'german
'功能：批量删除数据
Private Sub delete_any_Click()
    Dim loop_a As Integer
    Dim lcolInfo As New Collection
    Dim flag As Boolean
    
    flag = False
    
    For loop_a = 0 To cgrdInput.Rows Step 1
            If (cgrdInput.Value(loop_a, "是否删除") = "1") Then
            'string_debug = cgrdInput.Value(loop_a, "删除标志") & " " & string_debug
                'MsgBox "okay", , "okay"
                'lobjItem.编码 = cgrdInput.Value(loop_a, "编码")
                'lobjItem.subDelete_any
                pobj业务对象.Sub设置体检附加项目 3, lcolInfo, cgrdInput.Value(loop_a, "附加项目")
                'MsgBox CStr(loop_a), , "[DEBUG消息]"
                '参数1 操作类型 参数2 项目数据
                flag = True
            ElseIf (cgrdInput.Value(loop_a, "是否删除") <> "") Then
                MsgBox "是否删除 字段中要么不填 要么为1，请您填写正确的数据以便程序识别", 16, "消息"
                Exit Sub
            End If
        Next
        
    If (flag = False) Then
        MsgBox "你没有要批量删除的数据，请选择", 16, "消息"
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
    Dim lobjRec As Object
    Dim lcolInfo As Collection
    
    On Error GoTo errHandler
    '初始化录入网格：附加项目,录入标题,数据类型,数据长度,枚举值
    cgrdInput.InputTemplate.subRemoveAllItem
    With cgrdInput.InputTemplate
        .subAddItem "附加项目", 0, True, True, 30
        .subAddItem "录入标题", 0, False, True, 30
        .subAddItem "数据类型", 0, True, True, 20, 0, "1 日期型,2 数字型,3 文本型" ' dyInputSingleselecttext
        .subAddItem "数据长度", 3, True, True, 4, 0, , 300, 1
        .subAddItem "枚举值", 0, False, True, 50
    End With
    cgrdInput.subDraw
    
    '调用“pobj业务对象.所有体检附加项目”。
    '消息：数据库连接已经修改
    Set lobjRec = pobj业务对象.所有体检附加项目  '体检管理业务对象clsManageMedicalExam
    
    '把获取的项目加入录入网格中。
    gfsubLoadDyGridFromRec cgrdInput, lobjRec
    
    cgrdInput.subExpand '自动调整，单元格中数据的大小，以适应最长的数据
    lobjRec.Close
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病设置界面", "frmSetBaseItem", "Form_Load", 6666, lstrError, False
End Sub

'修改者: 何啸天
'german
Private Sub cgrdInput_AddNew(paraValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lcolInfo As Collection
    Dim llng数据类型 As Long
    Dim loop_a As Integer
    
    On Error GoTo errHandler
    '获取新增的项目信息。
    Set lcolInfo = New Collection
    
    '///////////////////////////////////////////////////////////
    '判断用户输入的数据是否是汉子
    'german
    For loop_a = 1 To Len(paraValue("附加项目")("值")) Step 1
        If (Asc(Mid(paraValue("附加项目")("值"), loop_a, 1)) > 0) Then
            MsgBox "[附加项目]:你输入的数据必须是汉字，数据记录动作已经被拦截", 16, "消息"
            Exit Sub
        End If
    Next
    
    For loop_a = 1 To Len(paraValue("枚举值")("值")) Step 1
        If (Asc(Mid(paraValue("枚举值")("值"), loop_a, 1)) >= 48 And Asc(Mid(paraValue("枚举值")("值"), loop_a, 1)) <= 48) Then
            MsgBox "[附加项目]:你输入的数据不能是数字，数据记录动作已经被拦截", 16, "消息"
            Exit Sub
        End If
    Next
    
    '///////////////////////////////////////////////////////////
    
    lcolInfo.Add paraValue("附加项目")("值"), "项目名称"
    lcolInfo.Add IIf(IsNull(paraValue("录入标题")("值")), "", paraValue("录入标题")("值")), "录入标题"
    
    llng数据类型 = Left(paraValue("数据类型")("值"), 1)
    
    lcolInfo.Add llng数据类型, "数据类型"
    lcolInfo.Add paraValue("数据长度")("值"), "数据长度"
    lcolInfo.Add IIf(IsNull(paraValue("枚举值")("值")), "", paraValue("枚举值")("值")), "枚举值"
    
    '新增到数据库中。
    pobj业务对象.Sub设置体检附加项目 1, lcolInfo  '体检管理业务对象clsManageMedicalExam
    
    Exit Sub
    
errHandler:
    ErrorInfo = func错误处理(Err.Number, Err.Description)
    Cancel = True
End Sub


Private Sub cgrdInput_RowChange(ByVal paraRow As Long, paraNewValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lcolInfo As Collection
    Dim llng数据类型 As Long
    
    On Error GoTo errHandler
    '获取新增的项目信息。
    Set lcolInfo = New Collection
    lcolInfo.Add paraNewValue("附加项目")("值"), "项目名称"
    lcolInfo.Add IIf(IsNull(paraNewValue("录入标题")("值")), "", paraNewValue("录入标题")("值")), "录入标题"
    llng数据类型 = Left(paraNewValue("数据类型")("值"), 1)
    lcolInfo.Add llng数据类型, "数据类型"
    lcolInfo.Add paraNewValue("数据长度")("值"), "数据长度"
    lcolInfo.Add IIf(IsNull(paraNewValue("枚举值")("值")), "", paraNewValue("枚举值")("值")), "枚举值"
    
    '新增到数据库中。
    pobj业务对象.Sub设置体检附加项目 2, lcolInfo, cgrdInput.Value(paraRow, "附加项目")

    Exit Sub
errHandler:
    ErrorInfo = func错误处理(Err.Number, Err.Description)
    Cancel = True
    cgrdInput.ItemSetfocus "附加项目"
    Exit Sub
    Resume
End Sub

Private Sub cgrdInput_Delete(ByVal paraRow As Long, Cancel As Boolean, ErrorInfo As String)
    Dim lcolInfo As New Collection
    
    On Error GoTo errHandler
    
    '从数据库中删除附加项目。
    'MsgBox CStr(paraRow), , "[DEBUG消息]"
    pobj业务对象.Sub设置体检附加项目 3, lcolInfo, cgrdInput.Value(paraRow, "附加项目")

    Exit Sub
errHandler:
    ErrorInfo = func错误处理(Err.Number, Err.Description)
    Cancel = True
End Sub

Private Sub ccmdExit_Click()
    Unload Me
End Sub
