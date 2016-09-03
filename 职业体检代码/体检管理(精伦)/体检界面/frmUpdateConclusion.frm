VERSION 5.00
Begin VB.Form frmUpdateConclusion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "体检结论设置"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "修改诊断处理意见"
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   8055
      Begin VB.CommandButton ccmdClear 
         Caption         =   "清空(&R)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1200
         Width           =   1020
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox cchkTemplate 
         Caption         =   "建议复查"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox ctxtDiagnosis 
         Height          =   1935
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "<<"
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
         Height          =   375
         Index           =   1
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   1020
      End
      Begin VB.ListBox clstAllDiagnosis 
         Height          =   1140
         ItemData        =   "frmUpdateConclusion.frx":0000
         Left            =   4440
         List            =   "frmUpdateConclusion.frx":0007
         TabIndex        =   8
         Top             =   480
         Width           =   3345
      End
      Begin VB.Label clblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请先进入“体检表设置”操作，选中体检表的“是否复查体检表”属性！"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2160
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   5760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊断处理意见："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所有可选意见："
         Height          =   180
         Index           =   1
         Left            =   4665
         TabIndex        =   10
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "修改体检结论"
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   8055
      Begin VB.ListBox clstSelectedConclusion 
         Height          =   1320
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   3015
      End
      Begin VB.ListBox clstAllConclusion 
         Height          =   1320
         Left            =   4440
         TabIndex        =   13
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "<<"
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
         Height          =   375
         Index           =   0
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   765
         Width           =   1020
      End
      Begin VB.CommandButton ccmdDel 
         Caption         =   ">>"
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
         Height          =   375
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检结论："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所有可选体检结论："
         Height          =   180
         Index           =   0
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "frmUpdateConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春
'最后修改：杨春

Private mstr系统编号 As String
Private mstr诊断意见 As String
Private mstr体检结论 As String
Private mstr复查体检表名 As String

Private Sub clstAllConclusion_DblClick()
    On Error Resume Next
    If clstAllConclusion.ListIndex >= 0 Then
        ccmdAdd_Click 0
    End If
    
End Sub

Private Sub clstAllDiagnosis_Click()
    On Error Resume Next
    If clstAllDiagnosis.ListIndex >= 0 Then
        ccmdAdd(1).Enabled = True
    End If
End Sub

Private Sub clstAllDiagnosis_DblClick()
    On Error Resume Next
    If clstAllDiagnosis.ListIndex >= 0 Then
        ccmdAdd_Click 1
    End If
    
End Sub

Private Sub clstSelectedConclusion_DblClick()
    On Error Resume Next
    If clstSelectedConclusion.ListIndex >= 0 Then
        ccmdDel_Click
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    Dim lobj体检表模板集 As Object 'clsMedicalExamTemplateSet。
    Dim lcolInfo As Collection     '获取的复查体检表名称。
    Dim i As Long
    On Error GoTo errHandler
    '创建体检表模板集对象。
    Set lobj体检表模板集 = CreateObject("体检对象.clsMedicalExamTemplateSet")
    
    '获取所有复查体检表名。
    lobj体检表模板集.体检表类型 = 2
    Set lcolInfo = lobj体检表模板集.元素集
    
    '显示在复查体检表下拉列表框中。
    ccmbTemplate.Clear
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
    Next
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmUpdateConclusion", "clstSelectedConclusion_Click", 6666, lstrError, False
End Sub

Public Property Get 系统编号() As String
    系统编号 = mstr系统编号
End Property

Public Property Let 系统编号(ByVal vNewValue As String)
    Dim lobj体检 As Object         'clsMedicalExam。
    Dim lobj体检表模板 As Object   'clsMedicalExamTemplate。
    Dim lcolInfo As Collection     '体检表模板的体检结论集属性。
    Dim lstrAllDiagnosis As String '体检表模板对象的属性“诊断处理意见”。
    Dim lstrItem As String
    Dim i As Long
    
    On Error GoTo errHandler
    If mstr系统编号 = vNewValue Then Exit Property
    
    '创建体检对象。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    lobj体检.系统编号 = vNewValue
    
    '创建体检表模板对象。
    Set lobj体检表模板 = CreateObject("体检对象.clsMedicalExamTemplate")
    lobj体检表模板.体检表名 = lobj体检.体检表.体检表名
    
    '把体检表模板的体检结论集中所有元素显示在clstAllConclusion中。
    Set lcolInfo = lobj体检表模板.体检结论集
    clstAllConclusion.Clear
    For i = 1 To lcolInfo.Count
        clstAllConclusion.AddItem lcolInfo(i)("名称")
    Next
    
    '把体检表模板对象的属性“诊断处理意见”中拆分后显示在clstAllDiagnosis中。
    lstrAllDiagnosis = lobj体检表模板.诊断处理意见
    clstAllDiagnosis.Clear
    i = 1
    lstrItem = gffuncGetItemFromList(lstrAllDiagnosis, i, ",")
    Do While lstrItem <> ""
        clstAllDiagnosis.AddItem lstrItem
        i = i + 1
        lstrItem = gffuncGetItemFromList(lstrAllDiagnosis, i, ",")
    Loop
    
    mstr系统编号 = vNewValue

    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmUpdateConclusion", "Property Let 系统编号", 6666, lstrError, True
End Property
Public Property Get 诊断处理意见() As String
    诊断处理意见 = mstr诊断意见
End Property

Public Property Let 诊断处理意见(ByVal vNewValue As String)
    On Error Resume Next
    mstr诊断意见 = vNewValue
    
    '把窗体属性"诊断处理意见"显示在“诊断处理意见”录入框中。
    ctxtDiagnosis.Text = mstr诊断意见
    
End Property

Public Property Get 体检结论() As String
    On Error Resume Next
    体检结论 = mstr体检结论
End Property
Public Property Let 体检结论(ByVal vNewValue As String)
    Dim lstrItem As String
    Dim i As Long
    
    On Error GoTo errHandler
    mstr体检结论 = vNewValue
    '把窗体属性"体检结论"拆分后显示在“选中体检结论”列表中。
    clstSelectedConclusion.Clear
    i = 1
    lstrItem = gffuncGetItemFromList(mstr体检结论, i, ",")
    Do While lstrItem <> ""
        clstSelectedConclusion.AddItem lstrItem
        i = i + 1
        lstrItem = gffuncGetItemFromList(mstr体检结论, i, ",")
    Loop
    
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmUpdateConclusion", "Property Let 体检结论", 6666, lstrError, True
End Property

Public Property Let 复查体检表名(ByVal vNewValue As String)
    On Error GoTo errHandler

    mstr复查体检表名 = vNewValue
    
    '若"mstr复查体检表名"为空，则不选中cchkTemplate，并且ccmbTemplate不可见；否则选中cchkTemplate，并且ccmbTemplate可见。
    If mstr复查体检表名 = "" Then
        cchkTemplate.Value = 0
        ccmbTemplate.Visible = False
    Else
        cchkTemplate.Value = 1
        ccmbTemplate.Visible = True
        
        '让复查体检表列表选中当前属性“复查体检表名”。
        ccmbTemplate.ListIndex = gffuncItemIsInComboBox(ccmbTemplate, mstr复查体检表名)
        
    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmUpdateConclusion", "Property Let 复查体检表名", 6666, lstrError, True
End Property

Public Property Get 复查体检表名() As String
    复查体检表名 = mstr复查体检表名
End Property

Private Sub cchkTemplate_Click()
    On Error Resume Next
    If cchkTemplate.Value = 1 Then
        ccmbTemplate.Visible = True
        If ccmbTemplate.ListCount = 0 Then
            clblInfo.Visible = True
        End If
        ccmbTemplate.SetFocus
    Else
        ccmbTemplate.Visible = False
    End If
End Sub

Private Sub ccmdAdd_Click(Index As Integer)
    On Error GoTo errHandler
    If Index = 0 Then
        '添加体检结论。
        clstSelectedConclusion.AddItem clstAllConclusion.Text
        clstAllConclusion.RemoveItem clstAllConclusion.ListIndex
        ccmdAdd(0).Enabled = False
    Else
        If ctxtDiagnosis.Text <> "" Then
            ctxtDiagnosis.Text = ctxtDiagnosis.Text & IIf(Right(Trim(ctxtDiagnosis.Text), 1) = ",", "", ",") & clstAllDiagnosis.Text
        Else
            ctxtDiagnosis.Text = clstAllDiagnosis.Text
        End If
        ccmdClear.Enabled = True
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmUpdateConclusion", "ccmdAdd_Click", 6666, lstrError, False
End Sub

Private Sub ccmdCancel_Click()
    '隐藏窗体。
    Me.Hide
End Sub

Private Sub ccmdClear_Click()
    ctxtDiagnosis.Text = ""
    
End Sub

Private Sub ccmdDel_Click()
    On Error GoTo errHandler
    '删除体检结论。
    clstAllConclusion.AddItem clstSelectedConclusion.Text
    clstSelectedConclusion.RemoveItem clstSelectedConclusion.ListIndex
    ccmdDel.Enabled = False
   
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmUpdateConclusion", "ccmdDel_Click", 6666, lstrError, False
End Sub

'功能：确定（根据录入结果设置窗体属性），并返回。
Private Sub ccmdOk_Click()
    Dim i As Long
    
    On Error GoTo errHandler
    '从选中的体检结论列表框中获取体检结论。
    mstr体检结论 = ""
    For i = 0 To clstSelectedConclusion.ListCount - 1
        mstr体检结论 = mstr体检结论 & clstSelectedConclusion.List(i) & ","
    Next
    If mstr体检结论 <> "" Then mstr体检结论 = Left(mstr体检结论, Len(mstr体检结论) - 1)
    
    mstr诊断意见 = Trim(ctxtDiagnosis.Text)
    
    If cchkTemplate.Value = 1 Then
        mstr复查体检表名 = ccmbTemplate.Text
    Else
        mstr复查体检表名 = ""
    End If
    
    '隐藏窗体。
    Me.Hide
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmUpdateConclusion", "ccmdOK_Click", 6666, lstrError, False
End Sub

Private Sub clstAllConclusion_Click()
    On Error Resume Next
    If clstAllConclusion.ListIndex >= 0 Then
        ccmdAdd(0).Enabled = True
    Else
        ccmdAdd(0).Enabled = False
    End If
End Sub

Private Sub clstSelectedConclusion_Click()
    On Error Resume Next
    ccmdDel.Enabled = True
End Sub

Private Sub ccmbTemplate_GotFocus()
    On Error Resume Next
    If ccmbTemplate.Text = "" And ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.ListIndex = 0
    End If
End Sub


