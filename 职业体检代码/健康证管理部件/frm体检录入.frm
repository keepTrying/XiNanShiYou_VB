VERSION 5.00
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm体检录入 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "健康证录入"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm体检录入.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox ccmb行业类别 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton ccmdClear 
      Caption         =   "清空(&C)"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "打印证(&P)"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "保存后自动自动清空"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   7320
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox ctxt体检号 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton ccmd定位 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "单位定位"
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox ccmbMZ 
      Height          =   300
      Left            =   1200
      TabIndex        =   38
      Top             =   2880
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   4800
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox ctxt体检日期 
      Height          =   300
      Left            =   6600
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton ccmdSaveAs 
      Caption         =   "另存照片(&A)"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdLoad 
      Caption         =   "载入照片(&L)"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox ccmb处置 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "frm体检录入.frx":0E42
      Left            =   1200
      List            =   "frm体检录入.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6000
      Width           =   3255
   End
   Begin VB.ComboBox ccmb发证单位 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frm体检录入.frx":0E60
      Left            =   1200
      List            =   "frm体检录入.frx":0E6A
      TabIndex        =   12
      Top             =   6480
      Width           =   3255
   End
   Begin VB.ComboBox ccmb培训结论 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "frm体检录入.frx":0E7C
      Left            =   1200
      List            =   "frm体检录入.frx":0E86
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "照像"
      Height          =   4215
      Left            =   5520
      TabIndex        =   28
      Top             =   240
      Width           =   4935
      Begin dyCatchPhoto.ctlCatchPhoto ctlCatchPhoto 
         Height          =   3615
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6376
         BackColor       =   0
         FontSize        =   11.25
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   27
      Top             =   6840
      Width           =   10815
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "返回(&C)"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton ccmdSave 
      Caption         =   "保存(&S)"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox ccmb体检结论 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "frm体检录入.frx":0E98
      Left            =   1200
      List            =   "frm体检录入.frx":0EA2
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5040
      Width           =   3255
   End
   Begin VB.ListBox clstDisease 
      Height          =   1530
      ItemData        =   "frm体检录入.frx":0EB4
      Left            =   1200
      List            =   "frm体检录入.frx":0EC1
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
   End
   Begin VB.ComboBox ccmbOcc 
      Height          =   300
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
   End
   Begin VB.ComboBox ccmbType 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox ctxtUnit 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox ctxtAge 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox ccmbSex 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox ctxtName 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin MSMask.MaskEdBox ctxt发证日期 
      Height          =   300
      Left            =   6600
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ctxt有效期至 
      Height          =   300
      Left            =   9240
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ctxt培训日期 
      Height          =   300
      Left            =   9240
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "培训日期："
      Height          =   180
      Index           =   16
      Left            =   8280
      TabIndex        =   42
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检号："
      Height          =   180
      Index           =   15
      Left            =   120
      TabIndex        =   41
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "民    族："
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   39
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "处    置："
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "有效期至："
      Height          =   180
      Index           =   12
      Left            =   8280
      TabIndex        =   34
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发证日期："
      Height          =   180
      Index           =   11
      Left            =   5640
      TabIndex        =   33
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检日期："
      Height          =   180
      Index           =   5
      Left            =   5640
      TabIndex        =   32
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检单位："
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   31
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "培训结论："
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   30
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检结论："
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检出病种："
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "职    业："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "种    类："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位名称："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年    龄："
      Height          =   180
      Index           =   2
      Left            =   2520
      TabIndex        =   21
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性    别："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓    名："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   900
   End
   Begin VB.Menu cmnuPop 
      Caption         =   "cmnuPop"
      Visible         =   0   'False
      Begin VB.Menu cmnuItemPop 
         Caption         =   "添加"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPop 
         Caption         =   "删除"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm体检录入"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstr系统编号 As String

Private Sub ccmbType_Click()
    Dim lobjRec As Object
    On Error GoTo errhandler
    
    ccmb行业类别.Clear
    Set lobjRec = dafuncGetData("select * from 系统管理_行业类别字典视图 where Parent=(select InnerID from 系统管理_卫生种类字典视图 where 名称='" & ccmbType.Text & "'" & IIf(ccmbType.Text = "公共卫生", " or 名称='公共场所卫生'", "") & ")")
    Do While Not lobjRec.EOF
        ccmb行业类别.AddItem lobjRec!名称
        lobjRec.movenext
    Loop
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm体检录入", "ccmbType_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ccmb处置_Click()
    On Error Resume Next
    If ccmb处置.ListIndex = 0 Then
        ctxt发证日期.Text = Format(Now, "yyyy-mm-dd")
        ctxt有效期至.Text = Format(DateAdd("y", 1, ctxt发证日期.Text), "yyyy-mm-dd")
    Else
        ctxt发证日期.Text = "____-__-__"
        ctxt有效期至.Text = "____-__-__"
    End If
End Sub

Private Sub ccmb体检结论_Click()
    On Error Resume Next
    If ccmb体检结论.Text = "合格" Then
        ccmb处置.ListIndex = 0
        ctxt发证日期.Text = Format(Now, "yyyy-mm-dd")
        ctxt有效期至.Text = Format(DateAdd("y", 1, ctxt发证日期.Text), "yyyy-mm-dd")
        
    Else
        ccmb处置.ListIndex = 1
        ctxt发证日期.Text = "____-__-__"
        ctxt有效期至.Text = "____-__-__"
    End If
    
End Sub

Private Sub ccmdClear_Click()
    ctxt体检号.Text = ""
    ctxtName.Text = ""
    pstr系统编号 = ""
    ctxtAge = ""
End Sub

Private Sub ccmdExit_Click()
    On Error Resume Next
    pobj记忆.sub覆盖记忆值 "健康证录入保存后清空", cchkClear.Value
    
    Unload Me
    Call frm健康证管理.subRefresh
End Sub

Private Sub ccmdLoad_Click()
    Dim lstrFile As String
    On Error GoTo errhandler
    
    ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片", vbDirectory) <> "" Then
        ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Photo"
    End If
    ccmdFile.FileName = pstr系统编号
    ccmdFile.ShowOpen
    lstrFile = ccmdFile.FileName
    If lstrFile <> "" Then
        If InStr(lstrFile, ".") > 0 Then
            Set ctlCatchPhoto.Photo = LoadPicture(lstrFile)
        End If
    End If
    Exit Sub
errhandler:
    MsgBox "载入照片失败！" & Error, vbOKOnly + vbExclamation, "系统提示"
End Sub

Private Sub ccmdPrint_Click()
    Dim lstrCN As String
    On Error GoTo errhandler
    Dim lstrCode As String
    
    '根据业务设置，判断是否需要自动生成健康证号。
    lstrCode = pobj体检管理.业务设置("健康证带条码")
    lstrCode = "是" '省内强制要求使用带条码的证。
    If lstrCode = "是" Or pobj体检管理.业务设置("手工输入健康证号") = "是" Then
    
        '用户输入健康证号的起始号。
        lstrCN = InputBox("请输入健康证号", "输入")
        If lstrCN = "" Then
            Exit Sub
        End If
        
        '判断输入健康证号是否为数字。
        If lstrCode = "是" Then
            Do While Not (IsNumeric(lstrCN))
                If MsgBox("你输入的健康证号格式不对。是否重新输入？", vbYesNo, "系统提示") = vbYes Then
                    lstrCN = InputBox("请输入健康证的起始号", "输入")
                Else
                    Exit Sub
                End If
            Loop
            
            '判断卡是否合法。
            Dim lobjEncrypt As Object
            Set lobjEncrypt = CreateObject("fycarddes.clsDataEncrypt")
            If Not lobjEncrypt.funcCheckJkzCardno(lstrCN) Then
                Err.Raise 6666, , "系统无法识别这张卡，请确定卡符合指定的格式或卡是否已损坏！"
            End If
            '不保存校验位。
            lstrCN = lobjEncrypt.卡号
            Set lobjEncrypt = Nothing
        End If
    Else
        '系统自动生成健康证号
        lstrCN = ""
    End If
    
    Dim lcolInfo As New Collection
    Dim lobj体检 As cls体检
    Set lobj体检 = New cls体检
    lobj体检.系统编号 = pstr系统编号
    
    If Not (lstrCode = "是" Or pobj体检管理.业务设置("手工输入健康证号") = "是") Then
        '需要系统自动生成许可证号。
        If lobj体检.状态 = "未打印" And lobj体检.健康证号 = "" Then
            lstrCN = pobj体检管理.func生成健康证号()
            lobj体检.健康证号 = lstrCN
        End If
    Else
        lobj体检.健康证号 = lstrCN
    End If
    
    '如果是体检系统流过来的记录，没有发证日期和发证单位。
    If lobj体检.发证日期 = "" Then
        lobj体检.发证日期 = Format(Date, "yyyy-mm-dd")
    End If
    If lobj体检.有效期至 = "" Then
        lobj体检.有效期至 = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Date)), "yyyy-mm-dd")
    End If
    If lobj体检.发证单位 = "" Then
        lobj体检.发证单位 = um防疫站名
    End If
                    
    lcolInfo.Add lobj体检
       
    pobj体检管理.sub打印健康证 lcolInfo
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "ccmdPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub Ccmdsave_Click()
    Dim lobj体检 As New cls体检
    Dim lstr检出病种 As String
    Dim i As Long
    On Error GoTo errhandler
    
    '检查输入合法性。
    If ctxtName.Text = "" Then
        MsgBox "必须输入姓名！其它黄色框也必须录入！", vbOKOnly + vbExclamation, "系统提示"
        ctxtName.SetFocus
        Exit Sub
    End If
    If Len(ccmbOcc.Text) > 10 Then
        MsgBox "职业最多输入10个汉字！", vbOKOnly + vbExclamation, "系统提示"
        ccmbOcc.SetFocus
        Exit Sub
    End If
    If Not IsDate(ctxt体检日期.Text) Then
        MsgBox "必须输入体检日期！其它黄色框也必须录入！", vbOKOnly + vbExclamation, "系统提示"
        ctxt体检日期.SetFocus
        Exit Sub
    End If
    If Not IsDate(ctxt培训日期.Text) Then
        MsgBox "必须输入培训日期！其它黄色框也必须录入！", vbOKOnly + vbExclamation, "系统提示"
        ctxt体检日期.SetFocus
        Exit Sub
    End If
    If ccmb处置.ListIndex = 0 Then
        If Not IsDate(ctxt发证日期.Text) Or Not IsDate(ctxt有效期至.Text) Then
            MsgBox "必须输入发证日期和有效期至！其它黄色框也必须录入！", vbOKOnly + vbExclamation, "系统提示"
            ccmb发证单位.SetFocus
            Exit Sub
        End If
        If Len(ccmb发证单位.Text) > 30 Then
            MsgBox "发证单位最多输入30个汉字！", vbOKOnly + vbExclamation, "系统提示"
            ccmbOcc.SetFocus
            Exit Sub
        End If
    End If
    
    '获取检出病重。
    lstr检出病种 = ""
    For i = 0 To clstDisease.ListCount - 1
        If clstDisease.Selected(i) Then
            lstr检出病种 = lstr检出病种 & clstDisease.List(i) & ","
        End If
    Next
    If lstr检出病种 <> "" Then lstr检出病种 = Left(lstr检出病种, Len(lstr检出病种) - 1)
    
    '保存。
    With lobj体检
        .系统编号 = pstr系统编号
        .体检号 = ctxt体检号.Text
        .姓名 = ctxtName.Text
        .性别 = ccmbSex.Text
        .年龄 = ctxtAge.Text
        .申请编号 = ctxtUnit.Tag
        .单位名称 = ctxtUnit.Text
        .种类 = ccmbType.Text
        .职业 = ccmbOcc.Text
        .民族 = ccmbMZ.Text
        .体检结论 = ccmb体检结论.Text
        .培训结论 = ccmb培训结论.Text
        .体检日期 = ctxt体检日期.FormattedText
        .培训日期 = ctxt培训日期.FormattedText
        .处置 = ccmb处置.Text
        
        If ctxt发证日期.FormattedText = "____-__-__" Then
            .发证日期 = ""
        Else
            .发证日期 = ctxt发证日期.FormattedText
        End If
        If ctxt有效期至.FormattedText = "____-__-__" Then
            .有效期至 = ""
        Else
            .有效期至 = ctxt有效期至.FormattedText
        End If
        .检出病种 = lstr检出病种
        .发证单位 = ccmb发证单位.Text
        
        If pobj体检管理.业务设置("是否照相") = "是" Then
            Set .照片 = ctlCatchPhoto.Photo
        End If
        
        .sub保存
    End With
    
    On Error Resume Next
    
    '保存枚举值。
    pobj记忆.sub添加记忆值 "卫生种类", Trim(ccmbType.Text)
    pobj记忆.sub添加记忆值 "职业", Trim(ccmbOcc.Text)
    pobj记忆.sub添加记忆值 "民族", Trim(ccmbMZ.Text)
    If Trim(ccmb发证单位.Text) <> "" Then
        pobj记忆.sub添加记忆值 "发证单位", Trim(ccmb发证单位.Text)
    End If
    
    
    If pstr系统编号 = "" And cchkClear.Value = 1 Then
        '清空必录项。
        ctxtName.Text = ""
        ctxtAge.Text = ""
        ctxt体检号.Text = ""
        ctxt体检号.SetFocus
        
        '恢复照相。
        If pobj体检管理.业务设置("是否照相") = "是" Then
            If ctlCatchPhoto.Status = "恢复" Then
                ctlCatchPhoto.sub转换状态
            End If
        End If
    Else
        pstr系统编号 = lobj体检.系统编号
        ccmdPrint.Enabled = True
    End If

    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm体检录入", "Ccmdsave_Click", Err.Number, Err.Description, False
End Sub

Private Sub ccmdSaveAs_Click()
    Dim lstrFile As String
    On Error GoTo errhandler
    
    ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "像片", vbDirectory) <> "" Then
        ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Photo"
    End If
    ccmdFile.FileName = pstr系统编号
    ccmdFile.ShowOpen
    lstrFile = ccmdFile.FileName
    If lstrFile <> "" Then
        SavePicture ctlCatchPhoto.Photo, lstrFile
    End If

    Exit Sub
errhandler:
    MsgBox "另存照片失败！" & Error, vbOKOnly + vbExclamation, "系统提示"
End Sub

Private Sub ccmd定位_Click()
    Dim lstrUnitName As String
    Dim lrstTmp As Object
    Dim lstrUnitNumber As String
    Dim lobj接口 As Object
    
    On Error GoTo errhandler
    
    Set lobj接口 = CreateObject("单位档案业务.ClsUnitInterface")
    Set lrstTmp = lobj接口.func单位简单定位(Screen.Width / 2, Screen.Height / 2)
    
    If lrstTmp Is Nothing Then
        ctxtUnit.SetFocus
        Exit Sub
    End If
    
    lstrUnitNumber = lrstTmp!申请编号
    ctxtUnit = lrstTmp!单位名称
    ccmbType.Text = lrstTmp!卫生种类
    ccmbType_Click
    
    ccmb行业类别.Text = IIf(IsNull(lrstTmp!行业类别), "", lrstTmp!行业类别)
    
    Me.ctxtUnit.Tag = lrstTmp!申请编号 '在单位名称的tag中记录单位申请编号。
    ctxtUnit.SetFocus
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm体检录入", "ccmd定位_Click", Err.Number, Err.Description, False
End Sub

Private Sub clstDisease_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbRightButton Then
        If clstDisease.ListIndex < 0 Then
            cmnuItemPop(2).Enabled = False
        Else
            cmnuItemPop(2).Enabled = True
        End If
        PopupMenu cmnuPop
    End If
End Sub

Private Sub cmnuItemPop_Click(Index As Integer)
    Dim lstrItem As String
    Dim i As Long
    
    On Error Resume Next
    Select Case Index
    Case 1 '新增
        '输入新增病种。
        lstrItem = Trim(InputBox("请输入新增病种：", "系统询问", ""))
        If lstrItem <> "" Then
            If InStr(lstrItem, "'") > 0 Then
                MsgBox "病种名称中不能含有非法字符“'”（即单引号）！", vbOKOnly + vbExclamation, "系统提示"
                Exit Sub
            End If
            i = 0
            For i = 0 To clstDisease.ListCount - 1
                If clstDisease.List(i) = lstrItem Then
                    Exit For
                End If
            Next
            If i = clstDisease.ListCount Then
                clstDisease.AddItem lstrItem
            End If
        End If
    Case 2 '删除。
        If clstDisease.ListIndex >= 0 Then
            If clstDisease.List(clstDisease.ListIndex) = "无" Then
                MsgBox "该项目不能删除！", vbOKOnly + vbExclamation, "系统提示"
                Exit Sub
            End If
            clstDisease.RemoveItem clstDisease.ListIndex
        End If
    End Select
    
End Sub

Private Sub ctxt发证日期_Change()
    On Error Resume Next
    If IsDate(ctxt发证日期.Text) Then
        ctxt有效期至.Text = Format(DateAdd("d", -1, DateAdd("yyyy", 1, ctxt发证日期.Text)), "yyyy-mm-dd")
    End If
End Sub

'功能：控制不能输入单印号，处理回车。
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        SendKeys Chr(9)
    ElseIf KeyCode = 39 Then
        KeyCode = 0
    End If
    

End Sub



Private Sub Form_Load()
    Dim lcolInfo As Collection
    Dim i As Long
    
    On Error GoTo errhandler
    
    ccmbSex.Clear
    ccmbSex.AddItem "男"
    ccmbSex.AddItem "女"
    ccmbSex.ListIndex = 0
    
    '获取种类。
    Set lcolInfo = pobj记忆.记忆项值("卫生种类", True)
    ccmbType.Clear
    For i = 1 To lcolInfo.Count
        ccmbType.AddItem lcolInfo(i)
    Next
    If ccmbType.ListCount > 0 Then ccmbType.ListIndex = 0
    
    '获取职业。
    Set lcolInfo = pobj记忆.记忆项值("职业", True)
    ccmbOcc.Clear
    For i = 1 To lcolInfo.Count
        ccmbOcc.AddItem lcolInfo(i)
    Next
    If ccmbOcc.ListCount > 0 Then ccmbOcc.ListIndex = 0
    
    
    '获取民族。
    Set lcolInfo = pobj记忆.记忆项值("民族", True)
    ccmbMZ.Clear
    For i = 1 To lcolInfo.Count
        ccmbMZ.AddItem lcolInfo(i)
    Next
    If ccmbMZ.ListCount > 0 Then ccmbMZ.ListIndex = 0
    
    '获取检出病种。
    Set lcolInfo = pobj记忆.记忆项值("检出病种", True)
    clstDisease.Clear
    clstDisease.AddItem "无从业禁忌症"
    For i = 1 To lcolInfo.Count
        clstDisease.AddItem lcolInfo(i)
    Next
    clstDisease.Selected(0) = True
    
    '获取发证单位。
    Set lcolInfo = pobj记忆.记忆项值("发证单位", True)
    ccmb发证单位.Clear
    ccmb发证单位.AddItem ""
    For i = 1 To lcolInfo.Count
        ccmb发证单位.AddItem lcolInfo(i)
    Next
    If ccmb发证单位.ListCount > 1 Then
        ccmb发证单位.ListIndex = 1
    Else
        ccmb发证单位.ListIndex = 0
    End If
    
    ccmb体检结论.ListIndex = 0
    ccmb培训结论.ListIndex = 0
    ccmb处置.ListIndex = 0
    
    ctxt体检日期.Text = Format(Date, "yyyy-mm-dd")
    ctxt发证日期.Text = Format(Date, "yyyy-mm-dd")
    ctxt培训日期.Text = Format(Date, "yyyy-mm-dd")
    ctxt有效期至.Text = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Date)), "yyyy-mm-dd")
    
    '显示修改人员信息。
    Dim lobj体检 As cls体检
    Dim lstr检出病种 As String
    If pstr系统编号 <> "" Then
        Set lobj体检 = New cls体检
        lobj体检.系统编号 = pstr系统编号
        
        ctxt体检号.Text = lobj体检.体检号
        ctxtName.Text = lobj体检.姓名
        If lobj体检.性别 = "男" Then
            ccmbSex.ListIndex = 0
        Else
            ccmbSex.ListIndex = 1
        End If
        ctxtAge.Text = lobj体检.年龄
        
        ccmbType.Text = lobj体检.种类
        ccmbOcc.Text = lobj体检.职业
        
        ctxtUnit.Text = lobj体检.单位名称
        ctxtUnit.Tag = lobj体检.申请编号
        
        lstr检出病种 = lobj体检.检出病种
        If lstr检出病种 <> "" Then
            lstr检出病种 = lstr检出病种 & ","
            For i = 0 To clstDisease.ListCount - 1
                If InStr(lstr检出病种, clstDisease.List(i) & ",") > 0 Then
                    clstDisease.Selected(i) = True
                End If
            Next
        End If
        
        If lobj体检.体检结论 = "合格" Then
            ccmb体检结论.ListIndex = 0
        Else
            ccmb体检结论.ListIndex = 1
        End If
    
        If lobj体检.培训结论 = "合格" Then
            ccmb培训结论.ListIndex = 0
        Else
            ccmb培训结论.ListIndex = 1
        End If
        If lobj体检.处置 = "发健康证" Then
            ccmb处置.ListIndex = 0
        Else
            ccmb处置.ListIndex = 1
        End If
        
        ctxt体检日期.Text = Format(lobj体检.体检日期, "yyyy-mm-dd")
        
        ccmb发证单位.Text = lobj体检.发证单位
        
        If lobj体检.发证日期 = "" Then
            ctxt发证日期.Text = "____-__-__"
        Else
            ctxt发证日期.Text = Format(lobj体检.发证日期, "yyyy-mm-dd")
        End If
        If lobj体检.有效期至 = "" Then
            ctxt有效期至.Text = "____-__-__"
        Else
            ctxt有效期至.Text = Format(lobj体检.有效期至, "yyyy-mm-dd")
        End If
        
        Set ctlCatchPhoto.Photo = lobj体检.照片
        ccmdPrint.Enabled = True
    Else
        ccmdPrint.Enabled = False
    End If
    
    '初始化照像控件。
    If pobj体检管理.业务设置("是否照相") = "是" Then
        Frame2.Caption = "照像"
        ctlCatchPhoto.Enabled = True
        ctlCatchPhoto.funcInitVideo
        ccmdLoad.Enabled = True
        ccmdSaveAs.Enabled = True
    Else
        Frame2.Caption = "业务设置为不照像"
        ctlCatchPhoto.Enabled = False
        ccmdLoad.Enabled = False
        ccmdSaveAs.Enabled = False
    End If
    If pobj记忆.记忆项值("健康证录入保存后清空") = "1" And pstr系统编号 = "" Then
        cchkClear.Value = 1
    Else
        cchkClear.Value = 0
    End If
    
    If Not umfunc校验用户权限("健康证管理_打印") Then
        ccmdPrint.Visible = False
    End If
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm体检录入", "Form_Load", Err.Number, Err.Description, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ctlCatchPhoto.subDisconnect

    On Error GoTo errhandler
    Dim lstrList As String
    Dim i As Long
    
    For i = 1 To clstDisease.ListCount - 1
        lstrList = lstrList & clstDisease.List(i) & ","
    Next
    
    If lstrList <> "" Then lstrList = Left(lstrList, Len(lstrList) - 1)
    pobj记忆.sub覆盖记忆值 "检出病种", lstrList
    
errhandler:

End Sub


