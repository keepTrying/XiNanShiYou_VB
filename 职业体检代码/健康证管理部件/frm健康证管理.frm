VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm健康证管理 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "食品"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11715
   ClipControls    =   0   'False
   Icon            =   "frm健康证管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton copt其它 
      Caption         =   "其它"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton copt食品 
      Caption         =   "食品"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox ctxt系统编号 
      Height          =   270
      Left            =   8640
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox ctxtNum 
      Height          =   270
      Left            =   6120
      TabIndex        =   6
      Text            =   "10"
      Top             =   840
      Width           =   495
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00FFDFFE&
      Caption         =   "调离"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00D1F7FE&
      Caption         =   "已打印"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00D2FCCF&
      Caption         =   "未打印"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10815
      _cx             =   25315972
      _cy             =   25307929
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
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "编号   |姓名    |性别    |年龄    |单位名称     |种类    |职业    |检出病种   | 体检结论 |健康证号"
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   -1  'True
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
   Begin MSComctlLib.Toolbar C工具栏 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   1200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "条码号(体检号/姓名)："
      Height          =   180
      Index           =   1
      Left            =   6720
      TabIndex        =   10
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "人"
      Height          =   180
      Left            =   2160
      TabIndex        =   7
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选中前面："
      Height          =   180
      Index           =   0
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "人数："
      Height          =   180
      Left            =   10800
      TabIndex        =   4
      Top             =   840
      Width           =   540
   End
   Begin VB.Menu cmnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu cmnuItemView 
         Caption         =   "查询(&Q)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "刷新(&R)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "退出(&Esc)"
         Index           =   4
      End
   End
   Begin VB.Menu cmnuInput 
      Caption         =   "录入(&I)"
      Visible         =   0   'False
      Begin VB.Menu cmnuItemInput 
         Caption         =   "新增(&N)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemInput 
         Caption         =   "修改(&E)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemInput 
         Caption         =   "删除(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu cmnuPrint 
      Caption         =   "打印(&p)"
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "健康证(&Z)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "调离通知"
         Index           =   2
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frm健康证管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象 '界面上引用的界面通用对
Attribute mobjGUI.VB_VarHelpID = -1

'查询条件。
Private mstr系统编号 As String
Private mstr姓名 As String
Private mstr单位 As String
Private mstr体检日期从 As String
Private mstr体检日期到 As String
Private mstr种类 As String
Private mstr发证单位 As String

Private mobjRec As Object

Private mcolIndex As Collection

Private Sub cchkType_Click(Index As Integer)
    subRefresh
End Sub

Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub cmnuItemInput_Click(Index As Integer)

    Dim lobj体检 As cls体检
    
    On Error GoTo errhandler
    Select Case Index
    Case 1 '新增
        frm体检录入.pstr系统编号 = ""
        frm体检录入.Show 1, Me
'        frm体检录入.Move Me.Left, Me.Top
        '刷新界面。
        subRefresh
    
    Case 2 '修改
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要修改的体检人员！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        frm体检录入.pstr系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, 0)
        frm体检录入.Show 1, Me
        
        '刷新界面。
        subRefresh
    
    Case 3 '删除
        Set lobj体检 = New cls体检
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要删除的体检人员！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        If MsgBox("你确信要删除“" & cgrdMain.TextMatrix(cgrdMain.Row, 1) & "”的体检记录吗？", vbYesNo + vbQuestion, "系统询问") = vbNo Then
            Exit Sub
        End If
        lobj体检.系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, 0)
        lobj体检.sub删除
        cgrdMain.RemoveItem cgrdMain.Row
    
    End Select
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm健康证管理", "cmnuItemInput_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemPrint_Click(Index As Integer)
    Dim i As Long
    Dim lobj体检 As cls体检
    Dim lstrSet As String
    Dim lobjRec As Object
    Dim lobjCN As Object
    On Error GoTo errhandler
    
    Select Case Index
    Case 1 '健康证
        Dim lcolInfo As Collection
        Dim lstrCN As String
        Dim lbln重新生成 As Boolean
        Dim lstr编号前缀 As String
        Dim lstr服务器编号 As String
        '根据业务设置，判断是否需要自动生成健康证号。
        '说明：省内限制必须使用带条码的健康证。
'        lstrSet = pobj体检管理.业务设置("健康证带条码")
'        lstrSet = "是"
'        If lstrSet = "是" Or pobj体检管理.业务设置("手工输入健康证号") = "是" Then
        
            '用户输入健康证号的起始号。
'            lstrCN = InputBox("请输入健康证的起始号", "输入")
'            If lstrCN = "" Then
'                Exit Sub
'            End If
            
            '判断输入健康证号是否为数字。
'            Do While Not (IsNumeric(lstrCN))
'                If MsgBox("你输入的健康证号格式不对。是否重新输入？", vbYesNo, "系统提示") = vbYes Then
'                    lstrCN = InputBox("请输入健康证的起始号", "输入")
'                Else
'                    Exit Sub
'                End If
'            Loop
            '检查卡号长度，12位为新版卡号，否则为老版卡号，要判断允许使用的最后期限
'            Dim lobjCheck As New clsCheck
'
'            If Len(lstrCN) <> 12 Then
''                If lobjCheck.funcCheckExpireDate() Then
''                    Err.Raise 6666, , "当前系统不能识别这种卡！"
''                End If
'                '未过期，判断卡是否合法。
'                Dim lobjEncrypt As Object
'                Set lobjEncrypt = CreateObject("fycarddes.clsDataEncrypt")
'                If Not lobjEncrypt.funcCheckJkzCardno(lstrCN) Then
'                    Err.Raise 6666, , "系统无法识别这张卡，请确定卡是否已损坏！"
'                End If
'            Else
'                '新版卡号，判断卡是否合法。
'                If Not funcCheckCardno(lstrCN) Then
'                    Err.Raise 6666, , "系统无法识别这张卡，请确定卡符合指定的格式！"
'                End If
'                '判断卡号是否超出公司的打印范围
'                If Not lobjCheck.funcCheckMaxNo(lstrCN) Then
'                    Err.Raise 6666, , "系统无法识别这张卡，该卡不是指定供应商发行的卡！"
'                End If
'            End If
'            lstrCN = Left(lstrCN, Len(lstrCN) - 2)  '去掉校验位
'        Else
'            '系统自动生成健康证号
'            lstrCN = ""
'        End If
        
        lstrCN = ""
        '获取选中的系统编号，创建体检对象。
        Set lcolInfo = New Collection
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.Cell(flexcpChecked, i, 1) = flexChecked Then
                Set lobj体检 = New cls体检
                lobj体检.系统编号 = cgrdMain.TextMatrix(i, 0)
                lobj体检.民族 = cgrdMain.TextMatrix(i, mcolIndex("血号"))
                If lobj体检.处置 = "调离" Then
                    Err.Raise 6666, , "调离人员不能打印健康证，请不要选中调离人员！"
                End If
                Set lobjCN = dafuncGetData("EXEC 健康证管理_生成健康证编号")
                lstrCN = lobjCN(0)
                Set lobjCN = Nothing
                Set lobjCN = dafuncGetData("SELECT 服务器代号 FROM 系统管理_系统基本配置表")
                lstr服务器编号 = lobjCN(0)
                
                Select Case lstr服务器编号
                    Case "1"
                        lstr编号前缀 = "A"
                    Case "2"
                        lstr编号前缀 = "B"
                    Case "3"
                        lstr编号前缀 = "C"
                    Case Else
                        lstr编号前缀 = "D"
                End Select
                
                lobj体检.健康证号 = lstr编号前缀 & Right(lstrCN, 7) '条码号

'                lobj体检.健康证号 = lobj体检.体检系统编号
                
                If lobj体检.种类 = "食品卫生" And lobj体检.身份证号 = "" Then   '食品证的证号存放在“身份证号”中，按年重编
                    lobj体检.身份证号 = Mid(pobj体检管理.func生成健康证号(), 2)     '系统生成的号是6位，只要后5位
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
                
                
                '健康证号自动递增。
'                lstrCN = Format(Val(lstrCN) + 1, String(Len(lstrCN), "0"))
                
            End If
        Next
        
        If lcolInfo.Count = 0 Then
            MsgBox "请选择要打印的体检人员（姓名上打勾）！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
'        pobj体检管理.sub打印健康证 lcolInfo
        Dim frm As Form
        Set frm = frmPrintPVCCard
        Set frm.Cards = lcolInfo
        frm.Show 1
        Unload frm
        '刷新界面。
        subRefresh
        
    Case 2 '调离通知。
        'frm调离管理.Show 1, Me
        
    End Select

    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm健康证管理", "cmnuItemPrint_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errhandler
    
    Select Case Index
    Case 1 '查询
        frm查询.Show 1, Me
        
        If frm查询.pblnOk Then
            mstr姓名 = frm查询.pstrName
            mstr系统编号 = frm查询.pstrNo
            mstr体检日期从 = frm查询.pstrStartDate
            mstr体检日期到 = frm查询.pstrEndDate
            mstr单位 = frm查询.pstrUnit
            'mstr种类 = frm查询.pstrType
            mstr种类 = IIf(copt食品.Value, "食品卫生", "")
            mstr发证单位 = frm查询.pstr发证单位
            
            subRefresh
        End If

    Case 2 '刷新。
        subRefresh
        
    Case 4
        Unload frm体检录入
        Unload Me
    End Select

    Exit Sub
errhandler:
    sfsub错误处理 "健康证界面部件", "frm健康证管理", "cmnuItemView_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Sub copt其它_Click()
    mstr种类 = ""
    subRefresh
End Sub

Private Sub copt食品_Click()
    mstr种类 = "食品卫生"
    subRefresh
End Sub

Private Sub ctxtNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim llngNum As Long
    On Error GoTo errhandler
    llngNum = Val(ctxtNum.Text)
    If llngNum > cgrdMain.Rows - 1 Then
        llngNum = cgrdMain.Rows - 1
    End If
    For i = 1 To cgrdMain.Rows - 1
        cgrdMain.Cell(flexcpChecked, i, 1) = flexUnchecked
    Next
    For i = 1 To llngNum
        cgrdMain.Cell(flexcpChecked, i, 1) = flexChecked
    Next
    Exit Sub
errhandler:
End Sub

Private Sub ctxt系统编号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    Dim lobjRec As Object
    On Error GoTo errhandler
    If KeyCode = 13 And ctxt系统编号 <> "" Then
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.TextMatrix(i, mcolIndex("体检系统编号")) = ctxt系统编号 Then
                cgrdMain.TopRow = i
                cgrdMain.Row = i
                Exit Sub
            End If
        Next
        
        '现在网格里没有找到这个编号。从库里找。
        Set lobjRec = pobj体检管理.func健康体检查询("", "", "", "", "", "", "", "", ctxt系统编号)
        If lobjRec.RecordCount > 0 Then
            cgrdMain.Rows = cgrdMain.Rows + 1
            i = cgrdMain.Rows - 1
            cgrdMain.Cell(flexcpChecked, i, 1) = flexUnchecked
            For j = 0 To cgrdMain.Cols - 1
                cgrdMain.TextMatrix(i, j) = IIf(IsNull(lobjRec(j)), "", lobjRec(j))
            Next
            cgrdMain.AutoSize 0, cgrdMain.Cols - 1
            
            '显示颜色。
            If lobjRec!处置 = "发健康证" Then
                If lobjRec!状态 = "未打印" Then
                    cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(0).BackColor
                Else
                    cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(1).BackColor
                End If
            Else
                cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(2).BackColor
            End If
            cgrdMain.ColWidth(1) = 1000
            
            '隐藏系统编号。
            cgrdMain.ColHidden(0) = True
            
            cgrdMain.TopRow = i
            cgrdMain.Row = i
            clblInfo.Caption = "人数：" & cgrdMain.Rows - 1
        End If
        
        ctxt系统编号.Text = ""
    End If
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm健康证管理", "ctxt系统编号_KeyDown", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Load()

    On Error GoTo errhandler

    If pblnInUse Then Exit Sub
    pblnInUse = True
    
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.C工具栏 = C工具栏
    lcol工具栏按钮.Add "刷新"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "添加"
    lcol工具栏按钮.Add "修改"
    lcol工具栏按钮.Add "删除"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "打印证(&Z)107"
    lcol工具栏按钮.Add "调离通知(&T)107"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    mobjGUI.subInitialize lcol工具栏按钮, ""
    '不允许使用本子系统录证
    C工具栏.Buttons(3).Visible = False
    C工具栏.Buttons(4).Visible = False
    C工具栏.Buttons(5).Visible = False
    C工具栏.Buttons(6).Visible = False
    
    '权限判断。
'    If Not umfunc校验用户权限("健康证管理_录入") Then
'        C工具栏.Buttons(3).Visible = False
'        C工具栏.Buttons(4).Visible = False
'        If Not umfunc校验用户权限("健康证管理_删除") Then
'            C工具栏.Buttons(5).Visible = False
'            C工具栏.Buttons(6).Visible = False
'            cmnuInput.Visible = False
'        Else
'            cmnuItemInput(1).Visible = False
'            cmnuItemInput(2).Visible = False
'        End If
'    Else
'        If Not umfunc校验用户权限("健康证管理_删除") Then
'            C工具栏.Buttons(5).Visible = False
'            cmnuItemInput(3).Visible = False
'        End If
'    End If
    If Not umfunc校验用户权限("健康证管理_打印") Then
        C工具栏.Buttons(7).Visible = False
        C工具栏.Buttons(8).Visible = False
        C工具栏.Buttons(9).Visible = False
        cmnuPrint.Visible = False
    End If
    
    C工具栏.Buttons(8).Visible = False
    
    '获取最近两周的未打印体检记录。
    mstr体检日期从 = Format(DateAdd("d", 1 - DatePart("w", Now, vbMonday), Now) - 7, "yyyy-mm-dd")
    mstr体检日期到 = Format(Now, "yyyy-mm-dd")
    mstr种类 = "食品卫生"
    
    subRefresh
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm健康证管理", "Form_Load", Err.Number, Err.Description, False
End Sub

'功能：根据查询条件显示查询结果。
Public Sub subRefresh()
    
    Dim lstr状态条件 As String
    Dim i As Long
    lstr状态条件 = ""
    
    '首先导入数据。
    On Error Resume Next
    dafuncGetData "exec 健康证管理_导入已体检完毕人员信息"
    
    On Error GoTo errhandler
    If cchkType(0).Value = 1 Or cchkType(1).Value = 1 Then
        lstr状态条件 = "(处置='发健康证'"
        If cchkType(0).Value = 1 And cchkType(1).Value = 0 Then
            lstr状态条件 = lstr状态条件 & " and 状态='未打印'"
        ElseIf cchkType(0).Value = 0 And cchkType(1).Value = 1 Then
            lstr状态条件 = lstr状态条件 & " and 状态='已打印'"
        End If
        lstr状态条件 = lstr状态条件 & ")"
    End If
    If cchkType(2).Value = 1 Then
        lstr状态条件 = lstr状态条件 & IIf(lstr状态条件 = "", "", " or ") & "处置='调离'"
    End If
    If lstr状态条件 = "" Then lstr状态条件 = "1=0"
    
    If mstr体检日期从 = "" Then mstr体检日期从 = DateAdd("d", -30, Date)
    
    Set mobjRec = pobj体检管理.func健康体检查询(mstr系统编号, mstr姓名, mstr单位, mstr体检日期从, mstr体检日期到, mstr种类, lstr状态条件, mstr发证单位)
    
    cgrdMain.FormatString = ""
    Set cgrdMain.DataSource = mobjRec
    For i = 1 To cgrdMain.Rows - 1
        cgrdMain.Cell(flexcpChecked, i, 1) = flexUnchecked
        '显示颜色。
        If mobjRec!处置 = "发健康证" Then
            If mobjRec!状态 = "未打印" Then
                cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(0).BackColor
                
            Else
                cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(1).BackColor
            End If
        Else
            cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(2).BackColor
        End If
        mobjRec.MoveNext
    Next
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.Cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    
    cgrdMain.ColWidth(1) = 1000
    
    '隐藏系统编号。
    cgrdMain.ColHidden(0) = True
    
    clblInfo.Caption = "人数：" & cgrdMain.Rows - 1

    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm健康证管理", "subRefresh", Err.Number, Err.Description, True
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjGUI = Nothing
    Set mobjRec = Nothing
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    
    Select Case Operate
    Case "刷新"
        cmnuItemView_Click 2
    Case "添加"
        cmnuItemInput_Click 1
    Case "修改"
        cmnuItemInput_Click 2
    Case "删除"
        cmnuItemInput_Click 3
    Case "打印证"
        cmnuItemPrint_Click 1
    Case "调离通知"
        cmnuItemPrint_Click 2
    End Select
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm健康证管理", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    
End Sub
Function funcCheckCardno(paraCardno As String) As Boolean
    Dim i As Long
    
    i = CLng(Left(paraCardno, 10))
    If Format(((i Mod 99) * 3) Mod 75, "00") = Right(paraCardno, 2) Then
        funcCheckCardno = True
    Else
        funcCheckCardno = False
    End If
End Function
