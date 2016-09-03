VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm健康证管理 
   Caption         =   "健康证管理"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm健康证打印.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.TextBox ctxtNum 
      Height          =   270
      Left            =   1080
      TabIndex        =   8
      Text            =   "10"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0C0FF&
      Caption         =   "调离"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0FFFF&
      Caption         =   "已打印"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0FFC0&
      Caption         =   "未打印"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1244
      ButtonWidth     =   2037
      ButtonHeight    =   1085
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询(&Q)"
            Key             =   "query"
            ImageKey        =   "query"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增(&N)"
            Key             =   "new"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改(&U)"
            Key             =   "update"
            ImageKey        =   "update"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除(&D)"
            Key             =   "delete"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "打印(&P)"
            Key             =   "print"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "调离通知(&D)"
            Key             =   "dl"
            ImageKey        =   "dl"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出(&E)"
            Key             =   "exit"
            ImageKey        =   "exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm健康证打印.frx":0E42
            Key             =   "query"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm健康证打印.frx":115C
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm健康证打印.frx":1476
            Key             =   "dl"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm健康证打印.frx":18C8
            Key             =   "new"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm健康证打印.frx":1D1A
            Key             =   "update"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm健康证打印.frx":216C
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm健康证打印.frx":2486
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   6975
      Left            =   -360
      TabIndex        =   1
      Top             =   1200
      Width           =   11655
      _cx             =   23810126
      _cy             =   23801871
      _ConvInfo       =   1
      Appearance      =   1
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
      ExplorerBar     =   0
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "人"
      Height          =   180
      Left            =   2160
      TabIndex        =   9
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "选中前面："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   8280
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "人数："
      Height          =   180
      Left            =   6840
      TabIndex        =   6
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检信息"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "frm健康证管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'查询条件。
Private mstr系统编号 As String
Private mstr姓名 As String
Private mstr单位 As String
Private mstr体检日期从 As String
Private mstr体检日期到 As String
Private mstr种类 As String

Private mobjRec As Object

Private Sub cchkType_Click(Index As Integer)
    subRefresh
End Sub

Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub ctbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lobj体检 As cls体检
    Dim i As Long
    Dim j As Long
    
    Select Case Button.Key
    Case "query"
        frm查询.Show 1, Me
        
        If frm查询.pblnOk Then
            mstr姓名 = frm查询.pstrName
            mstr系统编号 = frm查询.pstrNo
            mstr体检日期从 = frm查询.pstrStartDate
            mstr体检日期到 = frm查询.pstrEndDate
            mstr种类 = frm查询.pstrType
            
            subRefresh
        End If
        
    Case "new"
        frm体检录入.pstr系统编号 = ""
        frm体检录入.Show 1, Me
        '刷新界面。
        subRefresh
        
    Case "update"
        If cgrdMain.Row = 0 Then
            MsgBox "请选择要修改的体检人员！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        frm体检录入.pstr系统编号 = cgrdMain.TextMatrix(cgrdMain.Row, 0)
        frm体检录入.Show 1, Me
        
        '刷新界面。
        subRefresh
        
    Case "print" '打印健康证
        Dim lcolInfo As Collection
        Dim lstrCN As String
        
        '根据业务设置，判断是否需要自动生成健康证号。
        If pobj体检管理.业务设置("健康证带条码") = "是" Then
        
            '用户输入健康证号的起始号。
            lstrCN = InputBox("请输入健康证的起始号", "输入")
            If lstrCN = "" Then
                Exit Sub
            End If
            
            '判断输入健康证号是否为数字。
            Do While Not (IsNumeric(lstrCN))
                If MsgBox("你输入的健康证号格式不对。是否重新输入？", vbYesNo, "系统提示") = vbYes Then
                    lstrCN = InputBox("请输入健康证的起始号", "输入")
                Else
                    Exit Sub
                End If
            Loop
            
            '修改：2002-05-6（杨春）判断卡是否合法。
            Dim lobjEncrypt As Object
            Set lobjEncrypt = CreateObject("fycarddes.clsDataEncrypt")
            If Not lobjEncrypt.funcCheckJkzCardno(lstrCN) Then
                Err.Raise 6666, , "系统无法识别这张卡，请确定卡符合指定的格式或卡是否已损坏！"
            End If
            '不保存校验位。
            lstrCN = lobjEncrypt.卡号
            Set lobjEncrypt = Nothing
        Else
            lstrCN = pobj体检管理.func生成健康证号(False)
        End If
        
        
        '获取选中的系统编号，创建体检对象。
        Set lcolInfo = New Collection
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.Cell(flexcpChecked, i, 1) = flexChecked Then
                Set lobj体检 = New cls体检
                lobj体检.系统编号 = cgrdMain.TextMatrix(i, 0)
                lobj体检.健康证号 = lstrCN
                lcolInfo.Add lobj体检
                
                '健康证号自动递增。
                lstrCN = Format(Val(lstrCN) + 1, String(Len(lstrCN), "0"))
            End If
        Next
        
        If lcolInfo.Count = 0 Then
            MsgBox "请选择要打印的体检人员（姓名上打勾）！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        
        pobj体检管理.sub打印健康证 lcolInfo
        
        '刷新界面。
        subRefresh
        
    Case "delete"
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
        
    Case "dl" '调离
        frm调离管理.Show 1, Me
    Case "exit"
        End
    End Select
End Sub

Private Sub ctxtNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim llngNum As Long
    On Error GoTo errHandler
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
errHandler:
End Sub

Private Sub Form_Load()
    
    '获取最近两周的未打印体检记录。
    mstr体检日期从 = Format(DateAdd("d", 1 - DatePart("w", Now, vbMonday), Now) - 7, "yyyy-mm-dd")
    mstr体检日期到 = Format(Now, "yyyy-mm-dd")
    
    subRefresh
    
End Sub

'功能：根据查询条件显示查询结果。
Private Sub subRefresh()
    
    Dim lstr状态条件 As String
    Dim i As Long
    lstr状态条件 = ""
    If cchkType(0).Value = 1 Or cchkType(0).Value = 1 Then
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
        
    Set mobjRec = pobj体检管理.func健康体检查询(mstr系统编号, mstr姓名, mstr单位, mstr体检日期从, mstr体检日期到, mstr种类, lstr状态条件)
    
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
        mobjRec.movenext
    Next
    cgrdMain.ColWidth(1) = 1000
    
    '隐藏系统编号。
    cgrdMain.ColHidden(0) = True
    
    clblInfo.Caption = "人数：" & cgrdMain.Rows - 1

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
End Sub
