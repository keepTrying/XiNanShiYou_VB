VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm调离管理 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "调离管理"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   Icon            =   "frm调离管理.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox cchkPreview 
      Caption         =   "打印前预览"
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0FFFF&
      Caption         =   "已打印"
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H80000009&
      Caption         =   "未打印"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   12
      Top             =   960
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox ctxt备注 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   7320
      Width           =   8415
   End
   Begin VB.TextBox ctxt调离日期 
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox ctxt调离期限 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox ctxt调离编号 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   6840
      Width           =   2055
   End
   Begin VB.ListBox clstUnit 
      Height          =   5280
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   5415
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   7725
      _cx             =   25310522
      _cy             =   25306447
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
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
      FormatString    =   "编号   |姓名    |性别    |年龄    |单位名称     |卫生种类    |行业类别    |职业    |检出病种   | 体检结论 "
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
   Begin MSComctlLib.Toolbar C工具栏 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
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
      Caption         =   "备注："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   7440
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调离日期："
      Height          =   180
      Index           =   2
      Left            =   6600
      TabIndex        =   10
      Top             =   6960
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调离时限(天)："
      Height          =   180
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   6960
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调离编号："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调离名单："
      Height          =   180
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位名称："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   900
   End
End
Attribute VB_Name = "frm调离管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjGUI As cls界面通用对象 '界面上引用的界面通用对
Attribute mobjGUI.VB_VarHelpID = -1



'功能：控制不能输入单印号，处理回车。
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        SendKeys Chr(9)
    ElseIf KeyCode = 39 Then
        KeyCode = 0
    End If
    

End Sub
Private Sub cchkType_Click(Index As Integer)
    On Error GoTo errhandler
    subRefresh
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm调离管理", "cchkType_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub clstUnit_Click()
    Dim i As Long
    Dim lobjRec As Object
    On Error GoTo errhandler
    
    '显示选中单位的调离人员。
    Set lobjRec = pobj体检管理.func获取调离人员(clstUnit.List(clstUnit.ListIndex))
    cgrdMain.FormatString = ""
    Set cgrdMain.DataSource = lobjRec
    cgrdMain.ColHidden(0) = True
    If cgrdMain.Rows > 1 Then
        ctxt调离编号.Text = cgrdMain.TextMatrix(1, 1)
        ctxt调离日期.Text = cgrdMain.TextMatrix(1, cgrdMain.Cols - 4)
        ctxt调离期限.Text = cgrdMain.TextMatrix(1, cgrdMain.Cols - 3)
        ctxt备注.Text = cgrdMain.TextMatrix(1, cgrdMain.Cols - 2)
    End If
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm调离管理", "clstUnit_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Load()
        
    On Error GoTo errhandler
    
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.C工具栏 = C工具栏
    lcol工具栏按钮.Add "刷新"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "打印"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    '获取所有有调离人员的单位。
    subRefresh
    
    ctxt调离日期.Text = Format(Date, "yyyy-mm-dd")
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm调离管理", "Form_Load", Err.Number, Err.Description, False
    
End Sub
Private Sub subRefresh()
    Dim lobjRec As Object
    Dim lstr状态条件 As String
    
    On Error GoTo errhandler
    
    cgrdMain.Rows = 1
    
    If cchkType(0).Value = 1 And cchkType(1).Value = 0 Then
        lstr状态条件 = "(状态='未打印' or isnull(调离编号,'')='')"
    ElseIf cchkType(0).Value = 0 And cchkType(1).Value = 1 Then
        lstr状态条件 = "状态='已打印'"
    End If
    Set lobjRec = pobj体检管理.func获取调离单位(lstr状态条件)
    clstUnit.Clear
    Do While Not lobjRec.EOF
        clstUnit.AddItem lobjRec(0).Value
        lobjRec.MoveNext
    Loop
    If clstUnit.ListCount > 0 Then clstUnit.ListIndex = 0

    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm调离管理", "subRefresh", Err.Number, Err.Description, True
End Sub



Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Select Case Operate
    Case "刷新"
        subRefresh
    Case "打印"
        Dim i As Long
        Dim lcolInfo As Collection
        Dim lstr调离编号  As String
        
        If cgrdMain.Rows = 1 Then
            MsgBox "无内容可打印！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If
        If ctxt调离期限.Text = "" Then
            MsgBox "请输入调离时限！", vbOKOnly + vbExclamation, "系统提示"
            ctxt调离期限.SetFocus
            Exit Sub
        End If
        If ctxt调离日期.Text = "" Then
            MsgBox "请输入调离日期！", vbOKOnly + vbExclamation, "系统提示"
            ctxt调离日期.SetFocus
            Exit Sub
        End If
        '保存调离编号等信息。
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.TextMatrix(i, 1) = "" Then
                If lstr调离编号 = "" Then
                    lstr调离编号 = pobj体检管理.func生成调离编号(cgrdMain.TextMatrix(i, 0))
                End If
                dafuncGetData "update 健康证管理_办证申请信息表 set 调离编号='" & lstr调离编号 & "',调离期限=" & ctxt调离期限.Text & ",调离日期='" & ctxt调离日期.Text & "',备注='" & ctxt备注.Text & "' where 系统编号='" & cgrdMain.TextMatrix(i, 0) & "'"
            Else
                lstr调离编号 = cgrdMain.TextMatrix(i, 1)
            End If
        Next
        pobj体检管理.sub打印调离通知 lstr调离编号
    
    End Select
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm调离管理", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
End Sub
