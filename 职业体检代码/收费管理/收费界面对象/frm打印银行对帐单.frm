VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm打印银行对帐单 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "打印银行对帐单"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton coptType 
      Caption         =   "门诊收费"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "一般性缴费"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   17
      Top             =   720
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CheckBox cchkAll 
      Caption         =   "全选"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "关闭(&C)"
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "打印(&P)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton ccmdQuery 
      Caption         =   "查询(&Q)"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   720
      Width           =   1000
   End
   Begin VB.ComboBox ccmb开户行 
      Height          =   300
      Left            =   960
      TabIndex        =   9
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox ctxt票据号 
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox ctxt票据号 
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker cdtp截止日期 
      Height          =   300
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   25296897
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker cdtp开始日期 
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   25296897
      CurrentDate     =   36951
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
      Height          =   6420
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   10605
      _cx             =   163793170
      _cy             =   163785788
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
      BackColorAlternate=   16437167
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      TabBehavior     =   1
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
   Begin VB.Label clblTotal 
      AutoSize        =   -1  'True
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   14
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "总金额(元)："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   7920
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "开户银行"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Index           =   1
      Left            =   7440
      TabIndex        =   6
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "票据号："
      Height          =   180
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frm打印银行对帐单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrID As String
Private Sub cchkAll_Click()
    Dim i As Long

    For i = 1 To cgrdDetail.Rows - 1
        cgrdDetail.Cell(flexcpChecked, i, 2) = IIf(cchkAll.Value = 1, flexChecked, flexUnchecked)
    Next
    
    sub显示总额
End Sub

Private Sub sub显示总额()
    Dim i As Long
    Dim ldblTotal As Double
    Dim lIndex As Long
    
    For i = 0 To cgrdDetail.Cols - 1
        If cgrdDetail.TextMatrix(0, i) = "金额" Then
            lIndex = i
            Exit For
        End If
    Next
    For i = 1 To cgrdDetail.Rows - 1
        If cgrdDetail.Cell(flexcpChecked, i, 2) = flexChecked Then
            ldblTotal = Format(ldblTotal + cgrdDetail.ValueMatrix(i, lIndex), "0.00")
        End If
    Next
    clblTotal = ldblTotal
End Sub
Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdPrint_Click()
    Dim lobjRec As Object
    Dim lCount As Long
    Dim lcolItem As Collection
    Dim lcolInfo As Collection
    Dim lobj打印 As Object
    Dim i As Long
    On Error GoTo errhandler
    
    '把数据导入临时表。
    If mstrID <> "" Then dafuncGetData "delete temp_银行对帐单 where ID='" & mstrID & "'"
    
    Set lobjRec = dafuncGetData("select newid()")
    mstrID = lobjRec(0)
    lCount = 0
    For i = 1 To cgrdDetail.Rows - 1
        If cgrdDetail.Cell(flexcpChecked, i, 2) = flexChecked Then
            dafuncGetData "insert into temp_银行对帐单(ID,收费批号,收费项目编号) values('" & mstrID & "','" & cgrdDetail.TextMatrix(i, 0) & "','" & cgrdDetail.TextMatrix(i, 1) & "')"
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        sffuncMsg "请在要打印的记录前打勾！"
    Else
        Dim lstrFilter As String
        lstrFilter = IIf(IsNull(cdtp开始日期.Value), "", "日期：" & cdtp开始日期.Value) & IIf(IsNull(cdtp截止日期.Value), "", " 至：" & cdtp截止日期.Value) & IIf(ctxt票据号(0) <> "" Or ctxt票据号(1) <> "", "     票据号：" & ctxt票据号(0) & " 至 " & ctxt票据号(1), "") & IIf(ccmb开户行 = "", "", " 开户银行：" & ccmb开户行) & IIf(coptType(0).Value, "   一般性缴费", "   门诊收费")
        
        Set lcolInfo = New Collection
        Set lcolItem = New Collection
        lcolItem.Add "ID", "名称"
        lcolItem.Add mstrID, "值"
        lcolInfo.Add lcolItem
        
        Set lcolItem = New Collection
        lcolItem.Add "条件", "名称"
        lcolItem.Add lstrFilter, "值"
        lcolInfo.Add lcolItem
        
        Set lobj打印 = CreateObject("通用水晶文书打印.cls文书")
        lobj打印.funcPrintReport "银行对帐单", lcolInfo, App.Path, True
        
    End If
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm打印银行对帐单", "ccmdPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ccmdQuery_Click()
    Dim lobjRec As Object
    Dim i As Long
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData("exec 收费管理_查询银行对帐单 '" & IIf(IsNull(cdtp开始日期.Value), "", cdtp开始日期.Value) & "','" & IIf(IsNull(cdtp截止日期.Value), "", cdtp截止日期.Value) & "','" & ctxt票据号(0) & "','" & ctxt票据号(1) & "','" & ccmb开户行 & "','" & IIf(coptType(0).Value, "一般", "门诊") & "'")
    cgrdDetail.FormatString = ""
    Set cgrdDetail.DataSource = lobjRec
    
    cgrdDetail.Editable = True

    cgrdDetail.ColHidden(0) = True
    cgrdDetail.ColHidden(1) = True
    
    For i = 1 To cgrdDetail.Rows - 1
        cgrdDetail.Cell(flexcpChecked, i, 2) = IIf(cchkAll.Value = 1, flexChecked, flexUnchecked)
    Next
    cgrdDetail.ColWidth(0) = 1200
    sub显示总额
    If cgrdDetail.Rows > 0 Then
        
        ccmdPrint.Enabled = True
    Else
        ccmdPrint.Enabled = False
    End If
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm打印银行对帐单", "ccmdQuery_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cgrdDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    sub显示总额
End Sub

Private Sub cgrdDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 Then Cancel = True
    
End Sub

Private Sub coptType_Click(Index As Integer)
    If coptType(0).Value Then
        ccmb开户行.Enabled = True
    Else
        ccmb开户行.Enabled = False
        ccmb开户行.ListIndex = -1
    End If
    
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    On Error GoTo errhandler
    '获取开户行帐号.
    Set lobjRec = dafuncGetData("select 开户行+' '+帐号 from 收费管理_银行开户行设置表")
    ccmb开户行.Clear
    ccmb开户行.AddItem ""
    Do While Not lobjRec.EOF
        ccmb开户行.AddItem lobjRec(0)
        
        lobjRec.MoveNext
    Loop
    
    cdtp开始日期.Value = Format(Now, "yyyy/mm/dd")
    cdtp截止日期.Value = Format(Now, "yyyy/mm/dd")
    cgrdDetail.Editable = True
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm打印银行对帐单", "Form_Load", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If mstrID <> "" Then dafuncGetData "delete temp_银行对帐单 where ID='" & mstrID & "'"
End Sub
