VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm交款设置 
   Caption         =   "交款设置"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   12780
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox ctxtEnd 
      Height          =   270
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox ctxtBegin 
      Height          =   270
      Left            =   4920
      MaxLength       =   10
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox cchk显示已收款 
      Caption         =   "显示已收款"
      Height          =   255
      Left            =   7920
      TabIndex        =   13
      Top             =   720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox cchk显示应收款 
      Caption         =   "显示应收款"
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton ccmd保存 
      Caption         =   "保  存"
      Height          =   375
      Left            =   11160
      TabIndex        =   11
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton ccmd全部未交 
      Caption         =   "全部设为未核账"
      Height          =   495
      Left            =   11160
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox clstName 
      Height          =   300
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton ccmdQuery 
      Caption         =   "查  询"
      Height          =   375
      Left            =   10920
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton ccmd全部交款 
      Caption         =   "全部设为已核账"
      Height          =   495
      Left            =   11160
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   6285
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   10455
      _cx             =   25577481
      _cy             =   25570126
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
   Begin MSComCtl2.DTPicker cdtp截止日期 
      Height          =   300
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   25493505
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker cdtp开始日期 
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   25493505
      CurrentDate     =   36951
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   6000
      TabIndex        =   17
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票据号："
      Height          =   180
      Left            =   4200
      TabIndex        =   14
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   2520
      TabIndex        =   9
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期范围："
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收费员："
      Height          =   180
      Left            =   7800
      TabIndex        =   7
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "票据列表："
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "frm交款设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cchk显示已收款_Click()
    sub查询并显示记录
End Sub

Private Sub cchk显示应收款_Click()
    sub查询并显示记录
End Sub

Private Sub ccmdQuery_Click()
    sub查询并显示记录
End Sub

Private Sub ccmd保存_Click()
    Dim i As Integer
    Dim lobjRec As Object
    
    For i = 1 To cgrdMain.Rows - 1
        Set lobjRec = dafuncGetData("select * from 收费管理_票据交账记录表 where 票据号='" & cgrdMain.Cell(flexcpText, i, 0) & "'")
        If lobjRec.RecordCount > 0 Then
            dafuncGetData "update 收费管理_票据交账记录表 set 已交账=" & IIf(cgrdMain.Cell(flexcpChecked, i, 6) = flexChecked, "1", "0") & ",交账日期=" & IIf(cgrdMain.Cell(flexcpText, i, 7) = "", "null", "'" & cgrdMain.Cell(flexcpText, i, 7) & "'") & ",已收款=" & IIf(cgrdMain.Cell(flexcpChecked, i, 8) = flexChecked, "1", "0") & ",收款日期=" & IIf(cgrdMain.Cell(flexcpText, i, 9) = "", "null", "'" & cgrdMain.Cell(flexcpText, i, 9) & "'") & " where 票据号='" & cgrdMain.Cell(flexcpText, i, 0) & "'"
        Else
            dafuncGetData "insert into 收费管理_票据交账记录表 (票据号,已交账,交账日期,已收款,收款日期) values('" & cgrdMain.Cell(flexcpText, i, 0) & "'," & IIf(cgrdMain.Cell(flexcpChecked, i, 6) = flexChecked, "1", "0") & "," & IIf(cgrdMain.Cell(flexcpText, i, 7) = "", "null", "'" & cgrdMain.Cell(flexcpText, i, 7) & "") & "," & IIf(cgrdMain.Cell(flexcpChecked, i, 8) = flexChecked, "1", "0") & "," & IIf(cgrdMain.Cell(flexcpText, i, 9) = "", "null", "'" & cgrdMain.Cell(flexcpText, i, 9) & "'") & ")"
        End If
    Next
    
    MsgBox "保存成功！", vbInformation, "系统提示"
End Sub

Private Sub ccmd全部交款_Click()
    Dim i As Integer
    
    For i = 1 To cgrdMain.Rows - 1
        cgrdMain.Cell(flexcpChecked, i, 6) = flexChecked
        cgrdMain.Cell(flexcpText, i, 7) = Format(Date, "yyyy-mm-dd")
    Next
    cgrdMain.AutoSize 0, cgrdMain.Cols - 1
    MsgBox "修改后请注意进行保存！", vbInformation, "系统提示"
End Sub

Private Sub ccmd全部未交_Click()
    Dim i As Integer
    
    For i = 1 To cgrdMain.Rows - 1
        cgrdMain.Cell(flexcpChecked, i, 6) = flexUnchecked
        cgrdMain.Cell(flexcpText, i, 7) = ""
    Next
    MsgBox "修改后请注意进行保存！", vbInformation, "系统提示"
End Sub

Private Sub cgrdMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    
    If cgrdMain.Cell(flexcpChecked, Row, Col) = flexChecked Then
        cgrdMain.Cell(flexcpText, Row, Col + 1) = Format(Date, "yyyy-mm-dd")
    Else
        cgrdMain.Cell(flexcpText, Row, Col + 1) = ""
    End If
End Sub

Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 6 And Col <> 8 Then Cancel = True
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object, i As Integer
    Dim lobjRec1 As Object
    
    clstName.Clear
    clstName.AddItem ""
    Set lobjRec = dafuncGetData("select 编号,姓名 from 系统管理_员工基本信息视图 order by 编号")
    For i = 1 To lobjRec.RecordCount
        Set lobjRec1 = dafuncGetData("select * from 系统管理_用户操作权限表 where 用户编号='" & lobjRec(0) & "' and 权限名='收费管理_直接收费'")
        If lobjRec1.RecordCount > 0 Then
            clstName.AddItem lobjRec(0) & " " & lobjRec(1)
        End If
        lobjRec.MoveNext
    Next
    If clstName.ListCount > 0 Then
        clstName.ListIndex = 0
    Else
        MsgBox "当前没有设置具有收费权限的人员！", vbInformation, "系统提示"
    End If
    
    '默认显示当天的所有收费记录。
    cdtp开始日期.Value = Format(Date - 7, "yyyy-mm-dd")
    cdtp截止日期.Value = Format(Date, "yyyy-mm-dd")
    
    sub查询并显示记录

End Sub
Private Sub sub查询并显示记录()
    Dim lobjRec As Object, i As Integer
    Dim lstrWhere As String, lstrWhere1 As String
    
    On Error GoTo errhandle
    
    If cchk显示已收款.Value = Checked And cchk显示应收款.Value = Checked Then
        lstrWhere = "1=1"
    ElseIf cchk显示已收款.Value = Checked Then
        lstrWhere = " 已收款=1"
    ElseIf cchk显示应收款.Value = Checked Then
        lstrWhere = " (已收款=0 or 已收款 is null)"
    Else
        lstrWhere = "1=1"
    End If
    
    ctxtBegin = Trim(ctxtBegin)
    ctxtEnd = Trim(ctxtEnd)
    If ctxtBegin <> "" And ctxtEnd <> "" Then
        lstrWhere1 = " and 收据号 between ''" & ctxtBegin & "'' and ''" & ctxtEnd & "''"
    ElseIf ctxtBegin <> "" Then
        lstrWhere1 = " and 收据号 >= ''" & ctxtBegin & "''"
    ElseIf ctxtEnd <> "" Then
        lstrWhere1 = " and 收据号 <= ''" & ctxtEnd & "''"
    End If
    '查询收费记录.
    If clstName.ListIndex > 0 Then      '不针对收费员
        Set lobjRec = dafuncGetData("exec 收费管理_获取票据交账信息 '收费人=''" & Left(clstName.Text, InStr(clstName.Text, " ") - 1) & "'' and 收费日期 between ''" & Format(cdtp开始日期.Value, "yyyy-mm-dd") & "'' and ''" & Format(cdtp截止日期.Value, "yyyy-mm-dd") & "''" & lstrWhere1 & "','" & lstrWhere & "'")
    Else
        Set lobjRec = dafuncGetData("exec 收费管理_获取票据交账信息 '收费日期 between ''" & Format(cdtp开始日期.Value, "yyyy-mm-dd") & "'' and ''" & Format(cdtp截止日期.Value, "yyyy-mm-dd") & "''" & lstrWhere1 & "','" & lstrWhere & "'")
    End If
    Set cgrdMain.DataSource = lobjRec
    cgrdMain.ColFormat(4) = "999999.00"
    cgrdMain.ColDataType(6) = flexDTBoolean
    cgrdMain.ColDataType(8) = flexDTBoolean
    For i = 1 To cgrdMain.Rows - 1
        cgrdMain.Cell(flexcpChecked, i, 6) = IIf(cgrdMain.Cell(flexcpText, i, 6) = "1", flexChecked, flexUnchecked)
        cgrdMain.Cell(flexcpChecked, i, 8) = IIf(cgrdMain.Cell(flexcpText, i, 8) = "1", flexChecked, flexUnchecked)
    Next
    cgrdMain.AutoSize 0, cgrdMain.Cols - 1
    Exit Sub
errhandle:
    Call sfsub错误处理("收费界面部件", "frm收费管理", "sub查询并显示记录()", Err.Number, Err.Description, True)
    Exit Sub
    Resume
End Sub

