VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frm开户行设置 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "开户行设置"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ccmdExit 
      Caption         =   "返回(&X)"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton ccmdDelete 
      Caption         =   "删除(&D)"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton ccmdSave 
      Caption         =   "保存(&S)"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   3405
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      _cx             =   70004963
      _cy             =   69998454
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   15791081
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "帐户名称           |帐号                  |开户银行                       "
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
End
Attribute VB_Name = "frm开户行设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ccmdDelete_Click()
    On Error GoTo errHandler
    If cgrdMain.Row > 0 And cgrdMain.TextMatrix(cgrdMain.Row, 1) <> "" Then
        dafuncGetData "delete 收费管理_银行开户行设置表  where 帐号='" & cgrdMain.TextMatrix(cgrdMain.Row, 1) & "'"
        cgrdMain.RemoveItem cgrdMain.Row
    End If
    Exit Sub
    
    
errHandler:
    sfsub错误处理 "收费界面部件", "frm开户行设置", "ccmdDelete_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ccmdExit_Click()
    Unload Me
    
End Sub

Private Sub ccmdSave_Click()
    Dim i As Long
    On Error GoTo errHandler
    
    dasubBeginTran
    dafuncGetData "delete 收费管理_银行开户行设置表"
    With cgrdMain
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" And .TextMatrix(i, 1) <> "" Then
                dafuncGetData "insert into 收费管理_银行开户行设置表(帐户名称,帐号,开户行) values('" & .TextMatrix(i, 0) & "'," _
                    & "'" & .TextMatrix(i, 1) & "','" & .TextMatrix(i, 2) & "')"
            
            End If
        Next
    End With
    dasubCommitTran
    MsgBox "保存成功！", vbOKOnly + vbInformation, "系统提示"
    Exit Sub
errHandler:
    dasubRollBack
    sfsub错误处理 "收费界面部件", "frm开户行设置", "ccmdSave_Click", Err.Number, Err.Description, False
End Sub

Private Sub cgrdMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '如果编辑的是最后一行，自动追加一个空行。
    On Error Resume Next
    If cgrdMain.TextMatrix(Row, 0) <> "" And cgrdMain.TextMatrix(Row, 1) <> "" And Row = cgrdMain.Rows - 1 Then
        cgrdMain.Rows = cgrdMain.Rows + 1
    End If

End Sub
Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    Dim lstrName As String
    lstrName = cgrdMain.TextMatrix(0, Col)
    If lstrName <> "帐号" Then
        sub恢复中文输入法
    End If
End Sub
Private Sub cgrdMain_Click()
    On Error Resume Next
    If cgrdMain.Row > 0 Then
        cgrdMain.EditCell
    End If
End Sub


Private Sub Form_Load()
    Dim lobjRec As Object
    Dim i As Long
    On Error GoTo errHandler
    Set lobjRec = dafuncGetData("select * from 收费管理_银行开户行设置表")
    cgrdMain.Rows = lobjRec.RecordCount + 2
    i = 1
    Do While Not lobjRec.EOF
        cgrdMain.TextMatrix(i, 0) = lobjRec!帐户名称
        cgrdMain.TextMatrix(i, 1) = lobjRec!帐号
        cgrdMain.TextMatrix(i, 2) = IIf(IsNull(lobjRec!开户行), "", lobjRec!开户行)
        i = i + 1
        lobjRec.MoveNext
    Loop
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm开户行设置", "Form_Load", Err.Number, Err.Description, False
    
End Sub
