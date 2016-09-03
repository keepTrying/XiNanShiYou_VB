VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form frm明细 
   Caption         =   "收费明细"
   ClientHeight    =   5220
   ClientLeft      =   6930
   ClientTop       =   5550
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5550
   Begin VB.CommandButton ccmd返回 
      Caption         =   "返回"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdDetail 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   5535
      _cx             =   9763
      _cy             =   6588
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
      AutoSearchDelay =   2
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   3240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchk预览 
         Caption         =   "打印前预览"
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   120
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收费批号："
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "总金额："
      Height          =   180
      Left            =   2640
      TabIndex        =   2
      Top             =   4800
      Width           =   720
   End
End
Attribute VB_Name = "frm明细"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pNo As String    '收费批号
Dim WithEvents mobj界面通用对象 As cls界面通用对象
Attribute mobj界面通用对象.VB_VarHelpID = -1

Private Sub ccmd返回_Click()
    
    Label1.Caption = "收费批号："
    Label2.Caption = "总金额："
    cgrdDetail.Rows = 1
    Unload Me
End Sub

Private Sub mobj界面通用对象_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandle
    Dim lobjRec As Object
    Dim i As Integer
    Select Case Operate
    Case "导出Excel"
        Dim lstrFile As String
        ccmdFile.Filter = "Excel文件 (*.xls)|*.xls|文本文件 (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.filename
        If lstrFile <> "" Then
'            cgrdMain.ColDataType(mcolIndex("系统编号")) = flexDTString
'            cgrdMain.ColDataType(mcolIndex("体检表号")) = flexDTString
            cgrdDetail.SaveGrid lstrFile, flexFileExcel, True
        End If
'        MsgBox "导出成功！", vbOKOnly, "系统提示"
    Case "打印"
            cgrdDetail.PrintGrid
    Case "返回"
        ccmd返回_Click
    End Select
    Exit Sub
errHandle:
    sffuncMsg Operate & "不成功。" & Err.Description, sf警告
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lsngTotal As Single '总金额。
    Dim HealthNo As String
    Dim lcol工具栏按钮 As Collection
    Dim lLen As Integer
'    pblnInUse = True                              '指示窗体已启动
'
    '初始化工具栏
    Set mobj界面通用对象 = New cls界面通用对象
    Set mobj界面通用对象.Form = Me
    Set mobj界面通用对象.c工具栏 = ctlb工具栏
    Set lcol工具栏按钮 = New Collection
    lcol工具栏按钮.Add "导出Excel(&D)111"
    lcol工具栏按钮.Add "打印"
    lcol工具栏按钮.Add "退出"
    mobj界面通用对象.subInitialize lcol工具栏按钮, ""
    
    
    Label1.Caption = Label1.Caption & pNo
    If pNo <> "" Then
        
        Set lobjRec = dafuncGetData("select 系统编号 from 职业病体检_体检基本信息表 where 收费批号 = '" & pNo & "'")
        If Not lobjRec.EOF Then
            
            HealthNo = lobjRec("系统编号")
            
            Set lobjRec = dafuncGetData("select a.名称 as 项目名称,a.单价" _
                & ",b.体检类型,b.体检类别 from 职业病体检_体检项目设置表 a,职业病体检_体检收费视图 b " _
                & "where a.编码 = b.体检项目 and 系统编号 = '" & HealthNo & "'")
            If lobjRec.EOF Then Exit Sub
            
'            Do While Not lobjRec.EOF
'                If IsNull(lobjRec("单价")) = False Then
'                lsngTotal = Format(lsngTotal + lobjRec("单价"), "0.00")
'                End If
'                lobjRec.MoveNext
'            Loop
            Set cgrdDetail.DataSource = lobjRec
            cgrdDetail.SubtotalPosition = flexSTBelow
            cgrdDetail.Subtotal flexSTSum, 0, 1, , , vbRed, True, "合计", 1, True
            
            cgrdDetail.AutoResize = True
            cgrdDetail.AutoSize 0, cgrdDetail.Cols - 1
            
            Label2.Caption = Label2.Caption & cgrdDetail.TextMatrix(cgrdDetail.Rows - 1, 1) & " 元"
            
        End If
    
    End If
    
End Sub
