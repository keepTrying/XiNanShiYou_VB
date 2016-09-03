VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form FrmPrintBarCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打印新条码"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6645
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid cgrdSysNum 
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   1815
      _cx             =   2088766593
      _cy             =   2088768075
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "生成新条码号    "
      ScrollTrack     =   -1  'True
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
   Begin VB.CommandButton PrintBarCode 
      Caption         =   "打  印"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox PrintNumber 
      Height          =   270
      Left            =   3720
      TabIndex        =   1
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退  出"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "至"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label clblSysNo_Last 
      Caption         =   "12345678901234"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label clblSysNo_First 
      Caption         =   "12345678901234"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "人"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   645
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "打印"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   645
      Width           =   375
   End
End
Attribute VB_Name = "FrmPrintBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-02-25 于登淼 增加打印新条码窗体和相应功能
'输入打印的人数，自动生成所要打印的新条码号
'条码号生成规则同系统编号，倒数第5位标记为“1”。一般的体检则没有这一位

Option Explicit
'注意-->体检条码号 与 体检系统编号 几乎完全相同。只是在代码中不同位置，名称不同。
Private pstr系统编号 As String           '即存入数据库中的系统编号，也是接下来打印的体检条码号
Public pstrNumbers As Integer           '记录上一次写入的打印条码人数。用于进行系统编号退回的操作
Private pstr是否退回系统编号 As Boolean '若打印失败或没有打印，就退回系统编号，值为true；否则为false。

Private Sub ccmdExit_Click()
    Dim i As Integer
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病对象.clsmedicalexam")
    
    '没有打印，则退回当前生成的系统编号
    If pstr是否退回系统编号 = True Then
        For i = cgrdSysNum.Rows - 1 To 1 Step -1
            lobjTmp.Func退回职业病体检系统编号 (cgrdSysNum.Cell(flexcpText, i, 0))
            cgrdSysNum.RemoveItem (i)
        Next i
    End If
    
    '退出并清空窗口对象
    Unload Me
    Set FrmPrintBarCode = Nothing
End Sub

Private Sub Form_Load()
    Dim lobjTmp As Object
    On Error GoTo errHandler
    
    Set lobjTmp = CreateObject("职业病对象.clsMedicalExam")
    pstr是否退回系统编号 = True
    pstrNumbers = Val(PrintNumber.Text)
    PrintNumber.TabIndex = 1
    
    'form_load时，默认显示第一个系统编号
    pstr系统编号 = lobjTmp.Func分配职业病体检系统编号
    clblSysNo_First.Caption = pstr系统编号
    cgrdSysNum.AddItem pstr系统编号, 1
    sub生成并显示条码号 (1)
    
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmPrintBarCode", "Form_Load", 6666, lstrError, False
End Sub

Private Sub PrintBarCode_Click()
    Dim i As Integer

    On Error GoTo errHandler
    
    '批量打印一次过后，就不能修改打印人数，也不能继续打印。[不是很合理的设定吧？]
    PrintBarCode.Enabled = False
    PrintNumber.Enabled = False
    'For i = 1 To cgrdSysNum.Rows - 1
        'sub打印单个体检条码号 (cgrdSysNum.TextMatrix(i, 0))
    'Next i
    'pstr是否退回系统编号 = False
    
'2012-04-05 陶露
'打印多个条码
    Dim para体检条码号 As Collection
    Set para体检条码号 = New Collection
    For i = 1 To cgrdSysNum.Rows - 1
        para体检条码号.Add (cgrdSysNum.TextMatrix(i, 0))
    Next i
    sub打印多个体检条码号 para体检条码号
    pstr是否退回系统编号 = False
'2012-04-05 陶露
    
    Exit Sub

'错误处理：打印出错后，提示并退回系统编号。
errHandler:
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病对象.clsmedicalexam")
    sfsub错误处理 "职业病界面", "FrmPrintBarCode", "PrintBarCode_Click", Err.Number, Err.Description, False
    For i = cgrdSysNum.Rows - 1 To 1 Step -1
        lobjTmp.Func退回职业病体检系统编号 (cgrdSysNum.Cell(flexcpText, i, 0))
        cgrdSysNum.RemoveItem (i)
    Next i
    Exit Sub
End Sub

Private Sub PrintNumber_Change()
    Dim i, lobjInt, IfContinue As Integer
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病对象.clsmedicalexam")
    
    '判断输入格式
    If IsNumeric(PrintNumber.Text) = False Or CLng(Val(PrintNumber.Text)) <= 0 Then
        MsgBox ("人数必须为大于0的整数")
        Exit Sub
    End If
    
    '如果生成太多的编号，总编号超过9999的会出错。
    '由于，每天体检人员数 超过9999的可能性比较小，所以没有做更深入的修改，仅在这里提示。
    If Val(PrintNumber.Text) > 9000 Then IfContinue = MsgBox("一次生成这么多编号，体检编号可能不足。要继续吗？", vbYesNo)
    If IfContinue = vbNo Then
        PrintNumber.Text = CStr(pstrNumbers)
        Exit Sub
    End If
    
    '更改了人数，说明之前的体检系统编号没有打印。故需要退回，重新生成新编号。
    For i = Val(pstrNumbers) To 2 Step -1
        lobjTmp.Func退回职业病体检系统编号 (cgrdSysNum.Cell(flexcpText, i, 0))
        cgrdSysNum.RemoveItem (i)
    Next i
    sub生成并显示条码号 (Val(PrintNumber.Text))
    
End Sub

'将该次生成的所有条码号显示在窗体列表中
Sub sub生成并显示条码号(ByVal paraPrintNum As Integer)
    Dim i As Integer
    Dim lobjTmp As Object
    
    On Error GoTo errHandler
    
    Set lobjTmp = CreateObject("职业病对象.clsmedicalexam")
    For i = 2 To paraPrintNum
        cgrdSysNum.AddItem lobjTmp.Func分配职业病体检系统编号, cgrdSysNum.Rows
    Next i
    clblSysNo_Last.Caption = cgrdSysNum.Cell(flexcpText, cgrdSysNum.Rows - 1, 0)
    pstrNumbers = Val(PrintNumber.Text)
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "frmPrintBarCode", "sub生成并显示条码号", 6666, lstrError, True
End Sub


