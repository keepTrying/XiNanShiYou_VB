VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm测试对象 
   Caption         =   "测试"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   8865
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "Sub数据导入"
      Height          =   435
      Left            =   5040
      TabIndex        =   12
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Func获取mdb数据交换文件中的数据范围"
      Height          =   435
      Left            =   5040
      TabIndex        =   11
      Top             =   5040
      Width           =   3495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Func获取mdb文件中的数据分类清单"
      Height          =   435
      Left            =   5040
      TabIndex        =   10
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "func查看数据"
      Height          =   435
      Left            =   5040
      TabIndex        =   9
      Top             =   4560
      Width           =   3495
   End
   Begin VSFlex6DAOCtl.vsFlexGrid cgrdMain 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3625
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   7320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command6 
      Caption         =   "sub更改文件名"
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   1995
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sub导入体检人员登记"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   1995
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Func读报盘文件中体检人员信息"
      Height          =   435
      Left            =   2160
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置体检访问标志"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "func获取已体检完毕信息"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "属性"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1755
   End
End
Attribute VB_Name = "frm测试对象"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobj体检对外接口 As ClsManageTransmission

Private Sub Command8_Click()
    Dim lobjRec As Object
    Dim lcolRange As Collection
    Dim lcolItem As Collection
    Dim llngType As Integer
    If MsgBox("查看mdb库中数据吗？若选择“是”，则查看mdb库，否则查看网络数据库。", vbYesNo + vbQuestion + vbDefaultButton1, "系统询问") = vbYes Then
        llngType = 0
    Else
        llngType = 1
    End If
    Set lcolRange = New Collection
    Set lcolItem = New Collection
    lcolItem.Add "开始日期", "数据范围名"
    lcolItem.Add "2001-3-27", "数据范围值"
    lcolRange.Add lcolItem, lcolItem("数据范围名")
    
    Set lcolItem = New Collection
    lcolItem.Add "结束日期", "数据范围名"
    lcolItem.Add "2001-3-28", "数据范围值"
    lcolRange.Add lcolItem, lcolItem("数据范围名")
    
    Set lcolItem = New Collection
    lcolItem.Add "单位名称集", "数据范围名"
    lcolItem.Add "asd", "数据范围值"
    lcolRange.Add lcolItem, lcolItem("数据范围名")
    
    Set lcolItem = New Collection
    lcolItem.Add "系统编号范围", "数据范围名"
    lcolItem.Add "0103280204,0103280304", "数据范围值"
    lcolRange.Add lcolItem, lcolItem("数据范围名")
    Set lobjRec = mobj体检对外接口.Func查看数据(lcolRange, llngType)
    
    Set mobj体检对外接口 = Nothing
    
    gfsubLoadGridFromRec cgrdMain, lobjRec
    
End Sub

Private Sub Command9_Click()
    Dim lcolRange As Collection
    Dim lobjRec As Object
    Set lobjRec = mobj体检对外接口.Func查看数据(lcolRange, 0)
    
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    '初始化数据访问对象(连接testserver)。
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=防疫2001;Data Source=ychun"
    
    Set mobj体检对外接口 = New ClsManageTransmission
    
    If Not umfunc校验身份("7612", "") Then
        sffuncMsg "校验身份失败。", sf警告
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Form_load", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub



Private Sub Command1_Click()
    Dim lobj工作站配置 As Object
    On Error GoTo errHandler
    
    Text1.Text = "业务设置：是否打印体检单=" & mobj体检对外接口.业务设置("是否打印体检单") & Chr(13) & Chr(10)
    Set lobj工作站配置 = mobj体检对外接口.工作站配置
    Text1.Text = Text1.Text & "本地配置：内部导出文件=" & lobj工作站配置.内部导出文件
    
    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Command1_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command2_Click()
    Dim a As New Collection
    Dim b As New Collection
    Dim e As New Collection
    Dim lstrTemp As String
    On Error GoTo errHandler
    Set a = mobj体检对外接口.Func获取已完毕的体检信息("", "", "", "", , "从业人员体检表", "健康证", "")
    For Each b In a
        lstrTemp = lstrTemp & "系统编号:" & b("系统编号") & ", 姓名:" & b("姓名") & ", 单位名称:" & b("单位名称") & vbCrLf
        For Each e In b("附加项目")
            lstrTemp = lstrTemp & "项目名" & e("项目名") & ",  "
        Next
        lstrTemp = lstrTemp & vbCrLf & vbCrLf
    Next
    Text1.Text = lstrTemp

    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Command1_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command3_Click()
    On Error GoTo errHandler
    mobj体检对外接口.Sub设置体检访问标志 "111110103270106", "健康证1", "1"
    Exit Sub
errHandler:
    sfsub错误处理 "", "", "", Err.Number, Err.Description, False
End Sub

Private Sub Command4_Click()
    Dim c As String
    Dim i As Integer
    Dim b As New Collection
    Dim a As New Collection
    
    On Error GoTo errHandler
    
    Set a = mobj体检对外接口.Func读报盘文件中体检人员信息(App.Path & "\book1.xls")
    For Each b In a
        c = c & "单位名称:" & b("单位名称") & "                 "
        c = c & "姓    名:" & b("姓名") & vbCrLf
        c = c & "年    龄:" & b("年龄") & "                   "
        c = c & "性    别:" & b("性别") & vbCrLf
    Next b
    Text1.Text = c
    
    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Command4_Click", Err.Number, Err.Description
    Exit Sub
    Resume
    
End Sub

Private Sub Command5_Click()
    Dim a As New Collection
    On Error GoTo errHandler
    
    Set a = mobj体检对外接口.Func读报盘文件中体检人员信息(App.Path & "\book1.xls")
    mobj体检对外接口.Sub导入体检人员登记 "放射人员体检", ProgressBar1, a
    
    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Command5_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command6_Click()
    On Error GoTo errHandler
    mobj体检对外接口.sub更改文件名 App.Path & "\Form1.log"
    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Command6_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command7_Click()
    Dim a As New Collection
    Dim b As New Collection
    
    On Error GoTo errHandler
    
    b.Add "道源,大公司", "单位名称集"
    b.Add "00000103300001,00000103300099", "系统编号范围"

    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Command7_Click", Err.Number, Err.Description
    Exit Sub
    Resume
    
End Sub


