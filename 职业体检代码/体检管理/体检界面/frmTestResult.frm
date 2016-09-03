VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "录入控件.ocx"
Begin VB.Form frmTestResult 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin 录入控件.ctlInputFrame c显示板 
      Height          =   2400
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4233
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Cols            =   51
      DistanceofRow   =   0
      FormatString    =   $"frmTestResult.frx":0000
      Count           =   9
      titleInputBox0001=   "姓名"
      statusinfoInputBox0001=   ""
      lengthInputBox0001=   8
      orderInputBox0001=   1
      valueInputBox0001=   ""
      datatypeInputBox0001=   3
      colInputBox0001 =   1
      rowInputBox0001 =   1
      PassWordCharInputBox0001=   0   'False
      主键InputBox0001=   0   'False
      允许等于最大值InputBox0001=   0   'False
      允许等于最小值InputBox0001=   0   'False
      字典名称InputBox0001=   ""
      显示字典字段InputBox0001=   ""
      保存字典字段InputBox0001=   ""
      名称InputBox0001=   "姓名"
      缺省值InputBox0001=   ""
      保存缺省值InputBox0001=   ""
      长度InputBox0001=   0
      MaxInputBox0001 =   ""
      MinInputBox0001 =   ""
      VisibleInputBox0001=   -1  'True
      PermitNullInputBox0001=   -1  'True
      TriggerstrInputBox0001=   ""
      EnableInputBox0001=   0   'False
      允许多选InputBox0001=   0   'False
      titleInputBox0002=   "性别"
      statusinfoInputBox0002=   ""
      lengthInputBox0002=   4
      orderInputBox0002=   2
      valueInputBox0002=   ""
      datatypeInputBox0002=   3
      colInputBox0002 =   10
      rowInputBox0002 =   1
      PassWordCharInputBox0002=   0   'False
      主键InputBox0002=   0   'False
      允许等于最大值InputBox0002=   0   'False
      允许等于最小值InputBox0002=   0   'False
      字典名称InputBox0002=   ""
      显示字典字段InputBox0002=   ""
      保存字典字段InputBox0002=   ""
      名称InputBox0002=   "性别"
      缺省值InputBox0002=   ""
      保存缺省值InputBox0002=   ""
      长度InputBox0002=   0
      MaxInputBox0002 =   ""
      MinInputBox0002 =   ""
      VisibleInputBox0002=   -1  'True
      PermitNullInputBox0002=   -1  'True
      TriggerstrInputBox0002=   ""
      EnableInputBox0002=   0   'False
      允许多选InputBox0002=   0   'False
      titleInputBox0003=   "年龄"
      statusinfoInputBox0003=   ""
      lengthInputBox0003=   4
      orderInputBox0003=   3
      valueInputBox0003=   ""
      datatypeInputBox0003=   3
      colInputBox0003 =   15
      rowInputBox0003 =   1
      PassWordCharInputBox0003=   0   'False
      主键InputBox0003=   0   'False
      允许等于最大值InputBox0003=   0   'False
      允许等于最小值InputBox0003=   0   'False
      字典名称InputBox0003=   ""
      显示字典字段InputBox0003=   ""
      保存字典字段InputBox0003=   ""
      名称InputBox0003=   "年龄"
      缺省值InputBox0003=   ""
      保存缺省值InputBox0003=   ""
      长度InputBox0003=   0
      MaxInputBox0003 =   ""
      MinInputBox0003 =   ""
      VisibleInputBox0003=   -1  'True
      PermitNullInputBox0003=   -1  'True
      TriggerstrInputBox0003=   ""
      EnableInputBox0003=   0   'False
      允许多选InputBox0003=   0   'False
      titleInputBox0004=   "单位名称"
      statusinfoInputBox0004=   ""
      lengthInputBox0004=   30
      orderInputBox0004=   4
      valueInputBox0004=   ""
      datatypeInputBox0004=   3
      colInputBox0004 =   20
      rowInputBox0004 =   1
      PassWordCharInputBox0004=   0   'False
      主键InputBox0004=   0   'False
      允许等于最大值InputBox0004=   0   'False
      允许等于最小值InputBox0004=   0   'False
      字典名称InputBox0004=   ""
      显示字典字段InputBox0004=   ""
      保存字典字段InputBox0004=   ""
      名称InputBox0004=   "单位名称"
      缺省值InputBox0004=   ""
      保存缺省值InputBox0004=   ""
      长度InputBox0004=   0
      MaxInputBox0004 =   ""
      MinInputBox0004 =   ""
      VisibleInputBox0004=   -1  'True
      PermitNullInputBox0004=   -1  'True
      TriggerstrInputBox0004=   ""
      EnableInputBox0004=   0   'False
      允许多选InputBox0004=   0   'False
      titleInputBox0005=   "体检类别"
      statusinfoInputBox0005=   ""
      lengthInputBox0005=   10
      orderInputBox0005=   5
      valueInputBox0005=   ""
      datatypeInputBox0005=   3
      colInputBox0005 =   1
      rowInputBox0005 =   2
      PassWordCharInputBox0005=   0   'False
      主键InputBox0005=   0   'False
      允许等于最大值InputBox0005=   0   'False
      允许等于最小值InputBox0005=   0   'False
      字典名称InputBox0005=   ""
      显示字典字段InputBox0005=   ""
      保存字典字段InputBox0005=   ""
      名称InputBox0005=   "体检类别"
      缺省值InputBox0005=   ""
      保存缺省值InputBox0005=   ""
      长度InputBox0005=   0
      MaxInputBox0005 =   ""
      MinInputBox0005 =   ""
      VisibleInputBox0005=   -1  'True
      PermitNullInputBox0005=   -1  'True
      TriggerstrInputBox0005=   ""
      EnableInputBox0005=   0   'False
      允许多选InputBox0005=   0   'False
      titleInputBox0006=   "体检日期"
      statusinfoInputBox0006=   ""
      lengthInputBox0006=   12
      orderInputBox0006=   6
      valueInputBox0006=   ""
      datatypeInputBox0006=   3
      colInputBox0006 =   12
      rowInputBox0006 =   2
      PassWordCharInputBox0006=   0   'False
      主键InputBox0006=   0   'False
      允许等于最大值InputBox0006=   0   'False
      允许等于最小值InputBox0006=   0   'False
      字典名称InputBox0006=   ""
      显示字典字段InputBox0006=   ""
      保存字典字段InputBox0006=   ""
      名称InputBox0006=   "体检日期"
      缺省值InputBox0006=   ""
      保存缺省值InputBox0006=   ""
      长度InputBox0006=   0
      MaxInputBox0006 =   ""
      MinInputBox0006 =   ""
      VisibleInputBox0006=   -1  'True
      PermitNullInputBox0006=   -1  'True
      TriggerstrInputBox0006=   ""
      EnableInputBox0006=   0   'False
      允许多选InputBox0006=   0   'False
      titleInputBox0007=   "体检结论"
      statusinfoInputBox0007=   ""
      lengthInputBox0007=   25
      orderInputBox0007=   7
      valueInputBox0007=   ""
      datatypeInputBox0007=   3
      colInputBox0007 =   25
      rowInputBox0007 =   2
      PassWordCharInputBox0007=   0   'False
      主键InputBox0007=   0   'False
      允许等于最大值InputBox0007=   0   'False
      允许等于最小值InputBox0007=   0   'False
      字典名称InputBox0007=   ""
      显示字典字段InputBox0007=   ""
      保存字典字段InputBox0007=   ""
      名称InputBox0007=   "体检结论"
      缺省值InputBox0007=   ""
      保存缺省值InputBox0007=   ""
      长度InputBox0007=   0
      MaxInputBox0007 =   ""
      MinInputBox0007 =   ""
      VisibleInputBox0007=   -1  'True
      PermitNullInputBox0007=   -1  'True
      TriggerstrInputBox0007=   ""
      EnableInputBox0007=   0   'False
      允许多选InputBox0007=   0   'False
      titleInputBox0008=   "诊断和处理意见"
      statusinfoInputBox0008=   ""
      lengthInputBox0008=   30
      orderInputBox0008=   8
      valueInputBox0008=   ""
      datatypeInputBox0008=   3
      colInputBox0008 =   1
      rowInputBox0008 =   3
      PassWordCharInputBox0008=   0   'False
      主键InputBox0008=   0   'False
      允许等于最大值InputBox0008=   0   'False
      允许等于最小值InputBox0008=   0   'False
      字典名称InputBox0008=   ""
      显示字典字段InputBox0008=   ""
      保存字典字段InputBox0008=   ""
      名称InputBox0008=   "诊断和处理意见"
      缺省值InputBox0008=   ""
      保存缺省值InputBox0008=   ""
      长度InputBox0008=   0
      MaxInputBox0008 =   ""
      MinInputBox0008 =   ""
      VisibleInputBox0008=   -1  'True
      PermitNullInputBox0008=   -1  'True
      TriggerstrInputBox0008=   ""
      EnableInputBox0008=   0   'False
      允许多选InputBox0008=   0   'False
      titleInputBox0009=   "下结论医师"
      statusinfoInputBox0009=   ""
      lengthInputBox0009=   18
      orderInputBox0009=   9
      valueInputBox0009=   ""
      datatypeInputBox0009=   3
      colInputBox0009 =   32
      rowInputBox0009 =   3
      PassWordCharInputBox0009=   0   'False
      主键InputBox0009=   0   'False
      允许等于最大值InputBox0009=   0   'False
      允许等于最小值InputBox0009=   0   'False
      字典名称InputBox0009=   ""
      显示字典字段InputBox0009=   ""
      保存字典字段InputBox0009=   ""
      名称InputBox0009=   "下结论医师"
      缺省值InputBox0009=   ""
      保存缺省值InputBox0009=   ""
      长度InputBox0009=   0
      MaxInputBox0009 =   ""
      MinInputBox0009 =   ""
      VisibleInputBox0009=   -1  'True
      PermitNullInputBox0009=   -1  'True
      TriggerstrInputBox0009=   ""
      EnableInputBox0009=   0   'False
      允许多选InputBox0009=   0   'False
      ErrColor        =   16777215
   End
   Begin VB.Timer ctmMain 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   10200
      Top             =   3840
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "返  回"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   1
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "打  印"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   0
      Top             =   5880
      Width           =   2055
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdResult 
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   9615
      _cx             =   4211264
      _cy             =   4204279
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   49152
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648384
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
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "常规项目|^常规结果|化验项目|^化验结果"
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
   End
   Begin VB.PictureBox cpcMain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   9960
      ScaleHeight     =   2265
      ScaleWidth      =   1665
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检查询结果"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   0
      Left            =   3495
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'作者：杨春

Public 系统编号 As String
Private mlngCount As Long

Private Sub ccmdExit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub ccmdPrint_Click()
    Dim lcol编号 As Collection
    On Error GoTo errHandler
    '打印体检结果单。
    Set lcol编号 = New Collection
    lcol编号.Add 系统编号
    pobj业务对象.Sub打印文书 "体检结果单", lcol编号, False
    
errHandler:
    
End Sub

Private Sub ctmMain_Timer()
    On Error Resume Next
    '计时。
    mlngCount = mlngCount + 1
    
    '两分钟自动关闭本窗口。
    If mlngCount = 2 Then
        ctmMain.Enabled = False
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Dim lobj体检 As Object '体检对象。
    Dim lobjItem As Object '体检项目及其结果clsFactTestItem。
    Dim lcolInfo As Collection
    Dim lngWidth As Long
    Dim i As Long
    
    On Error GoTo errHandler
    mlngCount = 0
    
    '创建体检对象。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
        
    '获取系统编号内容。
    lobj体检.系统编号 = 系统编号
    If Not lobj体检.是否已存在 Then
        Err.Raise 6666, , "该系统编号的体检记录不存在。请重新输入体检的系统编号（或刷条码）。"
    End If
    If lobj体检.体检状态 <> P_ENDED_STATUS Then
        Err.Raise 6666, , "该系统编号的体检记录还未下体检结论，不允许查询。"
    End If
    
    '显示体检基本信息。
    With lobj体检
        c显示板.Box2("姓名").Text = .体检人员.姓名
        c显示板.Box2("性别").Text = .体检人员.性别
        c显示板.Box2("年龄").Text = .体检人员.年龄
        c显示板.Box2("单位名称").Text = .体检人员.单位名称
        
        Select Case .体检类别
        Case 0
            c显示板.Box2("体检类别").Text = "初检"
        Case 1
            c显示板.Box2("体检类别").Text = "复查"
        Case 2
            c显示板.Box2("体检类别").Text = "年检"
        End Select
        c显示板.Box2("体检日期").Text = .体检日期
        c显示板.Box2("体检结论").Text = .体检结论
        c显示板.Box2("诊断和处理意见").Text = .诊断和处理意见
        c显示板.Box2("下结论医师").Text = .下结论医师姓名
        
        If .体检人员.像片 Is Nothing Then
            cpcMain.Picture = Nothing
        Else
            cpcMain.Picture = .体检人员.像片
        End If
        
    End With
    
    cgrdResult.Redraw = False
    
    '获取常规体检项目并显示当前体检人员的体检结果。
    Set lcolInfo = lobj体检.体检表.体检项目集("常规")
    If lcolInfo.Count > 0 Then
        cgrdResult.Rows = lcolInfo.Count + 1
    Else
        cgrdResult.Rows = 1
    End If
    i = 1
    For Each lobjItem In lcolInfo
        cgrdResult.TextMatrix(i, 0) = lobjItem.体检项目名称
        cgrdResult.TextMatrix(i, 1) = lobjItem.体检结果
        i = i + 1
    Next
    '获取化验项目并显示当前体检人员的体检结果。
    Set lcolInfo = lobj体检.体检表.体检项目集("化验")
    i = 1
    For Each lobjItem In lcolInfo
        If i = cgrdResult.Rows Then
            cgrdResult.Rows = cgrdResult.Rows + 1
        End If
        cgrdResult.TextMatrix(i, 2) = lobjItem.体检项目名称
        cgrdResult.TextMatrix(i, 3) = lobjItem.体检结果
        i = i + 1
    Next
    Set lobj体检 = Nothing
    Set lobjItem = Nothing
    
    '刷新体检结果网格。
    cgrdResult.Redraw = True


    lngWidth = (cgrdResult.Width - 500) / 4
    cgrdResult.ColWidth(0) = lngWidth + 500
    cgrdResult.ColWidth(1) = lngWidth - 500
    cgrdResult.ColWidth(2) = lngWidth + 500
    cgrdResult.ColWidth(3) = lngWidth - 500
    
    ctmMain.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "frmTestResult", "Form_Load", 6666, lstrError, False
    Unload Me
    
    Exit Sub
    Resume
End Sub


