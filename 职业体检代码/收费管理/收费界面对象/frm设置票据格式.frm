VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm设置票据格式 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置票据格式"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10800
   ClipControls    =   0   'False
   Icon            =   "frm设置票据格式.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   10815
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         Caption         =   "票据信息"
         ForeColor       =   &H80000008&
         Height          =   5130
         Left            =   6720
         TabIndex        =   10
         Top             =   120
         Width           =   3972
         Begin VB.OptionButton copt汇总 
            Caption         =   "否"
            Height          =   252
            Index           =   1
            Left            =   1680
            TabIndex        =   19
            Top             =   3360
            Value           =   -1  'True
            Width           =   492
         End
         Begin VB.OptionButton copt汇总 
            Caption         =   "是"
            Height          =   252
            Index           =   0
            Left            =   960
            TabIndex        =   18
            Top             =   3360
            Width           =   612
         End
         Begin VB.CommandButton Ccmd浏览 
            Caption         =   "浏览…"
            Height          =   396
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3720
            Width           =   1125
         End
         Begin VB.ComboBox ccmb对应业务 
            Height          =   276
            ItemData        =   "frm设置票据格式.frx":0442
            Left            =   960
            List            =   "frm设置票据格式.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2100
            Width           =   2868
         End
         Begin VB.TextBox cinb票据格式文件名称 
            Height          =   372
            Left            =   120
            MaxLength       =   24
            TabIndex        =   6
            Top             =   4200
            Width           =   3708
         End
         Begin VB.TextBox cinb票据名称 
            Height          =   360
            Left            =   960
            MaxLength       =   25
            TabIndex        =   1
            Top             =   960
            Width           =   2832
         End
         Begin VB.TextBox cinb票据编号 
            Enabled         =   0   'False
            Height          =   360
            Left            =   960
            MaxLength       =   2
            TabIndex        =   0
            Top             =   360
            Width           =   1290
         End
         Begin VB.ComboBox ccmb票据类型 
            Height          =   276
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1560
            Width           =   2868
         End
         Begin VB.TextBox ctxt最大项数 
            Height          =   360
            Left            =   960
            MaxLength       =   2
            TabIndex        =   4
            Top             =   2640
            Width           =   1245
         End
         Begin MSComDlg.CommonDialog Ccmn票据 
            Left            =   3720
            Top             =   4680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "汇总项目"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   3360
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票据类型"
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "对应业务"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   2100
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "格式文件名"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   3960
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "编号"
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票据名称"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "最大项数"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   2760
            Width           =   720
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
         Height          =   5115
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   6465
         _cx             =   165162124
         _cy             =   165159742
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   7
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
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb设置 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
End
Attribute VB_Name = "frm设置票据格式"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

Private Sub Ccmd浏览_Click()
    On Error GoTo errHandler
    Dim lstrFileName As String
    Ccmn票据.InitDir = App.Path
    Ccmn票据.ShowOpen
    lstrFileName = funcGetFileName(Ccmn票据.filename)
    If Len(lstrFileName) > 14 Then
        sffuncMsg "对不起，文件名称过长，必须少于14个字！", sf警告
    Else
        cinb票据格式文件名称.Text = lstrFileName
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm设置票据格式", "Ccmd浏览_Click", Err.Number, Err.Description, False
    
End Sub

Private Function funcGetFileName(filename As String) As String
    Dim lintOffset As Integer
    Dim lstrResult As String
    
    lstrResult = filename
    lintOffset = InStr(lstrResult, "\")
    If lintOffset = 0 Then
        funcGetFileName = filename
        Exit Function
    End If
    
    Do While lintOffset <> 0
        lstrResult = Mid(lstrResult, lintOffset + 1)
        lintOffset = InStr(lstrResult, "\")
    Loop
    funcGetFileName = lstrResult
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        SendKeys Chr(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True

    '初始化工具栏
    Dim lcol工具栏按钮 As Collection
    Set lcol工具栏按钮 = New Collection
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.c工具栏 = ctlb设置

    lcol工具栏按钮.Add "添加"
    lcol工具栏按钮.Add "删除"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "保存"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select 编号,票据类型编号,票据名称,票据格式文件名称,对应业务,最大项数,项目汇总 from 收费管理_票据设置信息表 order by 编号")
    Set cgrdMain.DataSource = lobjRec
    cgrdMain.Row = 0
    cgrdMain.ColHidden(1) = True
    
    '初始化票据类型
    Set lobjRec = dafuncGetData("select * from 收费管理_票据类型字典视图")
    If (Not lobjRec.EOF) And (Not lobjRec.BOF) Then
        Do While (Not lobjRec.EOF)
            ccmb票据类型.AddItem lobjRec.Fields("名称").Value
            ccmb票据类型.ItemData(ccmb票据类型.NewIndex) = lobjRec.Fields("innerId").Value
            lobjRec.MoveNext
        Loop
        lobjRec.MoveFirst
    End If
    If ccmb票据类型.ListCount > 0 Then
        ccmb票据类型.ListIndex = 0
    End If
    
    
    ccmb对应业务.AddItem "一般", 0
    ccmb对应业务.AddItem "门诊", 1
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm设置票据格式", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Select Case Operate
    Case "添加"
        Cancel = True
        cinb票据编号.Text = ""
        cinb票据名称.Text = ""
        ctxt最大项数.Text = 4
        cinb票据格式文件名称.Text = ""
        cinb票据名称.SetFocus
    
    Case "保存"
        Cancel = True
        subValidate
         
        pobj收费管理.sub保存票据设置 cinb票据编号.Text, cinb票据名称.Text, cinb票据格式文件名称.Text, ccmb票据类型.ItemData(ccmb票据类型.ListIndex), ccmb对应业务.Text, Val(ctxt最大项数.Text), IIf(copt汇总(0).Value, "是", "否")
        
        Dim lobjRec As Object
        Set lobjRec = dafuncGetData("select 编号,票据类型编号,票据名称,票据格式文件名称,对应业务,最大项数,项目汇总 from 收费管理_票据设置信息表 order by 编号")
        Set cgrdMain.DataSource = lobjRec
        cgrdMain.Row = 0
        cgrdMain.ColHidden(1) = True
        mobjGUI_BeforeOperate "添加", True
        
    Case "删除"
        Cancel = True
    End Select
    
    Set lobjRec = Nothing
    Exit Sub
    
errHandler:
    sfsub错误处理 "收费界面部件", "frm设置票据格式", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
        
End Sub

Private Sub subValidate()
    '判断用户录入是否合理
    If cinb票据格式文件名称.Text = "" Or cinb票据名称.Text = "" Or ccmb对应业务.Text = "" Or ccmb票据类型.Text = "" Then
        Err.Raise 6666, , "录入不完整，请重新录入！"
    End If
    If ctxt最大项数.Text = "" Then
        Err.Raise 6666, , "请填上最大项数！"
    End If
    If Not IsNumeric(ctxt最大项数.Text) Then
        If ctxt最大项数.Text < 1 Then
            Err.Raise 6666, , "最大项数录入非法！"
        End If
    End If
    
    '判断该记录是否存在
    Dim linttemp As Integer
    If cgrdMain.Rows > 0 Then
        For linttemp = 1 To cgrdMain.Rows - 1
            If linttemp <> cgrdMain.Row Then
                If cgrdMain.Cell(flexcpText, linttemp, 1) = ccmb票据类型.ItemData(ccmb票据类型.ListIndex) And _
                    cgrdMain.Cell(flexcpText, linttemp, 3) = cinb票据格式文件名称.Text And _
                    cgrdMain.Cell(flexcpText, linttemp, 4) = ccmb对应业务.Text And _
                    cgrdMain.Cell(flexcpText, linttemp, 5) = ctxt最大项数.Text Then
                    Err.Raise 6666, , "该票据设置信息已存在，请修改！"
                End If
            End If
        Next
    End If
End Sub



Private Sub cgrdMain_Click()
    Dim lobjRec As Object
    Dim lstr编号 As String
    
    On Error GoTo errHandler
    
    If cgrdMain.RowSel = 0 Then
        Exit Sub
    End If
    Set lobjRec = dafuncGetData("select * from 收费管理_票据类型字典视图")
    cinb票据编号.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 0)
    cinb票据名称.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 2)
    cinb票据格式文件名称.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 3)
    ctxt最大项数.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 5)
    If (Not lobjRec.EOF) And (Not lobjRec.BOF) Then
        lobjRec.MoveFirst
        Do While (Not lobjRec.EOF)
            lstr编号 = lobjRec("innerID").Value
            If lstr编号 = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 1) Then
               ccmb票据类型.Text = lobjRec("名称")
               Exit Do
            Else
                lobjRec.MoveNext
            End If
            If lobjRec.EOF Then
                MsgBox "如果修改了票据类型字典表，请重新配置该项目的票据类型，该项目票据类型有问题！", vbExclamation, "系统提示"
                Exit Do
            End If
       Loop
    End If
    
    If cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 4) = "一般" Then
        ccmb对应业务.ListIndex = 0
    Else
        ccmb对应业务.ListIndex = 1
    End If
    
    If cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 6) = "是" Then
        copt汇总(0).Value = True
    Else
        copt汇总(1).Value = True
    End If
    
    cinb票据名称.SetFocus
    
    Set lobjRec = Nothing
    Exit Sub
    
errHandler:
    sfsub错误处理 "收费界面部件", "frm设置票据格式", "cgrdMain_Click", Err.Number, Err.Description, False
    
End Sub


