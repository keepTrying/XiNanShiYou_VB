VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Begin VB.Form frm设置打折 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置打折"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9735
   ClipControls    =   0   'False
   Icon            =   "frm设置打折.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   9615
      Begin VB.Frame Frame8 
         Caption         =   "打折控制项"
         Height          =   1755
         Left            =   4680
         TabIndex        =   9
         Top             =   4080
         Width           =   4680
         Begin VB.OptionButton copt打折 
            Caption         =   "可以打折，但严格控制"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   1185
            Width           =   2100
         End
         Begin VB.OptionButton copt打折 
            Caption         =   "可以打折，但不严格控制"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   735
            Width           =   2280
         End
         Begin VB.OptionButton copt打折 
            Caption         =   "不打折"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   300
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CommandButton ccmd单位定位 
         Caption         =   "单位定位"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox ctxt打折 
         Height          =   360
         Left            =   6000
         TabIndex        =   1
         Text            =   "1.00"
         Top             =   960
         Width           =   1065
      End
      Begin VB.VScrollBar cvsl打折 
         Height          =   360
         Left            =   6600
         Max             =   1
         Min             =   100
         TabIndex        =   2
         Top             =   960
         Value           =   100
         Width           =   675
      End
      Begin 录入控件.ctlInputBox cinp交费单位 
         Height          =   360
         Left            =   4680
         TabIndex        =   6
         Top             =   360
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   760
         Text            =   ""
         Label           =   "交费单位"
         Enabled         =   0   'False
         名称            =   ""
         长度            =   0
         允许等于最大值  =   0   'False
         允许等于最小值  =   0   'False
         允许多选        =   0   'False
      End
      Begin VSFlex6Ctl.vsFlexGrid cFlg打折 
         Height          =   5685
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   4395
         _cx             =   20192840
         _cy             =   20195116
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
         BackColorAlternate=   14737632
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
         Cols            =   4
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
         DataMember      =   ""
      End
      Begin 录入控件.ctlInputBox cinb单位编号 
         Height          =   360
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   760
         Text            =   ""
         Label           =   "单位编号"
         Enabled         =   0   'False
         名称            =   ""
         长度            =   0
         允许等于最大值  =   0   'False
         允许等于最小值  =   0   'False
         允许多选        =   0   'False
         BackgroundColor =   15791081
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打折比率"
         Height          =   180
         Left            =   4680
         TabIndex        =   12
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label clbl打折说明 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   4560
         TabIndex        =   11
         Top             =   1800
         Width           =   4965
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
      TabIndex        =   13
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
End
Attribute VB_Name = "frm设置打折"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

Private Sub ccmd单位定位_Click()
    On Error GoTo errhandler
    
    Dim lobj单位信息 As Object
    Set lobj单位信息 = pobj单位定位.func单位简单定位(8600, 1000)
    If Not (lobj单位信息 Is Nothing) Then   '已获取单位信息
        cinp交费单位.Text = lobj单位信息.Fields!单位名称
        cinb单位编号.Text = IIf(IsNull(lobj单位信息.Fields!申请编号), "", lobj单位信息.Fields!申请编号)
    Else
    
    End If

    ctxt打折.SetFocus
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置打折", "ccmd单位定位_Click", Err.Number, Err.Description, False
End Sub

Private Sub cFlg打折_Click()
    On Error Resume Next
    cinb单位编号.Text = cFlg打折.TextMatrix(cFlg打折.Row, 1)
    cinp交费单位.Text = cFlg打折.TextMatrix(cFlg打折.Row, 2)
    ctxt打折.Text = cFlg打折.TextMatrix(cFlg打折.Row, 3)
    cvsl打折.Value = Int(Val(cFlg打折.TextMatrix(cFlg打折.Row, 3)) * 100)
End Sub

Private Sub cvsl打折_Change()
    
    ctxt打折.Text = cvsl打折.Value / 100
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    
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
    
    Dim lint打折控制  As Integer
    lint打折控制 = Val(pobj收费管理.业务设置("打折控制"))
    copt打折(lint打折控制).Value = True
    
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select a.单位编号,b.单位名称,a.打折比率 from 收费管理_打折信息表 a inner join 单位档案_单位基本信息表 b on a.单位编号=b.申请编号")
    Set cFlg打折.DataSource = lobjRec
    
    cFlg打折.Row = 0
    
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置打折", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Select Case Operate
    Case "添加"
        Cancel = True
        ctxt打折.Enabled = True
        ccmd单位定位.Enabled = True
        cvsl打折.Enabled = True
        
        cinb单位编号.Text = ""
        cinp交费单位.Text = ""
        ctxt打折.Text = "1.00"
        
   Case "删除"
        Cancel = True
        If cinb单位编号.Text <> "" Then
            pobj收费管理.sub删除打折信息 LTrim(RTrim(cinb单位编号.Text))
            cFlg打折.RemoveItem cFlg打折.RowSel
        Else
            MsgBox "请选择具体的单位！", vbExclamation, "系统管理"
        End If
   Case "保存"
        Dim lint打折控制 As Integer
        Cancel = True
        If copt打折(0).Value Then
            lint打折控制 = 1
        ElseIf copt打折(1).Value Then
            lint打折控制 = 1
        Else
            lint打折控制 = 1
        End If
        pobj收费管理.业务设置("打折控制") = lint打折控制
        
        If cinb单位编号.Text <> "" Then
            pobj收费管理.sub保存打折信息 cinb单位编号.Text, IIf(IsNull(ctxt打折.Text), "1", ctxt打折.Text)
            '刷新网格。
            Dim lobjRec As Object
            Set lobjRec = dafuncGetData("select a.单位编号,b.单位名称,a.打折比率 from 收费管理_打折信息表 a inner join 单位档案_单位基本信息表 b on a.单位编号=b.申请编号")
            Set cFlg打折.DataSource = lobjRec
            
            cinb单位编号.Text = ""
            cinp交费单位.Text = ""
            ccmd单位定位.SetFocus
            
        End If
        
   End Select
   
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置打折", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume
   
End Sub
