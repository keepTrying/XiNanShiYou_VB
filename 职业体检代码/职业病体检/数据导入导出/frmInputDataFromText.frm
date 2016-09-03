VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frmInputData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "体检结果导入"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputDataFromText.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11355
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ccmdView 
      Caption         =   "预览(&V)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton ccmdImport 
      Caption         =   "导入(&I)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "返回(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   9600
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "浏览文件(&B)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   1275
   End
   Begin VB.Frame cfraPreview 
      Caption         =   "预览数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7395
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   11175
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   6915
         Left            =   7200
         TabIndex        =   7
         Top             =   360
         Width           =   3855
         _cx             =   6800
         _cy             =   12197
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
         BackColorAlternate=   13827279
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
         MultiTotals     =   0   'False
         SubtotalPosition=   1
         OutlineBar      =   1
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   0   'False
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdPreview 
         Height          =   6915
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   6975
         _cx             =   12303
         _cy             =   12197
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
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   1
         MultiTotals     =   0   'False
         SubtotalPosition=   1
         OutlineBar      =   1
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   0   'False
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
   Begin VB.TextBox ctxtDataSource 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4035
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   6840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label clabDataSource 
      AutoSize        =   -1  'True
      Caption         =   "数据来源："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmInputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mobj对外接口 As ClsManageTransmission
Public pblnInUse As Boolean

Private mstrID As String
Private mcolIndex As New Collection
Private Sub ccmdExit_Click()
    Unload Me
End Sub


Private Sub ccmdImport_Click()
    On Error GoTo errHandler
    
    If ctxtDataSource = "" Then
        MsgBox "请输入或选择要导入的文件！", vbOKOnly + vbExclamation, "系统提示"
        Exit Sub
    End If
    If mstrID = "" Then
        mstrID = mobj对外接口.func读入体检结果(ctxtDataSource)
    End If
    '写库。
    dafuncGetData "update 体检管理_体检结果信息表 set 体检管理_体检结果信息表.体检结果=b.体检结果 from 体检管理_体检结果信息表 a,temp_体检结果信息 b,体检管理_体检项目设置表 c  where a.系统编号=b.系统编号 and a.体检项目=c.编码 and b.体检项目名称=c.名称  and b.ID='" & mstrID & "'"
    
    MsgBox "导入成功！", vbOKOnly + vbExclamation, "系统提示"
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检数据导入导出", "frmInputData", "ccmdImport_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub ccmdView_Click()
    Dim lobjRec As Object
    Dim i As Long
    On Error GoTo errHandler
    
    If ctxtDataSource = "" Then
        MsgBox "请输入或选择要导入的文件！", vbOKOnly + vbExclamation, "系统提示"
        Exit Sub
    End If
    
    If mstrID <> "" Then
        dafuncGetData "delete temp_体检基本信息 where ID='" & mstrID & "'"
    End If
    
    cgrdPreview.Rows = 1
    cgrdDetail.Rows = 1
    
    mstrID = mobj对外接口.func读入体检结果(ctxtDataSource)
    Set lobjRec = dafuncGetData("select * from temp_体检基本信息 where ID='" & mstrID & "' order by 系统编号")
    Set cgrdPreview.DataSource = lobjRec
    cgrdPreview.ColHidden(cgrdPreview.Cols - 1) = True
    
    Set mcolIndex = New Collection
    For i = 0 To cgrdPreview.Cols - 1
        mcolIndex.Add i, cgrdPreview.TextMatrix(0, i)
    Next
    Exit Sub
errHandler:
    sfsub错误处理 "体检数据导入导出", "frmInputData", "ccmdView_Click", Err.Number, Err.Description, False
End Sub

Private Sub cgrdPreview_Click()
    Dim lobjRec As Object
    cgrdDetail.Rows = 1
    If cgrdPreview.Row > 0 Then
        Set lobjRec = dafuncGetData("select 体检项目名称,体检结果 from temp_体检结果信息 where ID='" & mstrID & "' and 系统编号='" & cgrdPreview.TextMatrix(cgrdPreview.Row, mcolIndex("系统编号")) & "'")
        Set cgrdDetail.DataSource = lobjRec
        cgrdDetail.AutoSize 0, cgrdDetail.Cols - 1
    End If
End Sub

'初始化界面,
Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection
    Dim lcolTemplateSet As Object
    Dim lcolInfo As Collection
    
    Dim i As Integer
    
    On Error GoTo errHandler
    
    Set mobj对外接口 = New ClsManageTransmission
    

    Exit Sub
errHandler:
    sfsub错误处理 "体检数据导入导出", "frmInputData", "Form_Load", Err.Number, Err.Description, False
End Sub

'功能: 弹出文件查找窗口.
Private Sub ccmdBrowse_Click() ' 设置“CancelError”为 True
    ccdgBrowse.CancelError = True
    On Error GoTo errHandler
    ' 设置标志
    ccdgBrowse.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    ccdgBrowse.Filter = "All Files (*.*)|*.*|文本文件" & _
        "(*.txt)|*.txt"
    ccdgBrowse.FilterIndex = 2
    ccdgBrowse.ShowOpen
    ctxtDataSource.Text = ccdgBrowse.FileName

    If ctxtDataSource.Text <> "" Then
        ccmdView_Click
    End If
    Exit Sub
errHandler:
    ' 用户按了“取消”按钮
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If mstrID <> "" Then
        dafuncGetData "delete temp_体检基本信息 where ID='" & mstrID & "'"
    End If
    
    Set mobj对外接口 = Nothing

End Sub


