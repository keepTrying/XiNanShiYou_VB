VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frm号段设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发票号段设置"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退  出"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox clstName 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox ctxtBegin 
      Height          =   270
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox ctxtEnd 
      Height          =   270
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton ccmdSave 
      Caption         =   "保  存"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取  消"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   5895
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "双击所选定的号段信息可以进行修改"
      Top             =   1080
      Width           =   3375
      _cx             =   52172609
      _cy             =   52177054
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
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "ID|起号|止号|是否用完"
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
   Begin VB.Label clblCurNo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5280
      TabIndex        =   12
      Top             =   4680
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "当前票号："
      Height          =   180
      Left            =   4200
      TabIndex        =   11
      Top             =   4680
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收费员："
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "已设置号段："
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "起号："
      Height          =   180
      Left            =   4560
      TabIndex        =   7
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "止号："
      Height          =   180
      Left            =   4560
      TabIndex        =   6
      Top             =   1800
      Width           =   540
   End
End
Attribute VB_Name = "frm号段设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Dim mlngID As Long      '当前修改的号段的ID

Private Sub ccmdCancel_Click()
    mlngID = 0
    ctxtBegin = ""
    ctxtEnd = ""
End Sub

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdSave_Click()
    Dim i As Integer
    Dim llngBegin As Long, llngEnd As Long
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    
    ctxtBegin = Trim(ctxtBegin)
    ctxtEnd = Trim(ctxtEnd)
    If ctxtBegin = "" Then
        MsgBox "起号不能为空！", vbInformation, "系统提示"
        ctxtBegin.SetFocus
        Exit Sub
    End If
    If ctxtEnd = "" Then
        MsgBox "止号不能为空！", vbInformation, "系统提示"
        ctxtEnd.SetFocus
        Exit Sub
    End If
    llngBegin = CLng(ctxtBegin)
    llngEnd = CLng(ctxtEnd)
    If CLng(ctxtEnd) < llngBegin Then
        MsgBox "止号必须小于起号！", vbInformation, "系统提示"
        ctxtEnd.SetFocus
        Exit Sub
    End If
    '检查该范围是否已在系统录入的号段范围内
    Set lobjRec = dafuncGetData("select * from 收费管理_收费员号段信息表 where (起号<='" & llngBegin & "' and 止号>='" & llngBegin & "' or 起号<='" & llngEnd & "' and 止号>='" & llngEnd & "' or 起号>'" & llngBegin & "' and 止号<'" & llngEnd & "') and ID<>" & mlngID)
    If lobjRec.RecordCount Then
        MsgBox "所添加的号段范围与已分配给该收费员或其他收费员的号段范围重叠，不能添加！", vbInformation, "系统提示"
        ctxtBegin.SetFocus
        Exit Sub
    End If
    If mlngID = 0 Then
        dafuncGetData "insert into 收费管理_收费员号段信息表 (用户编号,起号,止号,是否用完) values('" & Mid(clstName.Text, InStr(clstName.Text, " ") + 1) & "','" & llngBegin & "','" & llngEnd & "','否')"
    Else
        dafuncGetData "update 收费管理_收费员号段信息表 set 起号='" & llngBegin & "',止号='" & llngEnd & "' where ID=" & mlngID
    End If
    MsgBox "保存成功！", vbInformation, "系统提示"
    ctxtBegin = ""
    ctxtEnd = ""
    mlngID = 0
    '刷新表格
    clstName_Click
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm业务设置", "ccmdSave_Click", Err.Number, Err.Description, False
    Exit Sub
End Sub

Private Sub cgrdMain_DblClick()
    If cgrdMain.Row = 0 Then Exit Sub
    If cgrdMain.Cell(flexcpText, cgrdMain.Row, 3) = "是" Then
        MsgBox "该号段已经使用完毕，不能修改！", vbInformation, "系统提示"
    Else
        mlngID = CLng(cgrdMain.Cell(flexcpText, cgrdMain.Row, 0))
        ctxtBegin = cgrdMain.Cell(flexcpText, cgrdMain.Row, 1)
        ctxtEnd = cgrdMain.Cell(flexcpText, cgrdMain.Row, 2)
    End If
End Sub

Private Sub clstName_Click()
    Dim lobjRec As Object, i As Integer

    On Error GoTo errhandler
    
    cgrdMain.FormatString = ""
    Set lobjRec = dafuncGetData("select ID,起号,止号,是否用完 from 收费管理_收费员号段信息表 where 用户编号='" & Mid(clstName.Text, InStr(clstName.Text, " ") + 1) & "' order by ID desc")
    Set cgrdMain.DataSource = lobjRec
    cgrdMain.ColHidden(0) = True
        
        
    Set lobjRec = dafuncGetData("select 当前值 from 系统管理_系统编号生成记录表 where 业务名称='收费管理" & Mid(clstName.Text, InStr(clstName.Text, " ") + 1) & "' and 编号名称='收据号'")
    If lobjRec.RecordCount = 0 Then
        clblCurNo = ""
    Else
        clblCurNo = lobjRec(0)
    End If
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm业务设置", "clstName_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
Private Sub ctxtBegin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ctxtEnd.SetFocus
End Sub

Private Sub ctxtBegin_LostFocus()
    ctxtBegin = Trim(ctxtBegin)
    If ctxtBegin <> "" Then
        If Not IsNumeric(ctxtBegin) Then
            MsgBox "起号必须输入正确的数字！", vbInformation, "系统提示"
            ctxtBegin.SetFocus
        ElseIf CLng(ctxtBegin) <= 0 Then
            MsgBox "起号必须是大于0的整数！", vbInformation, "系统提示"
            ctxtBegin.SetFocus
        ElseIf ctxtEnd <> "" Then
            If CLng(ctxtEnd) < CLng(ctxtBegin) Then MsgBox "止号必须小于起号！", vbInformation, "系统提示"
        End If
    End If
End Sub

Private Sub ctxtEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ccmdSave.SetFocus
End Sub

Private Sub ctxtEnd_LostFocus()
    ctxtEnd = Trim(ctxtEnd)
    If ctxtEnd <> "" Then
        If Not IsNumeric(ctxtEnd) Then
            MsgBox "止号必须输入正确的数字！", vbInformation, "系统提示"
            ctxtEnd.SetFocus
        ElseIf CLng(ctxtEnd) <= 0 Then
            MsgBox "止号必须是大于0的整数！", vbInformation, "系统提示"
            ctxtEnd.SetFocus
        ElseIf ctxtBegin <> "" Then
            If CLng(ctxtEnd) < CLng(ctxtBegin) Then MsgBox "止号必须小于起号！", vbInformation, "系统提示"
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
    
    '获取当前票据号。
    Dim lobjRec As Object, i As Integer
    Dim lobjRec1 As Object
    
    mlngID = 0
    clstName.Clear
    Set lobjRec = dafuncGetData("select 编号,姓名 from 系统管理_员工基本信息视图 order by 编号")
    For i = 1 To lobjRec.RecordCount
        Set lobjRec1 = dafuncGetData("select * from 系统管理_用户操作权限表 where 用户编号='" & lobjRec(0) & "' and 权限名='收费管理_直接收费'")
        If lobjRec1.RecordCount > 0 Then
            clstName.AddItem lobjRec(1) & " " & lobjRec(0)
        End If
        lobjRec.MoveNext
    Next
    If clstName.ListCount > 0 Then
        clstName.ListIndex = 0
    Else
        MsgBox "当前没有设置具有收费权限的人员，不能设置票据号段，请先为收费员分配收费权限！", vbInformation, "系统提示"
    End If
    
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm业务设置", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub
