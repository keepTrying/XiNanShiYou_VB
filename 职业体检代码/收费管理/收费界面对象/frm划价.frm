VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm划价 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "划价"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm划价.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "费用清修改单 "
      ForeColor       =   &H80000008&
      Height          =   6120
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   10860
      Begin VB.Frame Frame1 
         Height          =   5500
         Left            =   6360
         TabIndex        =   19
         Top             =   120
         Width           =   4335
         Begin VB.OptionButton coptFind 
            Caption         =   "按助记符查找"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   5040
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.TextBox ctxtFind 
            Height          =   375
            Left            =   1680
            TabIndex        =   24
            Top             =   5040
            Width           =   1935
         End
         Begin VB.ComboBox ccmb收费标准 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox Ccbo收费项目大类 
            Height          =   300
            Left            =   1440
            TabIndex        =   6
            Top             =   600
            Width           =   2655
         End
         Begin VSFlex6Ctl.vsFlexGrid cgrdItem 
            Height          =   3975
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   4000
            _cx             =   60955536
            _cy             =   60955491
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
            Rows            =   50
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   "收费项目            |编号         |助记符     "
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
         Begin VB.Label Clab收费项目大类 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费项目大类"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费标准"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   720
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   5340
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6165
         _cx             =   115354234
         _cy             =   115352779
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
         TabBehavior     =   1
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
      Begin VB.Label clblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总金额："
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   5640
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "按“Del”键可以删除当前选中的项目"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   2970
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "基本信息"
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   10845
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   200
         Width           =   1665
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   0
         Top             =   200
         Width           =   1560
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   3
         Left            =   6360
         TabIndex        =   1
         Top             =   200
         Width           =   2715
      End
      Begin VB.ComboBox ccmb主管科室 
         Height          =   300
         Left            =   6360
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox ccmb卫生种类 
         Height          =   300
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox ccmb片区 
         Height          =   300
         Left            =   3600
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton ccmd定位 
         Caption         =   "..."
         Height          =   255
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费编号"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   15
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费人"
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   225
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费单位"
         Height          =   180
         Index           =   3
         Left            =   5520
         TabIndex        =   13
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主管科室"
         Height          =   180
         Index           =   4
         Left            =   5520
         TabIndex        =   12
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卫生种类"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "片区"
         Height          =   180
         Left            =   3000
         TabIndex        =   10
         Top             =   720
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   9840
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
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   979
      ButtonWidth     =   1455
      ButtonHeight    =   926
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchk清空 
         Caption         =   "保存后清空"
         Height          =   255
         Left            =   8520
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm划价"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

'启动参数。
Public pstr收费编号 As String
Public pstr单位编号 As String
Public pstr业务分类 As String

Dim WithEvents mobj界面通用对象 As cls界面通用对象
Attribute mobj界面通用对象.VB_VarHelpID = -1


Private Const 收费_收费编号 = 1
Private Const 收费_交费人 = 2
Private Const 收费_交费单位 = 3

Private Const 费用清单_收费项目编号 = 0
Private Const 费用清单_收费项目名称 = 1
Private Const 费用清单_单价 = 2
Private Const 费用清单_数量 = 3
Private Const 费用清单_金额 = 4


Dim mstrUndoCount As String          '用于保存表格中原来的字符串,以便在输入不合法时能够还原
Dim mstrUndoMoney As String          '用于保存表格中原来的字符串,以便在输入不合法时能够还原
Dim mcur最小单价 As Currency
Dim mcur最大单价 As Currency

Dim mstr交费单位编号 As String  '从单位定位接口得到的交费单位的编号
Dim mint交费方式编号 As Integer '交费方式的编号

Dim mint科目级数 As Integer

Private Sub Ccbo收费项目大类_Click()
    On Error GoTo errhandler
   
    Dim lobjRec As Object            '定义变量记录数据集
    
    '根据收费项目大类名称,获取收费编号前缀
    Set lobjRec = dafuncGetData("select 收费项目编号 from 收费管理_收费项目字典表 where 收费项目名称= '" & Ccbo收费项目大类.Text & "'")
    
    '获取下级收费项目
    Set lobjRec = dafuncGetData("select 收费项目名称 as 收费项目,收费项目编号 as 编号,助记符 from 收费管理_收费项目字典表 where left(收费项目编号,3)='" & Left$(lobjRec("收费项目编号"), 3) & "' and len(收费项目编号)>3")
    
    cgrdItem.FormatString = ""
    Set cgrdItem.DataSource = lobjRec
    cgrdItem.ColWidth(0) = 2000
    cgrdItem.Row = 0
    On Error Resume Next
    ctxtFind.SetFocus
    ctxtFind.Text = ""
    Exit Sub
errhandler:
    MsgBox "获取并显示指定大类的收费项目失败！" & Error, vbOKOnly + vbExclamation, "系统提示"
End Sub

Private Sub ccmb收费标准_Click()
    Dim lrds收费标准 As Object
    Dim i As Integer
    Dim lcurMoney As Currency
    
    On Error GoTo errhandler
    
    Set lrds收费标准 = dafuncGetData("select a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,金额=a.单价*a.数量 from 收费管理_收费标准信息表 a,收费管理_收费项目字典表 b where b.收费项目编号=a.收费项目编号 and 收费标准名称='" & ccmb收费标准.Text & "'")
    
    If lrds收费标准.EOF Then
        sffuncMsg "收费标准中无收费项目！", sf警告
        Exit Sub
    Else
        lrds收费标准.MoveFirst
        Dim llngItemCount As Long
        For i = 0 To lrds收费标准.RecordCount - 1
            If Not func检查项目是否已选(lrds收费标准("收费项目编号")) Then
                sub添加项目 lrds收费标准("收费项目编号")
                llngItemCount = llngItemCount + 1
            End If
            lrds收费标准.MoveNext
        Next
        
        sub计算总金额
        
        If llngItemCount = lrds收费标准.RecordCount Then
            MsgBox "收费标准中的所有收费项目(" & llngItemCount & "条)已添加到费用清单中！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
        ElseIf llngItemCount = 0 Then
            MsgBox "收费标准中的所有收费项目在费用清单中已添加！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
        Else
            MsgBox "收费标准中部分收费项目在费用清单中已添加,其余的 " & llngItemCount & " 条已添加到费用清单！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
        End If
    End If
                
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm划价", "ccmb收费标准_Click", Err.Number, Err.Description, False
End Sub

Private Sub ccmd定位_Click()
    Dim lrds打折信息 As Object               '单位的打折信息
    Dim lrdsTemp As Object
    
    On Error GoTo errhandler
    
    '调用单位档案的定位接口获取单位信息
    Set lrdsTemp = pobj单位定位.func单位简单定位(100, 100)
    If Not (lrdsTemp Is Nothing) Then
        If lrdsTemp.RecordCount > 0 Then
            '显示单位名称`
            ctxtInput(收费_交费单位).Text = lrdsTemp("单位名称")
            '显示卫生种类、片区
            ccmb卫生种类.Text = lrdsTemp("卫生种类")
            ccmb片区.Text = IIf(IsNull(lrdsTemp("片区")), "", lrdsTemp("片区"))
            
            '保存单位的申请编号
            mstr交费单位编号 = lrdsTemp("申请编号")
            ctxtInput(收费_交费单位).SetFocus
        End If
    End If


    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm划价", "ccmd定位_Click", Err.Number, Err.Description, False
End Sub



Private Sub cgrdDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim lcurMoney As Currency
    
    On Error GoTo errhandle
    ctlb工具栏.Buttons(1).Enabled = True
    Select Case cgrdDetail.TextMatrix(0, Col)
        Case "数量"
            '判断输入的是否数值
            If Len(cgrdDetail.TextMatrix(Row, Col)) > 4 Then
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
            Else
                If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) And Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    '是数值
                    '计算金额
                    cgrdDetail.TextMatrix(Row, 费用清单_金额) = cgrdDetail.TextMatrix(Row, 费用清单_单价) * cgrdDetail.TextMatrix(Row, 费用清单_数量)
                Else
                    '不是数值
                    'Undo
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
                End If
            End If
        Case "单价"
            Dim lcur单价 As Currency
            If mcur最小单价 = mcur最大单价 Then
                sffuncMsg "该收费项目单价已定,不可修改！", sf警告
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                Exit Sub
            End If
            
            If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) Then
                If Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    If Val(cgrdDetail.TextMatrix(Row, Col)) <= mcur最大单价 And Val(cgrdDetail.TextMatrix(Row, Col)) >= mcur最小单价 Then
                        cgrdDetail.TextMatrix(Row, 费用清单_金额) = cgrdDetail.TextMatrix(Row, 费用清单_单价) * cgrdDetail.TextMatrix(Row, 费用清单_数量)
                    Else
                        sffuncMsg "输入的单价超出范围！", sf警告
                        cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                    End If
                Else
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                End If
            Else
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
            End If
        Case Else
    End Select
    
    sub计算总金额
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm划价", "cing费用清单_AfterEdit", Err.Number, Err.Description, False
    
End Sub

Private Sub cgrdDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
     
    Select Case Col
        Case 费用清单_数量
            ctlb工具栏.Buttons(1).Enabled = False
            mstrUndoCount = cgrdDetail.TextMatrix(Row, Col)
            
        Case 费用清单_单价
            ctlb工具栏.Buttons(1).Enabled = False
                        
            '获取最小单价,最大单价.
            Dim lobjRec As Object
            Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & cgrdDetail.TextMatrix(Row, 0) & "'")
            If lobjRec.RecordCount > 0 Then
                mcur最小单价 = IIf(IsNull(lobjRec("最小单价").Value), 0, lobjRec("最小单价").Value)
                mcur最大单价 = IIf(IsNull(lobjRec("最大单价").Value), 99999999, lobjRec("最大单价").Value)
            Else
                sffuncMsg "未找到该收费项目的设置信息，该设置信息可能已被修改或删除，请退出收费界面，重新进入！"
            End If
            mstrUndoMoney = cgrdDetail.TextMatrix(Row, Col)
        Case Else
            ctlb工具栏.Buttons(1).Enabled = True
            Cancel = True
    End Select
End Sub



Private Sub cgrdDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete
            mobj界面通用对象_BeforeOperate "删除", False
    End Select

End Sub

Private Sub cgrdDetail_LostFocus()
    On Error Resume Next
    ctlb工具栏.Buttons(1).Enabled = True
End Sub


Private Sub cgrdItem_Click()
    On Error Resume Next
    ctxtFind.Text = ""
End Sub

Private Sub cgrdItem_DblClick()
    Dim lstrCode As String
    
    On Error GoTo errhandler
    '添加收费项目
    lstrCode = cgrdItem.TextMatrix(cgrdItem.Row, 1)
    lstrCode = Right(lstrCode, Len(lstrCode) - InStr(lstrCode, " "))
    If InStr(lstrCode, " ") > 0 Then lstrCode = Right(lstrCode, Len(lstrCode) - InStr(lstrCode, " "))
    lstrCode = Trim(lstrCode)
    
    If Not func检查项目是否已选(lstrCode) Then
        sub添加项目 lstrCode
    End If
                    
Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm划价", "clst收费项目_DblClick", Err.Number, Err.Description, False
          
End Sub

Private Sub sub添加项目(ByVal paraCode As String)
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & paraCode & "'")
    cgrdDetail.AddItem paraCode & vbTab & lobjRec("收费项目名称") & vbTab & _
                    lobjRec("单价") & vbTab & "1" & vbTab & lobjRec("单价")
    
    sub计算总金额
End Sub


Private Sub sub计算总金额()
    Dim lcurMoney As Double
    Dim i As Long
    
    For i = 1 To cgrdDetail.Rows - 1
        lcurMoney = Format(lcurMoney + cgrdDetail.ValueMatrix(i, 费用清单_金额), "0.00")
    Next

    clblTotal.Caption = "总金额：" & lcurMoney

End Sub


'功能：根据输入的助记符查找项目。
Private Sub ctxtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long
    If KeyCode = 13 Then
        '回车选择项目。
        If cgrdItem.Row > 0 Then
            cgrdItem_DblClick
        End If
    Else
        '定位。
        Dim lCol As Long
        cgrdItem.Row = 0
        If ctxtFind.Text <> "" Then
            For i = 1 To cgrdItem.Rows - 1
                If UCase(Left(cgrdItem.TextMatrix(i, 2), Len(ctxtFind))) = UCase(ctxtFind.Text) Then
                    cgrdItem.Select i, 0, i, cgrdItem.Cols - 1
                    cgrdItem.TopRow = i
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If ActiveControl = ctxtFind Then
        Else
            SendKeys Chr(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lcol工具栏 As Collection
    Dim i As Long
    Dim lobjRec As Object
    
    On Error GoTo errhandle
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
        
    Set mobj界面通用对象 = New cls界面通用对象
    Set mobj界面通用对象.Form = Me
    Set mobj界面通用对象.c工具栏 = ctlb工具栏
    
    Set lcol工具栏 = New Collection
    
    lcol工具栏.Add "保存"
    lcol工具栏.Add "|"
    lcol工具栏.Add "删除"
    lcol工具栏.Add "清空"
    lcol工具栏.Add "|"
    lcol工具栏.Add "退出"
    
    mobj界面通用对象.subInitialize lcol工具栏, ""
    
    mint科目级数 = Val(pobj收费管理.业务设置("科目级数"))
    
    sub初始化窗体
    
    If pstr收费编号 <> "" Then
        '内部收费,显示费用信息。
        sub显示费用信息
    Else
        mstr交费单位编号 = pstr单位编号
        If pstr单位编号 <> "" Then
            Set lobjRec = dafuncGetData("select * from 单位档案_单位基本信息表 where 申请编号='" & pstr单位编号 & "'")
            If lobjRec.RecordCount > 0 Then
                ctxtInput(收费_交费单位) = lobjRec!单位名称
                ccmb卫生种类.Text = lobjRec!卫生种类
                ccmb片区.Text = IIf(IsNull(lobjRec!片区), "", lobjRec!片区)
            End If
        End If
    End If
    
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm划价", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub sub显示费用信息()
    Dim lobjRec As Object
    Dim i As Long
    
    On Error GoTo errhandler
    
    If pstr收费编号 <> "" Then
        '修改收费记录。
        Set lobjRec = dafuncGetData("select a.收费批号,a.收费编号,a.收费项目编号,收费项目名称=(select 收费项目名称 from 收费管理_收费项目字典表 where 收费项目编号=a.收费项目编号),a.数量,计量单位=(select 计量单位 from 收费管理_收费项目字典表 where 收费项目编号=a.收费项目编号),a.单价,a.金额,a.收费状态,a.交费方式,a.交费人,a.交费单位编号,交费单位=(select 单位名称 from 单位档案_单位基本信息表 where 申请编号=a.交费单位编号),a.交费日期,a.退费日期,收费人编号=a.收费人,收费人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.收费人),退费人编号=a.退费人,退费人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.退费人) ,主管科室经手人编号=a.主管科室经手人,主管科室经手人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.主管科室经手人),主管科室编号,主管科室=(select 名称 from 系统管理_科室字典表 where 编号=a.主管科室编号),打折比率,备注1,备注2  from 收费管理_费用信息表 a where 收费编号='" & pstr收费编号 & "'  and 收费状态=0")
        
        cgrdDetail.Rows = 1
    
        Do While Not lobjRec.EOF
            cgrdDetail.AddItem lobjRec("收费项目编号") & vbTab & _
                lobjRec("收费项目名称") & vbTab & _
                lobjRec("单价") & vbTab & _
                lobjRec("数量") & vbTab & _
                lobjRec("金额")
            lobjRec.MoveNext
        Loop
        If lobjRec.RecordCount > 0 Then
            lobjRec.MoveFirst
            mstr交费单位编号 = IIf(IsNull(lobjRec("交费单位编号").Value), "", lobjRec("交费单位编号").Value)
        
            ctxtInput(收费_收费编号).Text = lobjRec("收费编号")
            
            If IIf(IsNull(lobjRec("主管科室")), "", lobjRec("主管科室")) <> "" Then
                For i = 0 To ccmb主管科室.ListCount - 1
                    If ccmb主管科室.List(i) = IIf(IsNull(lobjRec("主管科室")), "", lobjRec("主管科室")) Then
                        ccmb主管科室.ListIndex = i
                        Exit For
                    End If
                Next
            Else
                ccmb主管科室.ListIndex = -1
            End If
            
            ccmb卫生种类.Text = IIf(IsNull(lobjRec("备注1").Value), "", lobjRec("备注1").Value)
            ccmb片区.Text = IIf(IsNull(lobjRec("备注2").Value), "", lobjRec("备注2").Value)
        
        
            ctxtInput(收费_交费人).Text = lobjRec("交费人")
            ctxtInput(收费_交费单位).Text = IIf(IsNull(lobjRec("交费单位").Value), "", lobjRec("交费单位").Value)
        
        End If
        
    End If

    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm划价", "sub显示费用信息", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mobj界面通用对象 = Nothing
    
End Sub


Private Sub sub清除界面()
    Dim i As Integer
    
    On Error GoTo errhandle
    clblTotal = "总金额：0"
    
    '在保存后需要清空收费编号;徐冀川;2002/9/30
    mstr交费单位编号 = ""
  
    Dim lobjCtrl As Control
    For Each lobjCtrl In ctxtInput
        lobjCtrl.Text = ""
    Next
    cgrdDetail.Rows = 1
    
    ccmb主管科室.Text = um用户所属科室
    
    ctxtInput(收费_交费人).SetFocus
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm划价", "sub清除收费界面", Err.Number, Err.Description, True
End Sub


Private Sub sub初始化窗体()
On Error GoTo errhandle
    Dim lobj收费标准 As Object
    Dim lobj科室 As Object
    Dim lobj交费方式 As Object

    mstrUndoCount = ""
    mstrUndoMoney = ""
    mstr交费单位编号 = ""
    mint交费方式编号 = 0
    
    Dim i As Long
    Dim j As Long
    
    Set lobj收费标准 = dafuncGetData("select 收费标准名称,助记符 from 收费管理_收费标准信息表 group by 助记符,收费标准名称")
    Set lobj科室 = dafuncGetData("select * from 系统管理_科室字典表")
    Set lobj交费方式 = dafuncGetData("select * from 收费管理_交费方式字典表")
    
    
    '初始化 "cing费用清单"
    With cgrdDetail
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, 费用清单_收费项目编号) = "收费项目编号"
        .ColWidth(费用清单_收费项目编号) = 1310
        .ColAlignment(费用清单_收费项目编号) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_收费项目名称) = "收费项目名称"
        .ColWidth(费用清单_收费项目名称) = 1320
        
        .TextMatrix(0, 费用清单_单价) = "单价"
        .ColWidth(费用清单_单价) = 480
        
        .TextMatrix(0, 费用清单_数量) = "数量"
        .ColWidth(费用清单_数量) = 500
        
        .TextMatrix(0, 费用清单_金额) = "金额"
        .ColWidth(费用清单_金额) = 570
    End With
    
    '初始化 "收费标准"
    ccmb收费标准.Clear
    Do While Not lobj收费标准.EOF
        ccmb收费标准.AddItem lobj收费标准("收费标准名称").Value
        lobj收费标准.MoveNext
    Loop
    
    '初始化 "主管科室"列表
    ccmb主管科室.Clear
    If Not (lobj科室 Is Nothing) Then
        Do While Not lobj科室.EOF
            ccmb主管科室.AddItem lobj科室("名称").Value
            lobj科室.MoveNext
        Loop
    End If
    ccmb主管科室.Text = um用户所属科室
    
    '获取收费项目大类。
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select 收费项目编号,收费项目名称 from 收费管理_收费项目字典表 where len(收费项目编号)=3  order by 收费项目编号 ")
    Do While Not lobjRec.EOF
        Ccbo收费项目大类.AddItem lobjRec("收费项目名称")
        lobjRec.MoveNext
    Loop
    
    Ccbo收费项目大类.ListIndex = 0
    
    '获取卫生种类
    Set lobjRec = dafuncGetData("select * from 系统管理_卫生种类字典视图 order by 编号")
    ccmb卫生种类.Clear
    ccmb卫生种类.AddItem ""
    Do While Not lobjRec.EOF
        ccmb卫生种类.AddItem lobjRec("名称").Value
        lobjRec.MoveNext
    Loop
    
    '获取片区
    Set lobjRec = dafuncGetData("select * from 系统管理_片区字典视图 order by 编号")
    ccmb片区.Clear
    ccmb片区.AddItem ""
    Do While Not lobjRec.EOF
        ccmb片区.AddItem lobjRec("名称").Value
        lobjRec.MoveNext
    Loop
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面部件", "frm划价", "sub初始化窗体", Err.Number, Err.Description, True
End Sub






Private Sub mobj界面通用对象_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long, j As Long
    
    On Error GoTo errhandle
    Select Case Operate
        Case "保存"
            Cancel = True
            '校验数据合法性。
            If Not ValidateData Then Exit Sub
            
            '收集要保存的费用信息。
            Dim lstr主管科室编号 As String
            Dim lcol记录 As Collection
            Dim lcol数据 As Collection
            Dim lstr收费编号 As String
            
            If ccmb主管科室.ListIndex >= 0 Then
                lstr主管科室编号 = ccmb主管科室.ItemData(ccmb主管科室.ListIndex)
                lstr主管科室编号 = Right(lstr主管科室编号, Len(lstr主管科室编号) - 1)
            Else
                lstr主管科室编号 = um用户所属科室编号
            End If
            Set lcol数据 = New Collection
            For i = 1 To cgrdDetail.Rows - 1
                Set lcol记录 = New Collection
                For j = 0 To cgrdDetail.Cols - 1
                    lcol记录.Add cgrdDetail.TextMatrix(i, j), cgrdDetail.TextMatrix(0, j)
                Next
                '添加收费其他字段
                lcol记录.Add ctxtInput(收费_交费人).Text, "交费人"
                lcol记录.Add mstr交费单位编号, "交费单位编号"
                lcol记录.Add ctxtInput(收费_交费单位).Text, "交费单位名称"
                lcol记录.Add lstr主管科室编号, "主管科室编号"
                lcol记录.Add um用户编号, "主管科室经手人"
                lcol记录.Add ccmb卫生种类.Text, "备注1"
                lcol记录.Add ccmb片区.Text, "备注2"
                lcol数据.Add lcol记录
            Next
            
            '保存划价信息。
            lstr收费编号 = pobj收费管理.func划价保存(lcol数据, ctxtInput(收费_收费编号), pstr业务分类)
            
            If cchk清空.Value = 1 Then
                sub清除界面
            Else
                ctxtInput(收费_收费编号) = lstr收费编号
            End If
            pstr收费编号 = lstr收费编号
            ctxtInput(收费_交费人).SetFocus
            
        Case "删除"
            If cgrdDetail.Row > 0 Then
                cgrdDetail.RemoveItem cgrdDetail.Row
                
                sub计算总金额
            End If
            
        Case "清空"
            sub清除界面
            
        Case Else
    End Select
    Exit Sub
    
errhandle:
    sfsub错误处理 "收费界面部件", "frm划价", "mobj界面通用对象_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Function func检查项目是否已选(ByVal Para收费项目编号 As String) As Boolean
On Error GoTo errhandle
    Dim i As Long
    func检查项目是否已选 = False
    If cgrdDetail.Rows = 1 Then
        func检查项目是否已选 = False
        Exit Function
    End If
    
    For i = 1 To cgrdDetail.Rows - 1
        If Para收费项目编号 = cgrdDetail.TextMatrix(i, 费用清单_收费项目编号) Then
            func检查项目是否已选 = True
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub错误处理 "收费界面部件", "frm划价", " func检查项目是否已选()", Err.Number, Err.Description
End Function

Private Function ValidateData() As Boolean
    On Error GoTo errhandle
    ValidateData = False
    If ctxtInput(收费_交费人).Text = vbNullString And ctxtInput(收费_交费单位) = vbNullString Then
        sffuncMsg """交费人"" 和 ""交费单位"" 必须输入其中之一！", sf警告
        Exit Function
    End If
    If cgrdDetail.Rows = 1 Then
        sffuncMsg "无费用信息可以保存！", sf警告
        Exit Function
    End If
    

    ValidateData = True
    Exit Function
errhandle:
    sfsub错误处理 "收费界面部件", "frm划价", " ValidateData()", Err.Number, Err.Description, True
End Function


