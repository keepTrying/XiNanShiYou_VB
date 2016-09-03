VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frm查找人员 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "查找"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5820
   ClipControls    =   0   'False
   Icon            =   "frm查找人员.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex6Ctl.vsFlexGrid cgrdPerson 
      Height          =   2295
      Left            =   600
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      _cx             =   59384321
      _cy             =   59379664
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
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   1
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
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1155
   End
   Begin VB.CommandButton ccmdLocateUnit 
      Caption         =   "定位(&L)"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Width           =   945
   End
   Begin VB.OptionButton coptChoise 
      Caption         =   "其他"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Value           =   -1  'True
      Width           =   720
   End
   Begin VB.TextBox ctxtName 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1155
   End
   Begin VB.OptionButton coptChoise 
      Caption         =   "健康证号"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox ctxtHealthNo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   2745
   End
   Begin VB.ComboBox ccmbSex 
      Height          =   300
      ItemData        =   "frm查找人员.frx":000C
      Left            =   1560
      List            =   "frm查找人员.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1725
   End
   Begin VB.ComboBox ccmbQueryUnit 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox ctxtId 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   600
      Width           =   2745
   End
   Begin VB.OptionButton coptChoise 
      Caption         =   "身份证号"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请在列表中选择体检人员，双击鼠标（或按“确定”按钮）返回！"
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   600
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   5220
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位"
      Height          =   180
      Index           =   7
      Left            =   1080
      TabIndex        =   11
      Top             =   1800
      Width           =   360
   End
End
Attribute VB_Name = "frm查找人员"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstr体检类型 As String '初检/年检/复查。
Public pstr系统编号 As String '查找出来的系统编号。

Private Sub ccmdCancel_Click()
    pstr系统编号 = ""
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    Dim lobj体检集 As Object
    Dim lobj体检人员 As Object  '用来获取体检人员最近次体检记录。
    Dim lobj体检 As Object      '体检人员最近次体检。
    Dim lobjRec As Object       'Recordset，体检集对象返回的元素集。
    Dim lstr系统编号 As String  '健康证号对应的体检系统编号。
    Dim lstrError As String
    Dim i As Integer
    Dim j As Long
    
    On Error GoTo errHandler
    If cgrdPerson.Visible And cgrdPerson.Row > 0 Then
        '在列表中选择人员，并确定返回。
        cgrdPerson_DblClick
        
    Else
        '创建体检集对象。
        Set lobj体检集 = CreateObject("职业病对象.clsMedicalExamSet")
        lobj体检集.subClear
        If coptChoise(0).Value Or coptChoise(2).Value Then
            '其它、身份证。
            If coptChoise(0).Value Then
                If Trim(ctxtName.Text) = "" And ccmbSex.Text = "" And Trim(ccmbQueryUnit.Text) = "" Then
                    Err.Raise 6666, , "必须输入姓名、性别、单位名称！"
                End If
                
                '设置体检集对象的搜索定位条件属性。
                With lobj体检集
                    .姓名 = Trim(ctxtName.Text)
                    .性别 = ccmbSex.Text
                    .单位名称 = Trim(ccmbQueryUnit.Text)
                End With
            Else
                '按身份证号查询。
                If Trim(ctxtId.Text) = "" Then
                    Err.Raise 6666, , "你必须输入身份证号！"
                End If
                
                '设置体检集对象的搜索定位条件属性。
                With lobj体检集
                    .身份证号 = Trim(ctxtId.Text)
                End With
            
            End If
        ElseIf coptChoise(1).Value Then
            '健康证号、系统编号。
            If Trim(ctxtHealthNo.Text) = "" Then
                Err.Raise 6666, , "你必须输入" & coptChoise(1).Caption & "！"
            End If
            If coptChoise(1).Caption = "系统编号" Then
                lobj体检集.从系统编号 = Trim(ctxtHealthNo.Text)
                lobj体检集.到系统编号 = Trim(ctxtHealthNo.Text)
            Else
                '将健康证号转换成系统编号。
                lstr系统编号 = pobj业务对象.Func根据健康证条码号获取体检系统编号(Trim(ctxtHealthNo.Text))
                If lstr系统编号 = "" Then
                    Err.Raise 6666, , "你输入的健康证号没有对应的体检记录。"
                End If
                lobj体检集.从系统编号 = lstr系统编号
                lobj体检集.到系统编号 = lstr系统编号
            End If
        End If
        
        If pstr体检类型 = "复查" Then
            '需要复查，并且还没有进行复查登记。
            lobj体检集.复查标志 = 1
            lobj体检集.复查系统编号 = ""
        End If
        
        '获取满足定位条件的体检系统编号。
        '修改：2001-12-30（按姓名、体检日期排序）。
        Set lobjRec = lobj体检集.元素集("系统编号,姓名,单位名称,性别,年龄=datediff(year,出生日期,getdate()),体检日期" & IIf(pstr体检类型 = "复查", ",复查体检表名", ""), "姓名,单位名称,体检日期 desc")
        If lobjRec.RecordCount = 0 Then
            '没找到相应体检人员
            lstrError = "未查找到满足你输入条件的体检人员。"
            If pstr体检类型 = "复查" Then
                lstrError = lstrError & "可能是：" & Chr(13) & Chr(10) & "(1) 该体检不需要复查，即体检医师在下体检结论时没有设置要复查，以及复查的体检表；" & Chr(13) & Chr(10) & "(2) 或该体检已复查登记过了。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "解决办法：" & Chr(13) & Chr(10) & "(1) 请重新输入查找条件。" & Chr(13) & Chr(10) & "(2) 或重新下体检结论，设置为要复查。"
            End If
            Err.Raise 6666, , lstrError
        Else
            If lobjRec.RecordCount > 1 Then
                lstr系统编号 = ""
                If lobjRec.RecordCount > 100 Then
                    '查询结果太多，提示用户。
                    If Not sffuncMsg("你输入的查询条件范围太大，查询结果超过100条记录，这对你查找具体体检人员没有多大帮助，请缩小范围。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "你坚持要在这么多记录中选择吗？", sf询问) Then
                        Exit Sub
                    End If
                End If
                '查找到多条记录时弹出list，并加入到Grid中。
                cgrdPerson.Rows = lobjRec.RecordCount + 1
                cgrdPerson.Cols = lobjRec.Fields.Count
                For j = 0 To cgrdPerson.Cols - 1
                    cgrdPerson.TextMatrix(0, j) = lobjRec.Fields(j).Name
                Next
                i = 1
                Do While Not lobjRec.EOF
                    For j = 0 To cgrdPerson.Cols - 1
                        cgrdPerson.TextMatrix(i, j) = IIf(IsNull(lobjRec(j).Value), "", lobjRec(j).Value)
                    Next
                    lobjRec.MoveNext
                    i = i + 1
                Loop
                cgrdPerson.AutoSize 0, cgrdPerson.Cols - 1
                cgrdPerson.Visible = True
                clblInfo.Visible = True
                cgrdPerson.SetFocus
            Else
                lstr系统编号 = lobjRec(0)
            End If
            
        End If
        
        Set lobj体检人员 = Nothing
        Set lobj体检 = Nothing
        
        '查找成功，返回。
        If lstr系统编号 <> "" Then
            pstr系统编号 = lstr系统编号
            Unload Me
        Else
            '需要在列表中选择，或重新输入查询条件。
        End If
    End If
    
    Exit Sub
    
errHandler:
    
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frm查找人员", "ccmdOk_Click", 6666, lstrError, False
    
    If coptChoise(0).Value Then
        ctxtName.SetFocus
    ElseIf coptChoise(1).Value Then
        ctxtHealthNo.SetFocus
    Else
        ctxtId.SetFocus
    End If
    
    Exit Sub
    Resume
End Sub

Private Sub cgrdPerson_DblClick()
    On Error Resume Next
    If cgrdPerson.Row < 1 Then Exit Sub
    
    '返回。
    pstr系统编号 = cgrdPerson.TextMatrix(cgrdPerson.Row, 0)
    
    '本列表框消失。
    If pstr系统编号 <> "" Then
        cgrdPerson.Visible = False
    
        Unload Me
    End If
    
End Sub

Private Sub cgrdPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And cgrdPerson.Row > 0 Then
        cgrdPerson_DblClick
    End If

End Sub

Private Sub cgrdPerson_LostFocus()
    On Error Resume Next
    If ActiveControl.Name = "ccmdOk" And cgrdPerson.Row > 0 Then
    
    Else
        cgrdPerson.Visible = False
        clblInfo.Visible = False
    End If
End Sub

Private Sub coptChoise_Click(Index As Integer)
    On Error GoTo errHandler
    ctxtName.Enabled = False
    ccmbSex.Enabled = False
    ccmbQueryUnit.Enabled = False
    ccmdLocateUnit.Enabled = False
    ctxtId.Enabled = False
    ctxtHealthNo.Enabled = False
    
    If coptChoise(0).Value Then
        '选择输入姓名。
        ctxtName.Enabled = True
        ccmbSex.Enabled = True
        ccmbQueryUnit.Enabled = True
        ccmdLocateUnit.Enabled = True
        ctxtName.SetFocus
    ElseIf coptChoise(1).Value Then
        '选择输入健康档案编号。
        ctxtHealthNo.Enabled = True
        ctxtHealthNo.SetFocus
    ElseIf coptChoise(2).Value Then
        '选择输入身份证号。
        ctxtId.Enabled = True
        ctxtId.SetFocus
    End If
    
    Exit Sub
errHandler:
    'sfsub错误处理 "职业病界面部件", "frm查找人员", "coptChoise_Click", Err.Number, Err.Description, False
End Sub

Private Sub ctxtName_GotFocus()
    On Error Resume Next
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub
Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub
'功能：自动弹出列表框
Private Sub ccmbQueryUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbQueryUnit
    Exit Sub
errHandler:
End Sub

Private Sub ccmbQueryUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdOk.SetFocus
    End If
End Sub

Private Sub ccmbQueryUnit_LostFocus()
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '判断录入的单位是否在列表中存在，不存在则加入列表。
    i = gffuncItemIsInComboBox(ccmbQueryUnit, ccmbQueryUnit.Text)
    
    If i = -1 Then
        '加到ccmbQueryUnit中。
        ccmbQueryUnit.AddItem ccmbQueryUnit.Text
    End If
    
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbQueryUnit.SetFocus
    End If
End Sub
'调用单位定位
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '单位定位返回的结果记录。

    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbQueryUnit.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
        End If
    End If
    
    '把焦点回到单位录入框。
    ccmbQueryUnit.SetFocus
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "frm查找人员", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub


Private Sub ctxtHealthNo_GotFocus()
    On Error Resume Next
    With ctxtHealthNo
        .SelStart = 0
        .SelLength = Len(Trim(ctxtHealthNo.Text))
    End With
End Sub

Private Sub ctxtHealthNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdOk.SetFocus
    End If
End Sub
Private Sub ctxtId_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdOk.SetFocus
    End If
End Sub



Private Function func根据健康档案编号获取系统编号(ByVal para健康档案编号 As String) As String
    Dim lobj体检人员  As Object 'clsPersonExamed.
    Dim lobj体检 As Object      'clsMedicalExam
    Dim lstr系统编号 As String
    
    func根据健康档案编号获取系统编号 = ""
    
    '获取该人最近一次体检记录。
    '创建体检人员对象。
    Set lobj体检人员 = CreateObject("职业病对象.clsPersonExamed")
    lobj体检人员.健康档案编号 = para健康档案编号
    Set lobj体检 = lobj体检人员.Func获取本人最近一次体检
    If Not lobj体检 Is Nothing Then
        lstr系统编号 = lobj体检.系统编号
    Else
        Err.Raise 6666, , "该体检人员还没有在本体检中心体检过，无法进行年检登记。请选择初检登记。"
    End If
            
    func根据健康档案编号获取系统编号 = lstr系统编号
    
End Function

Private Sub Form_Load()
    Dim lcolInfo  As Collection
    Dim i As Long
    On Error Resume Next
    
    If pstr体检类型 = "年检" Then
       coptChoise(1).Caption = "健康证号"
    Else
       coptChoise(1).Caption = "系统编号"
    End If
    '从当日工作及已簿中获取当天录入过的单位名称。
    Set lcolInfo = pobj业务对象.当日工作记忆簿.单位名称集
    For i = 1 To lcolInfo.Count
        ccmbQueryUnit.AddItem lcolInfo(i)
    Next

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cgrdPerson.Visible = False
    clblInfo.Visible = False
End Sub
