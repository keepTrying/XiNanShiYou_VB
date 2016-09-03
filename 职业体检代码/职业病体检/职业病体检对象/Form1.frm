VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.2#0"; "dyCatchPhoto.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "体检人员"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "体检"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   240
      Width           =   1455
   End
   Begin VSFlex6DAOCtl.vsFlexGrid cgrdResult 
      Height          =   2175
      Left            =   2280
      TabIndex        =   8
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3836
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
      SelectionMode   =   1
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
   Begin dyCatchPhoto.ctlCatchPhoto ccpMain 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6165
      BackColor       =   0
      FontSize        =   11.25
   End
   Begin VB.CommandButton Command8 
      Caption         =   "体检项目"
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "体检管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "体检结论"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "体检医师"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "体检集"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "体检表模板"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2280
      TabIndex        =   11
      Top             =   3720
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "或取的体检人员相片"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    On Error GoTo errHandler
    
    '初始化数据访问对象(连接本机)。
'    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=防疫2001;Data Source=YANGCHUN"
'    If Not umfunc校验身份("5555", "") Then
'        sffuncMsg "校验身份失败5555。", sf警告
'    End If
        
    '初始化数据访问对象(Tdcserver)。
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=防疫2001;Data Source=Tdcserver"
    If Not umfunc校验身份("5555", "") Then
        sffuncMsg "校验身份失败5555。", sf警告
    End If
        
    '初始化数据访问对象(Testserver)。
'    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=dyfy;Persist Security Info=True;User ID=sa;Initial Catalog=卫生防疫新版数据库;Data Source=TESTSERVER"
'    If Not umfunc校验身份("0008", "") Then
'        sffuncMsg "校验身份失败0008。", sf警告
'    End If
        

    Exit Sub
errHandler:
    sfsub错误处理 "测试", "", "Form_load", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub
Private Sub Command9_Click()
    Dim lobj体检 As clsMedicalExam
    Dim lobjRec As Object
    Dim lstrTemp As String
    
    Set lobj体检 = New clsMedicalExam
    With lobj体检
        '新增体检记录。
        lstrTemp = lobj体检.Func分配系统编号
        .系统编号 = lstrTemp
        
        '设置其它属性。
        .体检类别 = P_EXAM_FIRST '初检。
        .体检人员.姓名 = "张三"
        .体检人员.单位名称 = "体检测试小组"
        '...
        
        .体检表.体检表名 = "从业人员体检表"
        Debug.Print .体检表.试管编号字母
        
        Debug.Assert 1 = 2
        
        '检查试管编号字母是否已分配。
        If .体检表.试管编号字母 = "" Then
            .体检表.试管编号字母 = "A"
        End If
        .体检人员.姓名 = ""
        Debug.Assert 1 = 2
        
        .Sub保存体检登记信息
        
        '核对库中：体检基本信息表、体检人员信息表、体检结果信息表中是否增加了相应记录。
        Debug.Assert 1 = 2
        
    End With
    
    With lobj体检
        '设置系统编号为库中存在的。
        .系统编号 = "123401010402002"
        
        '检查属性（并与数据库核对是否一致）。
        '...
        Debug.Assert 1 = 2
        
        '测试方法。
        lstrTemp = .func获取系统编号的前一个号("00000103280021")
        Debug.Assert 1 = 2
        lstrTemp = .func获取系统编号的后一个号("00000103280021")
        Debug.Assert 1 = 2
        
        lstrTemp = .Func分配系统编号
        '检查“体检管理_最大流水号表”是否递加。
        Debug.Assert 1 = 2
        
        .sub退回系统编号 lstrTemp
        
        '检查“体检管理_最大流水号表”是否恢复。
        Debug.Assert 1 = 2
        
        
    End With
End Sub



Private Sub Command1_Click()
    On Error GoTo errHandler
   '测试体检人员。
    Dim lobjPerson As clsPersonExamed
    Dim lobjRec As Object
    Dim lcolInfo As New Collection
    
    Set lobjPerson = New clsPersonExamed
    
    With lobjPerson
        '设置体检人员属性。
        .健康档案编号 = .Func分配健康档案编号(lcolInfo)
    
        .姓名 = "张宏"
        .公民身份号码 = "510223450608120"
        .单位名称 = "电力厂"
                
        Debug.Assert 1 = 2
        
        .Sub保存
        
        Debug.Assert 1 = 2
        
        '获取相片。
        Picture1.Picture = .像片
        
        Debug.Assert 1 = 2
        
        '获取本人最近一次体检。
        Set lobjRec = .Func获取本人最近一次体检
        
        '显示查询结果。
        gfsubLoadGridFromRec cgrdResult, lobjRec
    End With

    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command1_Click", Err.numer, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler
    Dim lobj体检表 As ClsMedicalExamSheet
    
    Set lobj体检表 = New ClsMedicalExamSheet
    With lobj体检表
        .系统编号 = "11111200103200004"
        
        Debug.Assert 1 = 2
        '检查属性。
        '...
        
        '重新选择体检表。
        .体检表名 = "从业人员体检"
        Debug.Assert 1 = 2
        
        '填体检结果。
        .Sub填体检结果 "0101", "100", "1234", Date
                
        Debug.Assert 1 = 2
        
        '保存体检结果。
        .Sub保存体检结果
        
    End With
    
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command2_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command3_Click()
    On Error GoTo errHandler
    Dim lobj体检表模板 As clsMedicalExamTemplate
    
    Set lobj体检表模板 = New clsMedicalExamTemplate
    With lobj体检表模板
        '设置属性。
        .体检表名 = "复查已肝两对办"
        .代号 = 9
        .体检单名称 = "单个试管体检单"
        .试管字母编号 = "B"
        .是否复查体检表 = True
        .收费标准 = "复查已肝两对办收费"
        
        .Sub添加附加项目 "卫生种类", True
        .Sub添加附加项目 "片区", True
        
        
        .Sub添加体检项目 "0001"
        .Sub添加体检项目 "0002"
                
        .Sub添加体检结论 3435
        
        Debug.Assert 1 = 2
        .Sub保存模板
        
        Debug.Assert 1 = 2
        .Sub复制模板 "复查已肝三对办"
        .代号 = 10
        .Sub保存模板
        
        Debug.Assert 1 = 2
        .Sub删除模板
        
    End With
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command3_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command4_Click()
    Dim lobj体检集 As clsMedicalExamSet
    Dim lobjRec As Object
    
    Set lobj体检集 = New clsMedicalExamSet
    
    With lobj体检集
        .从体检日期 = "2001-3-01"
        .到体检日期 = "2001-4-01"
    
        '.单位名称 = "111"
    
        '.从试管编号 = "A:0001"
        '.到试管编号 = "A:0010"
    
        '获取需要复查的。
        '.复查标志 = 1
    
        '获取需要复查但还未复查的。
        .复查标志 = 1
        .复查系统编号 = ""
    End With
        
    '获取查询结果。
    Set lobjRec = lobj体检集.元素集
    
    '显示查询结果。
    clblInfo = "从体检日期：" & lobj体检集.从体检日期 & "，到：" & lobj体检集.到体检日期 & "，需要复查但还未复查的体检记录"
    
    gfsubLoadGridFromRec cgrdResult, lobjRec
    
End Sub

Private Sub Command5_Click()
    '测试体检医师。
    Dim lobj体检医师 As ClsMedicalExaminer
    Dim lblnCan As Boolean
    Dim lcolInfo As Collection
    
    Set lobj体检医师 = New ClsMedicalExaminer
    With lobj体检医师
        .编号 = "5555"
        
        '获取可作体检项目。
        Set lcolInfo = .可作体检项目
        
        Debug.Assert 1 = 2
'        .Sub添加体检项目 "0101"
'        .Sub添加体检项目 "0102"
'        .Sub添加体检项目 "0201"
'        .Sub删除体检项目 "0201"
'        .Sub添加体检项目 "0202"
'
'        Set lcolInfo = .可作体检项目
       
        lblnCan = .func是否可作项目("0101")
        Debug.Assert 1 = 2
        
        lblnCan = .func是否可作项目("0201")
        Debug.Assert 1 = 2
        
        Set lcolInfo = .Func获取本人指定体检表上可作的体检项目("11111200103200001", "常规")
        Debug.Assert 1 = 2
    End With

End Sub

Private Sub Command6_Click()
    '测试体检结论。
    Dim lobj体检结论 As ClsMedicalExamConclusion
    Dim lobj体检结论条件 As ClsConclusionFilter
    Dim lcolInfo As Collection
    Dim lbln满足 As Boolean
    Set lobj体检结论条件 = New ClsConclusionFilter
    With lobj体检结论条件
        .ID = 3406
        .编号 = 2
        .SubAddFilter 1, "0001", "眼", "=", "异常"
        .SubAddFilter 2, "0001", "眼", "=", "异常"
        
        .subSave
        Debug.Assert 1 = 2

        .SubRemoveFilter 2
        .subSave
        Debug.Assert 1 = 2

        lbln满足 = .Func判断是否满足条件("11111200103200001")
        Debug.Assert 1 = 2
        
    End With
    
    Set lobj体检结论 = New ClsMedicalExamConclusion
    With lobj体检结论
        .ID = 3406
        Set lobj体检结论条件 = .判断条件(1)
        Debug.Assert 1 = 2
        
        .Sub删除条件分组 1
        Set lcolInfo = .所有判断条件
        Debug.Assert 1 = 2
        
        lbln满足 = .Func判断是否可下本结论("11111200103200001")
        Debug.Assert 1 = 2
    End With


End Sub

Private Sub Command7_Click()
    '测试体检管理对象。
    Dim lobj体检管理 As clsManageMedicalExam
    Dim lobjRec As Object
    Dim lcolInfo As Collection
    Dim lstrTemp As String
    
    Set lobj体检管理 = New clsManageMedicalExam
    With lobj体检管理
        Set lobjRec = .当日工作记忆簿
        '检查获取的属性。
        Debug.Assert 1 = 2
        
        Set lobjRec = .工作站配置
        '检查获取的属性。
        Debug.Assert 1 = 2
    
        Set lcolInfo = .所有业务设置
        '检查获取的属性。
        Debug.Assert 1 = 2
        
        Set lobjRec = .所有体检附加项目
        '检查获取的属性。
        Debug.Assert 1 = 2
        
        Set lobjRec = .所有体检结论条件
        '检查获取的属性。
        Debug.Assert 1 = 2
      
        Set lcolInfo = .所有体检收费标准
        '检查获取的属性。
        Debug.Assert 1 = 2
      
        lstrTemp = .业务设置("是否收费")
        '检查获取的属性。
        Debug.Assert 1 = 2
        .Sub修改业务配置 "是否收费", "是"
        .Sub修改业务配置 "是否照相", "否"
        .Sub修改业务配置 "复查周期", "30"
        .Sub修改业务配置 "是否打印体检单", "是"
        
        '检查库中：业务设置表是否已修改。
        Debug.Assert 1 = 2
        
        '添加项目[项目名称,录入标题,数据类型,数据长度,枚举值]
        Set lcolInfo = New Collection
        With lcolInfo
            .Add "身份证号", "项目名称"
            .Add "身份证号", "录入标题"
            .Add 3, "数据类型"
            .Add "20", "数据长度"
            .Add "", "枚举值"
        End With
        .Sub设置体检附加项目 1, lcolInfo
        '检查库中：体检人员附加项目设置表。
        Debug.Assert 1 = 2
        
        '修改项目。
        Set lcolInfo = New Collection
        With lcolInfo
            .Add "性别", "项目名称"
            .Add "性别", "录入标题"
            .Add 3, "数据类型"
            .Add "6", "数据长度"
            .Add "性别字典", "枚举值"
        End With
        .Sub设置体检附加项目 2, lcolInfo, "性别"
        '检查库中：体检人员附加项目设置表。
        Debug.Assert 1 = 2
        
        '删除项目。
        .Sub设置体检附加项目 3, lcolInfo, "性别"
        '检查库中：体检人员附加项目设置表。
        Debug.Assert 1 = 2
        
        Set lobjRec = .Func获取可修改的体检记录("", "", "")
        '检查获取的内容。
        Debug.Assert 1 = 2
        
        Set lobjRec = .Func获取体检结论已确定的体检记录("", "", "")
        '检查获取的内容。
        Debug.Assert 1 = 2
        Set lobjRec = .Func获取需要复查的体检记录()
        
        '检查获取的内容。
        Debug.Assert 1 = 2
        Set lobjRec = .Func获取已下结论但未确定的体检记录("", "", "", "")
        lstrTemp = .Func根据健康证条码号获取体检系统编号("")
        
        Debug.Assert 1 = 2
        
        '没有单位对外接口，此方法暂时不能测试。
        'Set lobjRec = .func单位定位()
        
        Dim lobj体检 As clsMedicalExam
        
        Set lobj体检 = New clsMedicalExam
        lstrTemp = lobj体检.Func分配系统编号
        lobj体检.系统编号 = lstrTemp
        lobj体检.体检类别 = P_EXAM_FIRST
        lobj体检.体检表.体检表名 = "从业人员体检"
        lobj体检.体检人员.姓名 = "杨洪"
        lobj体检.体检人员.单位名称 = "电机厂"
        lobj体检.体检人员.公民身份号码 = "510223470812110"
        lobj体检.体检人员.像片 = Picture1.Picture
        
        lobj体检管理.Sub体检登记 lobj体检
        '检查库中内容。
        Debug.Assert 1 = 2
        
        Dim lcolResult As Collection
        Dim lcolItem As Collection
        
        Set lcolInfo = New Collection
        lcolInfo.Add "11111200103200001"
        lcolInfo.Add "11111200103200002"
        
        '[体检项目，体检结果]
        Set lcolResult = New Collection
        Set lcolItem = New Collection
        lcolItem.Add "0001", "体检项目"
        lcolItem.Add "正常", "体检结果"
        lcolResult.Add lcolItem, lcolItem("体检项目")
        Set lcolItem = New Collection
        lcolItem.Add "0002", "体检项目"
        lcolItem.Add "正常", "体检结果"
        lcolResult.Add lcolItem, lcolItem("体检项目")
        
        .Sub填写体检结果 lcolInfo, lcolResult, "1234", Date
        '检查库中内容。
        Debug.Assert 1 = 2
        
        .Sub确定体检结论 "11111200103200002", "乙肝", "建议调离", "", "", False
        '检查库中内容。
        Debug.Assert 1 = 2
        
        .Sub取消体检结论 "11111200103200001"
        '检查库中内容。
        Debug.Assert 1 = 2
        
        '.Sub打印文书
        
    End With

End Sub

Private Sub Command8_Click()
    Dim lobj体检项目 As ClsTestItem
    Dim lobjRec As Object
    
    Set lobj体检项目 = New ClsTestItem
    With lobj体检项目
        .编码 = "0010"
        .名称 = "粪其他"
        .缺省值 = "20"
        .体检大类 = 4
        .属性 = "化验"
        .枚举来源 = ""
        .subSave
        Debug.Assert 1 = 2
        
        .subDelete
        Debug.Assert 1 = 2
    End With
    
    '测试体检项目集
    Dim lobj体检项目集 As clsTestItemSet
    Set lobj体检项目集 = New clsTestItemSet
    With lobj体检项目集
        '.体检大类 = 2
        .属性 = "化验"
        '.体检项目编码 = "0001"
        Set lobjRec = .体检项目
        Debug.Assert 1 = 2
    End With
    
End Sub

