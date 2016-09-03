VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5610
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "财务监管"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "号段设置"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "划价"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "业务设置"
      Height          =   495
      Index           =   5
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "收费管理"
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobj界面管理对象 As Object

Private Sub Command1_Click(Index As Integer)
    On Error GoTo errHandler
    mobj界面管理对象.funcStart "收费管理_" & Command1(Index).Caption
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command1_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command2_Click()
    Dim lobj接口 As Object
    Dim lstrTemp As String
    Dim lstr收费编号 As String
    Dim lstr单位编号 As String
    Dim lcol As Collection
    On Error GoTo errHandler
    
    'mobj界面管理对象.funcStart "收费管理_划价"
    Set lobj接口 = CreateObject("收费接口对象.cls对外接口")
    lstr收费编号 = "00102062400012"
    lstr单位编号 = "0000000017"
    lstrTemp = lobj接口.func划价_数据集合(lcol, lstr收费编号, True, lstr单位编号, "新办证收费")
    
    Exit Sub
errHandler:
    MsgBox Error, vbOKOnly + vbInformation, "系统提示"
End Sub

Private Sub Form_Load()
    Dim lstrServer As String
    Dim lstrData As String
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    
    On Error Resume Next
    Dim lstrError As String
    Dim i As Long
    i = 0
retry:    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    If Err <> 0 And i < 3 Then
        '重试。
        Err.Clear
        i = i + 1
        GoTo retry
    End If
    lstrError = Error
    On Error GoTo errHandler
    If lstrError <> "" Then
        Err.Raise 6666, , "初始化数据访问对象失败！" & lstrError
    End If
    
    Set mobj界面管理对象 = CreateObject("收费界面部件.cls界面管理")
    
    
    '校验身份。
    If Not umfunc校验身份("0001", "") Then
        sffuncMsg "校验身份失败！", sf警告
    End If
    
    
    Exit Sub
    
errHandler:
    sfsub错误处理 "工程1", "Form1", "Form_Load", Err.Number, Err.Description, False
    End
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    On Error Resume Next
    Set mobj界面管理对象 = Nothing
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Form_Unload", Err.Number, Err.Description, False
End Sub
