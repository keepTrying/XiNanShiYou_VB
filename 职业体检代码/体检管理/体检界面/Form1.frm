VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5700
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "体检表设置"
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "业务设置"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "体检登记"
      Height          =   495
      Index           =   8
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "文书打印"
      Height          =   495
      Index           =   7
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "体检结论录入"
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "体检结果录入"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1080
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
    mobj界面管理对象.funcStart "体检管理_" & Command1(Index).Caption
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command1_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("体检设置界面.clsManageConfigureForm")
    lobj.funcStart "体检管理_" & Command2(Index).Caption
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command1_Click", Err.Number, Err.Description, False

End Sub

Private Sub Form_Load()
    Dim lstrServer As String
    Dim lstrData As String
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    lstrServer = "."
    lstrData = "jkz2006"
    
    '初始化数据访问对象。
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
    Set mobj界面管理对象 = CreateObject("体检界面.clsManageTestForm")
    
    
    '校验身份。
    If Not umfunc校验身份("0001", "") Then
        sffuncMsg "校验身份失败！", sf警告
    End If
    
    '测试获取助推信息。
    Dim lstrTemp  As String
   ' lstrTemp = mobj界面管理对象.Func获取主推信息("复查登记")
    
    'lstrTemp = mobj界面管理对象.Func获取主推信息("")
    
    Exit Sub
    
errHandler:
    sfsub错误处理 "工程1", "Form1", "Form_Load", Err.Number, Err.Description, False
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
