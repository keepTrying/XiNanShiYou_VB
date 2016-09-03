VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4095
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "业务设置"
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "健康证管理"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   0
      Top             =   120
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
    mobj界面管理对象.funcStart "健康证管理_" & Command1(Index).Caption
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
    
    'Internet版测试。
    'lstrServer = "sql01\bloodbank"
    'lstrData = "防疫2001"
    
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
    
    Set mobj界面管理对象 = CreateObject("健康证管理部件.clsManageForm")
    
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
