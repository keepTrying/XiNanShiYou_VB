VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm测试界面 
   Caption         =   "测试界面"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Output"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "InputFromMdb"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "InputFromExcel"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frm测试界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj业务对象 As Object

Private Sub Command1_Click()
    mobj业务对象.funcStart "外单位数据导入"
End Sub

Private Sub Command2_Click()
    mobj业务对象.funcStart "内部数据导入"
End Sub

Private Sub Command3_Click()
    mobj业务对象.funcStart "内部数据导出"
End Sub

Private Sub Form_Load()
    Dim lstrServer  As String
    Dim lstrData As String
    
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")

    '初始化数据访问对象。
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
    'umsub数据导入 "c:\a.mdb", False, ProgressBar1
    Set mobj业务对象 = CreateObject("体检对外接口部件.clsManageTransmission")
    
    If Not umfunc校验身份("7612", "") Then
        sffuncMsg "校验身份失败。", sf警告
        End
    End If
    
End Sub
