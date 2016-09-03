VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   8775
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "报告管理"
      Height          =   495
      Index           =   9
      Left            =   3240
      TabIndex        =   27
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "老系统与不用按钮"
      Height          =   1815
      Left            =   480
      TabIndex        =   19
      Top             =   5760
      Width           =   7815
      Begin VB.CommandButton Command5 
         Caption         =   "退出"
         Height          =   495
         Left            =   5400
         TabIndex        =   25
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "体检表设置"
         Height          =   495
         Index           =   0
         Left            =   5400
         TabIndex        =   24
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "体检清单"
         Height          =   495
         Left            =   2760
         TabIndex        =   23
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "体检结果录入"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "文书打印"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "体检结论_原始版本"
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   20
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "单位统计"
      Height          =   495
      Index           =   4
      Left            =   3120
      TabIndex        =   18
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询统计"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   17
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "结果录入"
      Height          =   3855
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   7815
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   375
         Left            =   5520
         TabIndex        =   28
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "生化科结果录入"
         Height          =   495
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "染色体化验科结果录入"
         Height          =   495
         Index           =   11
         Left            =   5400
         TabIndex        =   16
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "肺功能影像科结果录入"
         Height          =   495
         Index           =   10
         Left            =   2760
         TabIndex        =   15
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "血常规化验科结果录入"
         Height          =   495
         Index           =   9
         Left            =   5400
         TabIndex        =   14
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "尿常规化验科结果录入"
         Height          =   495
         Index           =   8
         Left            =   5400
         TabIndex        =   13
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "心电科结果录入"
         Height          =   495
         Index           =   7
         Left            =   2760
         MaskColor       =   &H8000000F&
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "电测听科结果录入"
         Height          =   495
         Index           =   6
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "免疫科结果录入"
         Height          =   495
         Index           =   5
         Left            =   5400
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "外科结果录入"
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "内科结果录入"
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "五官科结果录入"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X光影像科结果录入"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "B超影像科结果录入"
         Height          =   495
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   2520
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "受检者个人信息录入科"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "业务设置"
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "体检登记"
      Height          =   495
      Index           =   8
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最终结论录入"
      Height          =   495
      Index           =   2
      Left            =   5880
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
Public mobjUI As Object
Private mobj界面管理对象 As Object

Private Sub Command1_Click(Index As Integer)
    On Error GoTo errHandler
    mobj界面管理对象.funcStart "职业病体检_" & Command1(Index).Caption
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command1_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("职业病设置.clsmageconfform_zyb")
    lobj.funcStart "职业病体检_" & Command2(Index).Caption
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command2_Click", Err.Number, Err.Description, False

End Sub

Private Sub Command3_Click()
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("职业病史录入.clscareerhstmage")
    lobj.funcStart "职业病体检_" & Command3.Caption
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command3_Click", Err.Number, Err.Description, False

End Sub

Private Sub Command4_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("职业病体检结果录入.clscommon")
    lobj.funcStart "职业病体检_" & Command4(Index).Caption
    Exit Sub
errHandler:
    sfsub错误处理 "工程1", "Form1", "Command4_Click", Err.Number, Err.Description, False

End Sub


Private Sub Command5_Click()
    End
End Sub

'''''''''''''''体检清单测试窗体，用完后会删掉
Private Sub Command6_Click()
    Dim lobj As Object
    Set lobj = CreateObject("职业病文书.cls文书")
    lobj.func预览体检清单
End Sub

Private Sub Command7_Click()
frmshenghuashow.Show 1
End Sub

Private Sub Form_Load()
    Dim lstrServer As String
    Dim lstrData As String
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
''    lstrServer = "192.168.1.104"
'    lstrServer = "192.168.0.186"
'    lstrServer = "ROMAN-T43"
'    lstrData = "jk2006"
''    lstrServer = "CDMBP-CFD6FB023"
''    lstrData = "jk2006"
''    lstrData = "TEST1"

    '初始化数据访问对象。
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
   

    '校验身份。
    If Not umfunc校验身份("0001", "") Then
        sffuncMsg "校验身份失败！", sf警告
    End If
    
     Set mobj界面管理对象 = CreateObject("职业病界面.clsManageTestForm")
    
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
