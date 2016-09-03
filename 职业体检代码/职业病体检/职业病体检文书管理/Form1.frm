VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "预览"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin Crystal.CrystalReport cRepPrint 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    cRepPrint.Connect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=user26;PWD = welcome;Initial Catalog=防疫2001;Data Source=testserver"
    
    '卫生监测管理_公卫现场监测报告
    cRepPrint.ReportFileName = "C:\Program Files\卫生防疫管理信息系统\公用组件\卫生监测管理_公卫现场监测报告.rpt"

    '设置打印参数
    cRepPrint.Formulas(0) = "监测系统编号='1200100006'"
    
    
    '显示报表
    cRepPrint.WindowState = crptMaximized
    cRepPrint.WindowControlBox = True
    cRepPrint.WindowLeft = 0
    cRepPrint.WindowTop = 0
    cRepPrint.Destination = crptToWindow
    cRepPrint.Action = 1
    Visible = True

End Sub

