VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frm打印水晶报表 
   Caption         =   "正在预览文书..."
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin Crystal.CrystalReport cRepPrint 
      Left            =   0
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton ccmdClose 
      Caption         =   "关闭(&C)"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frm打印水晶报表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'cRepPrint
Private Sub ccmdClose_Click()
    Unload Me
End Sub

