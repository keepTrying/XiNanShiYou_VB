VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm提示 
   BorderStyle     =   0  'None
   Caption         =   "frm提示"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5865
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar cprg提示 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Max             =   20
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请等待，系统正在初始化……"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2340
   End
End
Attribute VB_Name = "frm提示"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
