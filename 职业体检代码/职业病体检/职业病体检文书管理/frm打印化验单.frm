VERSION 5.00
Object = "{D0044432-16F0-11D5-8F5A-0050BA637F0B}#2.3#0"; "DyBigCheck.ocx"
Begin VB.Form frm打印化验单 
   Caption         =   "打印化验单"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdClose 
      Caption         =   "关闭 (&C)"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "打印 (&P)"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
   Begin dyBigCheck.ctlDyBigCheck c化验单 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11668
      BackColor       =   0
      FontSize        =   9
      试管编号每行段数=   6
      化验项目每行段数=   8
   End
End
Attribute VB_Name = "frm打印化验单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春

Public pobj化验单对象 As Object

Private Sub ccmdClose_Click()
    On Error Resume Next
    Unload Me
End Sub

'功能：打印化验单。
'作者：杨春
Private Sub ccmdPrint_Click()
    On Error GoTo errHandler
    c化验单.subPrint
    
    Unload Me
    Exit Sub
errHandler:
    sfsub错误处理 "体检文书管理", "frm打印化验单", "ccmdPrint_Click", Err.Number, Err.Description, False
    
End Sub

'功能：窗体初始化，并显示化验单内容。
'作者：杨春
Private Sub Form_Load()
    On Error GoTo errHandler
    If pobj化验单对象 Is Nothing Then
        Err.Raise 6666, , "启动本界面前必须先设置化验单对象属性。"
    Else
        Set c化验单.化验单对象 = pobj化验单对象
        c化验单.subShow
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "体检文书管理", "frm打印化验单", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

