VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm打印体检表 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label clbl试管号 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   90
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccMain 
      Height          =   765
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.Image cimgPhoto 
      Height          =   1845
      Left            =   8400
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frm打印体检表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pobj文书内容 As Object 'recordset[系统编号，姓名，身份证号，性别，单位名称]

Private Sub Form_Load()
    Dim lstrTmp As String
    
    On Error GoTo errhandler
    
    '填文书内容。
    cbccMain.Value = pobj文书内容("系统编号")
    
    '创建体检对象，获取照片。
    Dim lobj体检 As Object
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    lobj体检.系统编号 = pobj文书内容("系统编号")
    
      
    '显示像片。
    If Not lobj体检.体检人员.像片 Is Nothing Then
        cimgPhoto.Picture = lobj体检.体检人员.像片
    End If
    
    Exit Sub
errhandler:
    sfsub错误处理 "职业病文书管理", "frm打印体检登记表", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


