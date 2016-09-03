VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm打印试管标签 
   BackColor       =   &H00FFFFFF&
   Caption         =   "打印试管标签"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label clblName 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCode 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2775
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
End
Attribute VB_Name = "frm打印试管标签"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-06-20 于登淼
'打印试管标签
'根据省疾控要求，试管标签打印成贴纸，每次两张

Option Explicit

Public sysNo As String
Public paraName As String
Public paraSex As String
Public paraAge As String
Public paraDeptName As String

Public Function PrintLabel()
    Dim lstrWidth As String
    Dim lstrHeight As String
    Dim pobj业务对象 As Object
    
    Set pobj业务对象 = CreateObject("职业病对象.clsManageMedicalExam")
     
     lstrWidth = pobj业务对象.业务设置("X")
     lstrHeight = pobj业务对象.业务设置("Y")
     
    If Not (lstrWidth = "" Or IsNumeric(lstrWidth) = False) Then
        Me.Width = CLng(lstrWidth)
    Else
        Me.Width = 2460
    End If
    
    If Not (lstrHeight = "" Or IsNumeric(lstrHeight) = False) Then
        Me.Height = CLng(lstrHeight)
    Else
        Me.Height = 1530
    End If
    BarCode.Value = sysNo
   
    
'    clblName.Caption = Trim(paraName) + IIf(Len(paraDeptName) = 0, "", "（" & paraDeptName & "）")
    '加入性别和年龄 2015-12-25 by 牟俊
    clblName.Caption = Trim(paraName) + Trim(paraAge) + Trim(paraSex) + IIf(Len(paraDeptName) = 0, "", "（" & paraDeptName & "）")
       '设置打印机名
    Dim devPrinter As Printer
     For Each devPrinter In Printers
        If devPrinter.DeviceName = "标签打印机" Then
           '设定为系统缺省打印机。
          Set Printer = devPrinter
           ' 终止查找打印机。
        Me.PrintForm
        'modify by 2015-12-28
        
         Exit For
         End If
        Next
         
          
'          Dim tips
'         tips = MsgBox("没有设置标签打印机", vbOKOnly + vbCritical, "错误")
'        End If
'     Next
     
   
End Function

