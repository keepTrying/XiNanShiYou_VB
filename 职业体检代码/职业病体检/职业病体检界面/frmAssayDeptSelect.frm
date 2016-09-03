VERSION 5.00
Begin VB.Form frmAssayDeptSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "打印试管标签"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "选择标签类型"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox cchkType 
         Caption         =   "肝功2，肾功，GLU，血脂，ACP"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton ccmdCancel 
         Caption         =   "取消"
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton ccmdConfirm 
         Caption         =   "确定"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "染色体"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "肝功1，肾功，两对半，GLU，血脂，ACP"
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "免疫.血清"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "尿常规"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "血常规.静脉血"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAssayDeptSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-08-19 于登淼
'选择打印标签类型，点击“确定”后，即可打印标签

Option Explicit
Public pblnOk As Boolean
Public selectedDeptName As Collection

Private Sub ccmdCancel_Click()
    pblnOk = False
    subExit
End Sub

Private Sub ccmdConfirm_Click()
    Dim i As Integer
    
    Set selectedDeptName = New Collection
    selectedDeptName.Add ""
    For i = 0 To cchkType.Count - 1
        If cchkType(i).Value = 1 Then
            selectedDeptName.Add cchkType(i).Caption
        End If
    Next i
    
    pblnOk = True
    subExit
End Sub

Sub subExit()
    'End
    Unload frmAssayDeptSelect
End Sub
