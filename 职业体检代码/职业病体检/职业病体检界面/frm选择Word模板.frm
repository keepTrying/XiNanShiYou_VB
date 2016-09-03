VERSION 5.00
Begin VB.Form frm选择Word模板 
   Caption         =   "请选择用于生成报告的Word模板"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   5160
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取  消"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "确  定"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.ListBox clstFile 
      Height          =   4020
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "可选的模板文件："
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "frm选择Word模板"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrFilename As String
Public pstrWordname As String

Private Sub ccmdCancel_Click()
    pstrFilename = ""
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    If clstFile.ListIndex >= 0 Then
        pstrFilename = clstFile.Text
    Else
        pstrFilename = ""
    End If
    Unload Me
End Sub

Private Sub clstFile_DblClick()
    ccmdOk_Click
End Sub

Private Sub Form_Load()
    Dim lstrFile As String
    Dim lobjRec As Object
    
    lstrFile = Dir(App.Path & "\通用_*.dot")
    Do While lstrFile <> ""
       clstFile.AddItem lstrFile
       lstrFile = Dir
    Loop
    '寻找当前用户所属科室所对应的专用模板名前缀
    Set lobjRec = dafuncGetData("select 描述 from 系统管理_科室字典表 where 编号='" & um用户所属科室编号 & "'")
'    If lobjRec(0) <> "" Then
'        lstrFile = Dir(App.Path & "\" & lobjRec(0) & "_四川省" & Left(pstrWordname, 2) & "*.dot")
        lstrFile = Dir(App.Path & "\职业病体检_四川省" & Left(pstrWordname, 2) & "*.dot")
'    lstrFile = Dir(App.Path & "\职业病体检_四川省*.dot")
'        MsgBox lstrFile
'        Do While lstrFile <> ""
           clstFile.AddItem lstrFile
'           lstrFile = Dir
'        Loop
    
'    End If
    If clstFile.ListCount = 0 Then
        MsgBox "没有找到Word模板文件！", vbInformation, "系统提示"
        pstrFilename = ""
        'Unload Me
'    ElseIf clstFile.ListCount = 1 Then
'        pstrFilename = lstrFile
'        Unload Me
    Else
    
        pstrFilename = lstrFile
'        clstFile.ListIndex = 0
    End If
End Sub
