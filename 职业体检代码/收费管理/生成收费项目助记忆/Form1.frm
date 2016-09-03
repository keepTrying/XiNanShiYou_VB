VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "生成收费项目助记符"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdStop 
      Caption         =   "终止(&T)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton ccmdStart 
      Caption         =   "开始(&S)"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label clblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnStop As Boolean

Private Sub ccmdStart_Click()
    Dim lstrTemp As String
    Dim lobjRec As Object
    Dim i As Long
    MousePointer = 11
    i = 0
    mblnStop = False
    ccmdStop.Enabled = True
    ccmdStart.Enabled = False
    Set lobjRec = dafuncGetData("select 收费项目编号,收费项目名称 from 收费管理_收费项目字典表 where isnull(助记符,'')=''")
    clblInfo.Caption = "共要生成：" & lobjRec.recordcount & "个项目的助记符。" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Do While Not lobjRec.EOF
        lstrTemp = guf_GetFirstLetter(lobjRec("收费项目名称"))
        lstrTemp = Left(lstrTemp, 20)
        dafuncGetData "update 收费管理_收费项目字典表 set 助记符='" & lstrTemp & "' where 收费项目编号='" & lobjRec!收费项目编号 & "'"
        i = i + 1
        clblInfo.Caption = "共要生成：" & lobjRec.recordcount & "个项目的助记符。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                & "项目：" & lobjRec("收费项目名称") & ", 助记符：" & lstrTemp & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                & "还剩下" & (lobjRec.recordcount - i) & "个。"
        
        DoEvents
        If mblnStop Then Exit Do
        
        lobjRec.movenext
    Loop
    
    ccmdStop.Enabled = False
    ccmdStart.Enabled = True
    MousePointer = 0
End Sub

Private Sub ccmdStop_Click()
    mblnStop = True
    
End Sub

Private Sub Form_Load()
    Dim lstrTemp As String
    Dim lstrServer As String
    Dim lstrData As String
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    
    On Error Resume Next
    Dim lstrError As String
    Dim i As Long
    i = 0
retry:    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    If Err <> 0 And i < 3 Then
        '重试。
        Err.Clear
        i = i + 1
        GoTo retry
    End If
    lstrError = Error
    On Error GoTo errHandler
    If lstrError <> "" Then
        Err.Raise 6666, , "初始化数据访问对象失败！" & lstrError
    End If
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbOKOnly
    End
    Exit Sub
    Resume
    
End Sub
