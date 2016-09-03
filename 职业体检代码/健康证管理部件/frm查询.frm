VERSION 5.00
Begin VB.Form frm查询 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "健康体检查询"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4230
   Icon            =   "frm查询.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox ccmb发证单位 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frm查询.frx":0E42
      Left            =   1320
      List            =   "frm查询.frx":0E4C
      TabIndex        =   6
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox ctxtUnit 
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox ccmbType 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox ctxtEndDate 
      Height          =   270
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox ctxtStartDate 
      Height          =   270
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox ctxtName 
      Height          =   270
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox ctxtNo 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检单位："
      Height          =   180
      Index           =   10
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位名称："
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "种    类："
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "到："
      Height          =   180
      Index           =   3
      Left            =   960
      TabIndex        =   12
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检日期从："
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓    名："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体检号："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frm查询"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrNo As String
Public pstrName As String
Public pstrUnit As String
Public pstrStartDate As String
Public pstrEndDate As String
Public pstrType As String
Public pstr发证单位 As String

Public pblnOk As Boolean


Private Sub ccmdCancel_Click()
    pblnOk = False
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    On Error Resume Next
    pstrNo = ctxtNo.Text
    pstrName = ctxtName.Text
    pstrUnit = ctxtUnit.Text
    pstrStartDate = ctxtStartDate.Text
    pstrEndDate = ctxtEndDate.Text
    pstrType = ccmbType.Text
    pstr发证单位 = ccmb发证单位.Text
    
    pblnOk = True
    Unload Me
End Sub

'功能：控制不能输入单印号，处理回车。
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        SendKeys Chr(9)
    ElseIf KeyCode = 39 Then
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    Dim lcolInfo As Collection
    Dim i As Long
    
    On Error GoTo errhandler
    
    '获取种类。
    Set lcolInfo = pobj记忆.记忆项值("卫生种类", True)
    ccmbType.Clear
    ccmbType.AddItem ""
    For i = 1 To lcolInfo.Count
        ccmbType.AddItem lcolInfo(i)
    Next
    ccmbType.ListIndex = 0
    
    '获取发证单位。
    Set lcolInfo = pobj记忆.记忆项值("发证单位", True)
    ccmb发证单位.Clear
    ccmb发证单位.AddItem ""
    For i = 1 To lcolInfo.Count
        ccmb发证单位.AddItem lcolInfo(i)
    Next
    ccmb发证单位.ListIndex = 0
    
    'ctxtStartDate = Format(DateAdd("d", -30, Date), "yyyy-mm-dd")
    'ctxtEndDate = Format(Date, "yyyy-mm-d")
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm业务设置", "Form_Load", Err.Number, Err.Description, False
    
End Sub
