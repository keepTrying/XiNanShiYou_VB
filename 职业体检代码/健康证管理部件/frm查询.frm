VERSION 5.00
Begin VB.Form frm��ѯ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������ѯ"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4230
   Icon            =   "frm��ѯ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox ccmb��֤��λ 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frm��ѯ.frx":0E42
      Left            =   1320
      List            =   "frm��ѯ.frx":0E4C
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
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ��(&O)"
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
      Caption         =   "��쵥λ��"
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
      Caption         =   "��λ���ƣ�"
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
      Caption         =   "��    �ࣺ"
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
      Caption         =   "����"
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
      Caption         =   "������ڴӣ�"
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
      Caption         =   "��    ����"
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
      Caption         =   "���ţ�"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frm��ѯ"
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
Public pstr��֤��λ As String

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
    pstr��֤��λ = ccmb��֤��λ.Text
    
    pblnOk = True
    Unload Me
End Sub

'���ܣ����Ʋ������뵥ӡ�ţ�����س���
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
    
    '��ȡ���ࡣ
    Set lcolInfo = pobj����.������ֵ("��������", True)
    ccmbType.Clear
    ccmbType.AddItem ""
    For i = 1 To lcolInfo.Count
        ccmbType.AddItem lcolInfo(i)
    Next
    ccmbType.ListIndex = 0
    
    '��ȡ��֤��λ��
    Set lcolInfo = pobj����.������ֵ("��֤��λ", True)
    ccmb��֤��λ.Clear
    ccmb��֤��λ.AddItem ""
    For i = 1 To lcolInfo.Count
        ccmb��֤��λ.AddItem lcolInfo(i)
    Next
    ccmb��֤��λ.ListIndex = 0
    
    'ctxtStartDate = Format(DateAdd("d", -30, Date), "yyyy-mm-dd")
    'ctxtEndDate = Format(Date, "yyyy-mm-d")
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frmҵ������", "Form_Load", Err.Number, Err.Description, False
    
End Sub
