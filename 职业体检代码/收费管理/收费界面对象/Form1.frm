VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5610
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ŷ�����"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ҵ������"
      Height          =   495
      Index           =   5
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�շѹ���"
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobj���������� As Object

Private Sub Command1_Click(Index As Integer)
    On Error GoTo errHandler
    mobj����������.funcStart "�շѹ���_" & Command1(Index).Caption
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command1_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command2_Click()
    Dim lobj�ӿ� As Object
    Dim lstrTemp As String
    Dim lstr�շѱ�� As String
    Dim lstr��λ��� As String
    Dim lcol As Collection
    On Error GoTo errHandler
    
    'mobj����������.funcStart "�շѹ���_����"
    Set lobj�ӿ� = CreateObject("�շѽӿڶ���.cls����ӿ�")
    lstr�շѱ�� = "00102062400012"
    lstr��λ��� = "0000000017"
    lstrTemp = lobj�ӿ�.func����_���ݼ���(lcol, lstr�շѱ��, True, lstr��λ���, "�°�֤�շ�")
    
    Exit Sub
errHandler:
    MsgBox Error, vbOKOnly + vbInformation, "ϵͳ��ʾ"
End Sub

Private Sub Form_Load()
    Dim lstrServer As String
    Dim lstrData As String
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    lstrData = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")
    
    On Error Resume Next
    Dim lstrError As String
    Dim i As Long
    i = 0
retry:    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    If Err <> 0 And i < 3 Then
        '���ԡ�
        Err.Clear
        i = i + 1
        GoTo retry
    End If
    lstrError = Error
    On Error GoTo errHandler
    If lstrError <> "" Then
        Err.Raise 6666, , "��ʼ�����ݷ��ʶ���ʧ�ܣ�" & lstrError
    End If
    
    Set mobj���������� = CreateObject("�շѽ��沿��.cls�������")
    
    
    'У����ݡ�
    If Not umfuncУ�����("0001", "") Then
        sffuncMsg "У�����ʧ�ܣ�", sf����
    End If
    
    
    Exit Sub
    
errHandler:
    sfsub������ "����1", "Form1", "Form_Load", Err.Number, Err.Description, False
    End
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    On Error Resume Next
    Set mobj���������� = Nothing
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Form_Unload", Err.Number, Err.Description, False
End Sub
