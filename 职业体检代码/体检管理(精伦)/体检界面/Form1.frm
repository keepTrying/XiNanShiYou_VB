VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5700
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ҵ������"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���Ǽ�"
      Height          =   495
      Index           =   8
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����ӡ"
      Height          =   495
      Index           =   7
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������¼��"
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����¼��"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1080
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
    mobj����������.funcStart "������_" & Command1(Index).Caption
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command1_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("������ý���.clsManageConfigureForm")
    lobj.funcStart "������_" & Command2(Index).Caption
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command1_Click", Err.Number, Err.Description, False

End Sub

Private Sub Form_Load()
    Dim lstrServer As String
    Dim lstrData As String
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    lstrData = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")
    lstrServer = "."
    lstrData = "jkz2006"
    
    '��ʼ�����ݷ��ʶ���
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
    Set mobj���������� = CreateObject("������.clsManageTestForm")
    
    
    'У����ݡ�
    If Not umfuncУ�����("0001", "") Then
        sffuncMsg "У�����ʧ�ܣ�", sf����
    End If
    
    '���Ի�ȡ������Ϣ��
    Dim lstrTemp  As String
   ' lstrTemp = mobj����������.Func��ȡ������Ϣ("����Ǽ�")
    
    'lstrTemp = mobj����������.Func��ȡ������Ϣ("")
    
    Exit Sub
    
errHandler:
    sfsub������ "����1", "Form1", "Form_Load", Err.Number, Err.Description, False
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
