VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm���Խ��� 
   Caption         =   "���Խ���"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Output"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "InputFromMdb"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "InputFromExcel"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frm���Խ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjҵ����� As Object

Private Sub Command1_Click()
    mobjҵ�����.funcStart "�ⵥλ���ݵ���"
End Sub

Private Sub Command2_Click()
    mobjҵ�����.funcStart "�ڲ����ݵ���"
End Sub

Private Sub Command3_Click()
    mobjҵ�����.funcStart "�ڲ����ݵ���"
End Sub

Private Sub Form_Load()
    Dim lstrServer  As String
    Dim lstrData As String
    
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    lstrData = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")

    '��ʼ�����ݷ��ʶ���
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
    'umsub���ݵ��� "c:\a.mdb", False, ProgressBar1
    Set mobjҵ����� = CreateObject("������ӿڲ���.clsManageTransmission")
    
    If Not umfuncУ�����("7612", "") Then
        sffuncMsg "У�����ʧ�ܡ�", sf����
        End
    End If
    
End Sub
