VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   8775
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�������"
      Height          =   495
      Index           =   9
      Left            =   3240
      TabIndex        =   27
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ϵͳ�벻�ð�ť"
      Height          =   1815
      Left            =   480
      TabIndex        =   19
      Top             =   5760
      Width           =   7815
      Begin VB.CommandButton Command5 
         Caption         =   "�˳�"
         Height          =   495
         Left            =   5400
         TabIndex        =   25
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��������"
         Height          =   495
         Index           =   0
         Left            =   5400
         TabIndex        =   24
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "����嵥"
         Height          =   495
         Left            =   2760
         TabIndex        =   23
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�����¼��"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�����ӡ"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "������_ԭʼ�汾"
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   20
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��λͳ��"
      Height          =   495
      Index           =   4
      Left            =   3120
      TabIndex        =   18
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ѯͳ��"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   17
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "���¼��"
      Height          =   3855
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   7815
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   375
         Left            =   5520
         TabIndex        =   28
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�����ƽ��¼��"
         Height          =   495
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ⱦɫ�廯��ƽ��¼��"
         Height          =   495
         Index           =   11
         Left            =   5400
         TabIndex        =   16
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ι���Ӱ��ƽ��¼��"
         Height          =   495
         Index           =   10
         Left            =   2760
         TabIndex        =   15
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ѫ���滯��ƽ��¼��"
         Height          =   495
         Index           =   9
         Left            =   5400
         TabIndex        =   14
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�򳣹滯��ƽ��¼��"
         Height          =   495
         Index           =   8
         Left            =   5400
         TabIndex        =   13
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ĵ�ƽ��¼��"
         Height          =   495
         Index           =   7
         Left            =   2760
         MaskColor       =   &H8000000F&
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "������ƽ��¼��"
         Height          =   495
         Index           =   6
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "���߿ƽ��¼��"
         Height          =   495
         Index           =   5
         Left            =   5400
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "��ƽ��¼��"
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ڿƽ��¼��"
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "��ٿƽ��¼��"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X��Ӱ��ƽ��¼��"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "B��Ӱ��ƽ��¼��"
         Height          =   495
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   2520
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ܼ��߸�����Ϣ¼���"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ҵ������"
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���Ǽ�"
      Height          =   495
      Index           =   8
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ս���¼��"
      Height          =   495
      Index           =   2
      Left            =   5880
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
Public mobjUI As Object
Private mobj���������� As Object

Private Sub Command1_Click(Index As Integer)
    On Error GoTo errHandler
    mobj����������.funcStart "ְҵ�����_" & Command1(Index).Caption
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command1_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("ְҵ������.clsmageconfform_zyb")
    lobj.funcStart "ְҵ�����_" & Command2(Index).Caption
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command2_Click", Err.Number, Err.Description, False

End Sub

Private Sub Command3_Click()
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("ְҵ��ʷ¼��.clscareerhstmage")
    lobj.funcStart "ְҵ�����_" & Command3.Caption
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command3_Click", Err.Number, Err.Description, False

End Sub

Private Sub Command4_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobj As Object
    Set lobj = CreateObject("ְҵ�������¼��.clscommon")
    lobj.funcStart "ְҵ�����_" & Command4(Index).Caption
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command4_Click", Err.Number, Err.Description, False

End Sub


Private Sub Command5_Click()
    End
End Sub

'''''''''''''''����嵥���Դ��壬������ɾ��
Private Sub Command6_Click()
    Dim lobj As Object
    Set lobj = CreateObject("ְҵ������.cls����")
    lobj.funcԤ������嵥
End Sub

Private Sub Command7_Click()
frmshenghuashow.Show 1
End Sub

Private Sub Form_Load()
    Dim lstrServer As String
    Dim lstrData As String
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    lstrData = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")
''    lstrServer = "192.168.1.104"
'    lstrServer = "192.168.0.186"
'    lstrServer = "ROMAN-T43"
'    lstrData = "jk2006"
''    lstrServer = "CDMBP-CFD6FB023"
''    lstrData = "jk2006"
''    lstrData = "TEST1"

    '��ʼ�����ݷ��ʶ���
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
   

    'У����ݡ�
    If Not umfuncУ�����("0001", "") Then
        sffuncMsg "У�����ʧ�ܣ�", sf����
    End If
    
     Set mobj���������� = CreateObject("ְҵ������.clsManageTestForm")
    
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
