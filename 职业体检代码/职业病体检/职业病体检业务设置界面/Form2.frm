VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7350
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ҵ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ģ������"
      Height          =   495
      Left            =   840
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

Private Sub Command1_Click()
    On Error GoTo errHandler
    mobj����������.funcStart "��������"
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command1_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler
    mobj����������.funcStart "ҵ������"
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command2_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    '��ʼ�����ݷ��ʶ������ӱ���)��
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=����2001(��ʼ��);Data Source=YANGCHUN"
    
    'Tdcserver��
    'dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=����2001;Data Source=Tdcserver"
    
    Set mobj���������� = CreateObject("���ҵ�����ý���.clsManageConfigureForm")
    
    If Not umfuncУ�����("5555", "") Then
        sffuncMsg "У�����ʧ�ܡ�", sf����
    End If
    
    Exit Sub
    
errHandler:
    sfsub������ "����1", "Form1", "Form_Load", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    On Error Resume Next
    Set mobj���������� = Nothing
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Form_Unload", Err.Number, Err.Description, False
End Sub
