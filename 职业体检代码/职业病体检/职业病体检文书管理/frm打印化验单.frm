VERSION 5.00
Object = "{D0044432-16F0-11D5-8F5A-0050BA637F0B}#2.3#0"; "DyBigCheck.ocx"
Begin VB.Form frm��ӡ���鵥 
   Caption         =   "��ӡ���鵥"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7545
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ccmdClose 
      Caption         =   "�ر� (&C)"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "��ӡ (&P)"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
   Begin dyBigCheck.ctlDyBigCheck c���鵥 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11668
      BackColor       =   0
      FontSize        =   9
      �Թܱ��ÿ�ж���=   6
      ������Ŀÿ�ж���=   8
   End
End
Attribute VB_Name = "frm��ӡ���鵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��

Public pobj���鵥���� As Object

Private Sub ccmdClose_Click()
    On Error Resume Next
    Unload Me
End Sub

'���ܣ���ӡ���鵥��
'���ߣ��
Private Sub ccmdPrint_Click()
    On Error GoTo errHandler
    c���鵥.subPrint
    
    Unload Me
    Exit Sub
errHandler:
    sfsub������ "����������", "frm��ӡ���鵥", "ccmdPrint_Click", Err.Number, Err.Description, False
    
End Sub

'���ܣ������ʼ��������ʾ���鵥���ݡ�
'���ߣ��
Private Sub Form_Load()
    On Error GoTo errHandler
    If pobj���鵥���� Is Nothing Then
        Err.Raise 6666, , "����������ǰ���������û��鵥�������ԡ�"
    Else
        Set c���鵥.���鵥���� = pobj���鵥����
        c���鵥.subShow
    End If
    Exit Sub
errHandler:
    sfsub������ "����������", "frm��ӡ���鵥", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

