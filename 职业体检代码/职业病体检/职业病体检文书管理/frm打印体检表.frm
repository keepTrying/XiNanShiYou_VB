VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm��ӡ���� 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label clbl�Թܺ� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   90
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccMain 
      Height          =   765
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.Image cimgPhoto 
      Height          =   1845
      Left            =   8400
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frm��ӡ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pobj�������� As Object 'recordset[ϵͳ��ţ����������֤�ţ��Ա𣬵�λ����]

Private Sub Form_Load()
    Dim lstrTmp As String
    
    On Error GoTo errhandler
    
    '���������ݡ�
    cbccMain.Value = pobj��������("ϵͳ���")
    
    '���������󣬻�ȡ��Ƭ��
    Dim lobj��� As Object
    Set lobj��� = CreateObject("������.clsMedicalExam")
    lobj���.ϵͳ��� = pobj��������("ϵͳ���")
    
      
    '��ʾ��Ƭ��
    If Not lobj���.�����Ա.��Ƭ Is Nothing Then
        cimgPhoto.Picture = lobj���.�����Ա.��Ƭ
    End If
    
    Exit Sub
errhandler:
    sfsub������ "ְҵ���������", "frm��ӡ���ǼǱ�", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


