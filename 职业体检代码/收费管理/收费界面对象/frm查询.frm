VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm��ѯ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��ѯ"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5100
   Icon            =   "frm��ѯ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   22
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   21
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   720
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   4815
   End
   Begin VB.TextBox ctxt���ѵ�λ 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox ctxt������ 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox ctxt�վݺ� 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox ctxt�շ����� 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox Ccboҵ����� 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3480
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker cdtp��ֹ���� 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   52363265
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker cdtp��ʼ���� 
      Height          =   300
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   52363265
      CurrentDate     =   36951
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ѵ�λ"
      Height          =   180
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ֹ  ��"
      Height          =   180
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��  ��"
      Height          =   180
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҵ�����"
      Height          =   180
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   1200
      TabIndex        =   8
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ڷ�Χ"
      Height          =   180
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   720
   End
End
Attribute VB_Name = "frm��ѯ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pstr�շ����� As String
Public pstr�վݺ� As String
Public pstr��λ���� As String
Public pstr������ As String
Public pstr��ʼ���� As String
Public pstr��ֹ���� As String
Public pstrҵ����� As String

Public pblnOk As Boolean

Private Sub Clabҵ�����_Click()

End Sub

Private Sub ccmdCancel_Click()
    pblnOk = False
    Unload Me
   
End Sub

Private Sub ccmdOk_Click()
    If cchkType(0).Value = 1 Then
        pstr�շ����� = ctxt�շ�����.Text
    Else
        pstr�շ����� = ""
    End If
    If cchkType(1).Value = 1 Then
        pstr�վݺ� = ctxt�վݺ�.Text
    Else
        pstr�վݺ� = ""
    End If
    If cchkType(2).Value = 1 Then
        pstr������ = ctxt������.Text
    Else
        pstr������ = ""
    End If
    If cchkType(3).Value = 1 Then
        pstr��λ���� = ctxt���ѵ�λ.Text
    Else
        pstr��λ���� = ""
    End If
    If cchkType(4).Value = 1 Then
        pstr��ʼ���� = Format(cdtp��ʼ����.Value, "yyyy-mm-dd")
        pstr��ֹ���� = Format(cdtp��ֹ����.Value, "yyyy-mm-dd")
    Else
        pstr��ʼ���� = ""
        pstr��ֹ���� = ""
    End If
    If cchkType(5).Value = 1 Then
        pstrҵ����� = Ccboҵ�����.Text
    Else
        pstrҵ����� = ""
    End If
    
    pblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    cdtp��ʼ����.Value = Date               '��ʼ����ʼ���������Ϊ��������
    cdtp��ֹ����.Value = Date               '��ʼ���������������Ϊ��������

    Dim lobjRec As Object
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData("select ��Ӧҵ�� from �շѹ���_������Ϣ�� where isnull(��Ӧҵ��,'')<>'' group by ��Ӧҵ��  order by ��Ӧҵ�� ")
        
    Ccboҵ�����.Clear
        
    Ccboҵ�����.AddItem ""
        
    Do While Not lobjRec.EOF
        Ccboҵ�����.AddItem lobjRec("��Ӧҵ��").Value
        lobjRec.MoveNext
    Loop
    
    Ccboҵ�����.ListIndex = 0
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm��ѯ", "Form_Load", Err.Number, Err.Description, False
End Sub
