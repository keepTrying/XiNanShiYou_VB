VERSION 5.00
Begin VB.Form formdct 
   BackColor       =   &H00FFFFFF&
   Caption         =   "��������"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10515
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton btnright 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�鿴�Ҷ����"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton btnleft 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�鿴������"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   360
      ScaleHeight     =   4335
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   600
      Width           =   8835
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������O"
      Height          =   225
      Left            =   7080
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ǵ���X"
      Height          =   225
      Left            =   8400
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ȣ�"
      Height          =   225
      Left            =   6000
      TabIndex        =   24
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   7800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8400
      X2              =   9120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ʽ��"
      Height          =   225
      Left            =   6000
      TabIndex        =   23
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   495
      Left            =   1680
      TabIndex        =   22
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ע��"
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label���� 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���/�Ҷ�"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ʾͼƬΪ��"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�������ԣ�"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "˹�ϣ�"
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   7200
      Width           =   3375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Τ�ϣ�"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ϣ�"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "        ��          ��"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ԣ�"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ԣ�"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ʽ��"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ϵͳ��ţ�"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "formdct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim mpicPhoto As StdPicture
 Private Sub btnleft_Click()
 Set mpicPhoto = pmfunc��ȡͼƬ(Label2.Caption & "left", "ְҵ�����")
 Picture1.Picture = mpicPhoto
 Label����.Caption = "���"
End Sub

Private Sub btnright_Click()
 Set mpicPhoto = pmfunc��ȡͼƬ(Label2.Caption & "right", "ְҵ�����")
 Picture1.Picture = mpicPhoto
 Label����.Caption = "�Ҷ�"
End Sub

