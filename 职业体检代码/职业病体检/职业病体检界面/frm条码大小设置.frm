VERSION 5.00
Begin VB.Form frm�����С���� 
   Caption         =   "frm�����С����"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6240
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox ctxtWidth 
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox ctxtHeight 
      Height          =   270
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��  �ȣ�"
      Height          =   180
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��  �ȣ�"
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "frm�����С����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    pobjҵ�����.Sub�޸�ҵ������ "Y", ctxtHeight.Text
    pobjҵ�����.Sub�޸�ҵ������ "X", ctxtWidth.Text
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lstrWidth As String
    Dim lstrHeight As String
    
    lstrWidth = pobjҵ�����.ҵ������("X")
    lstrHeight = pobjҵ�����.ҵ������("Y")
     
    If Not (lstrWidth = "" Or IsNumeric(lstrWidth) = False) Then
        ctxtWidth = lstrWidth
    Else
        ctxtWidth = 2460
    End If
    
    If Not (lstrHeight = "" Or IsNumeric(lstrHeight) = False) Then
        ctxtHeight = lstrHeight
    Else
        ctxtHeight = 1530
    End If
End Sub
