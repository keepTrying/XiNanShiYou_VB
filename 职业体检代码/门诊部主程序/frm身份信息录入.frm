VERSION 5.00
Begin VB.Form frm�����Ϣ¼�� 
   Caption         =   "�����Ϣ¼��"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5040
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ  ��"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1740
      Width           =   975
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "��  ֤"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1740
      Width           =   975
   End
   Begin VB.TextBox ctxtCode 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox ctxtNo 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2���û���š���֤����������ṩ����ȡ��"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   3510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "1����֤ʱ�뱣֤�������ܷ���Internet�������޷���֤��"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   4590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��֤�룺"
      Height          =   180
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�û���ţ�"
      Height          =   180
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frm�����Ϣ¼��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ccmdCancel_Click()
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    ctxtCode = Trim(ctxtCode)
    ctxtNo = Trim(ctxtNo)
    If ctxtNo = "" Then
        MsgBox "���������û���ţ�", vbInformation, "ϵͳ��ʾ"
        ctxtNo.SetFocus
        Exit Sub
    End If
    If ctxtCode = "" Then
        MsgBox "����������֤�룡", vbInformation, "ϵͳ��ʾ"
        ctxtCode.SetFocus
        Exit Sub
    End If
    
    Dim lobjServer As New cls�û����
    Dim lobjRec As Recordset
    
    On Error GoTo errHandle
    
    If lobjServer.funcCheckUser(ctxtNo, ctxtCode) = False Then
        MsgBox "��֤��Ϣ����ȷ�����������룡", vbInformation, "ϵͳ��ʾ"
        ctxtNo.SetFocus
        Exit Sub
    End If
    Unload Me
    Exit Sub
errHandle:
    MsgBox "�û���֤ʧ�ܣ����鱾�����Ƿ����������ʻ�������" + Error, vbCritical, "ϵͳ����"
End Sub

Private Sub ctxtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ccmdOk.SetFocus

End Sub

Private Sub ctxtNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ctxtCode.SetFocus
End Sub
