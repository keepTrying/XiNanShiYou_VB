VERSION 5.00
Begin VB.Form frmPublicQueryResult 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "���ڲ�ѯ�����"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.TextBox ctxtNo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4088
      MaxLength       =   20
      TabIndex        =   0
      Top             =   6600
      Width           =   3720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ҫ��ѯ���ӡ���������ˢ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2460
      TabIndex        =   3
      Top             =   5880
      Width           =   6750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��ӭ��***"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2880
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   9105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ϣ��ѯϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   7845
   End
End
Attribute VB_Name = "frmPublicQueryResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��

Private mlngϵͳ��ų��� As Long
Private mblnInUse As Boolean

Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub ctxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 And Trim(ctxtNo) <> "" Then
        frmTestResult.ϵͳ��� = Trim(ctxtNo)
        frmTestResult.Show 1
        ctxtNo = ""
    End If
    
    Exit Sub
errHandler:
    ctxtNo = ""

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    mblnInUse = True
    
    Label1(1).Caption = "��ӭ��" & um����վ��
    
    '���������󣬻�ȡϵͳ��ŵĳ��ȡ�
    Dim lobj��� As Object
    Set lobj��� = CreateObject("������.clsMedicalExam")
    mlngϵͳ��ų��� = lobj���.ϵͳ��ų���
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmPublicQueryResult", "Form_Load", 6666, lstrError, False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub
Private Sub Form_Activate()
    On Error Resume Next
    ctxtNo = ""
    ctxtNo.TabIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mblnInUse = False
End Sub
