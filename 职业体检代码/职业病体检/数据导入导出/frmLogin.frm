VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00F0F0EC&
   BorderStyle     =   0  'None
   Caption         =   "�û���¼"
   ClientHeight    =   2520
   ClientLeft      =   1710
   ClientTop       =   1950
   ClientWidth     =   5400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2520
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox ctxtUserNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1180
      Width           =   2100
   End
   Begin VB.TextBox ctxtPassWord 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      IMEMode         =   3  'DISABLE
      Left            =   1740
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1780
      Width           =   2100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�������û���źͿ����Ե�¼��Դ����������Ϣ����ϵͳ"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label ccmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label ccmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'���������û����
Private Sub ctxtpassword_Change()
    If ctxtUserNo.Text = "" Then
        Call sffuncMsg("���������û����", sf����)
        ctxtUserNo.SetFocus
    End If
End Sub

'�õ����㣬ѡ�������ַ�
Private Sub ctxtPassWord_GotFocus()
    ctxtPassWord.SelStart = 0
    ctxtPassWord.SelLength = Len(ctxtPassWord.Text)
End Sub

'�õ����㣬ѡ�������ַ�
Private Sub ctxtUserNo_GotFocus()
    ctxtUserNo.SelStart = 0
    ctxtUserNo.SelLength = Len(ctxtUserNo.Text)
End Sub

'��Ӧ�س���
Private Sub ctxtUSERNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ctxtPassWord.SetFocus
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

'��Ӧ�س���
Private Sub ctxtPAssword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '����У���û�
        Call ccmdOk_Click
    End If
End Sub

'�û�ȡ�����˳�ϵͳ���ͷŶ���
Private Sub ccmdCancel_Click()

    Unload Me
End Sub

Private Sub ccmdOk_Click()
    Dim lstrģʽ As String
    
    On Error GoTo errHandle
    Dim lfrm������ As Form
    
    '��֤���
    If umfuncУ�����(Trim(ctxtUserNo.Text), ctxtPassWord.Text) Then
        '�Ϸ��û�
        '�޸ģ�2002-1-28���Ǽ����ñ���ǰ�û���ţ���
        On Error Resume Next
        
        Set lfrm������ = New frmMain
        
        On Error GoTo errHandle
        If Trim(um�û��������ұ��) = "" Then
            If Trim(ctxtUserNo.Text) = "0000" Then
                '��ʾ������
                Unload Me
                lfrm������.Show
            Else
                Call sffuncMsg("������û�����ÿ��ң�����ʹ�ñ�ϵͳ��", sf����)
                Call ccmdCancel_Click
            End If
        Else
            Unload Me
            lfrm������.Show
        End If
        
    Else '�Ƿ��û���������
        Call sffuncMsg("�û���Ż�������", sf����)
        ctxtUserNo.SetFocus
    End If
errHandle:
    Set lfrm������ = Nothing
    If Err.Number = 0 Then Exit Sub

    Call sfsub������("������ݵ��뵼��", "frmLogin", "ccmdOk_click", Err.Number, Err.Description, False)
End Sub



'�޸ģ�2002-1-28�������ע����л�ȡ�ϴε�¼���û���š�
Private Sub Form_Load()
    Dim lstrUser As String
    
    On Error Resume Next
    lstrUser = sffuncGetSetting("ϵͳ����", "��������", "�û����")
    If lstrUser = "" Then lstrUser = "0001"
    ctxtUserNo.Text = lstrUser
    Label1 = "�������û���źͿ����Ե�¼������뵼�뵼������"
    
End Sub
