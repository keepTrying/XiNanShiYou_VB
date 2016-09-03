VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   1725
   ClientTop       =   1965
   ClientWidth     =   5550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   5550
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox ctxtUserNo 
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
      Left            =   2415
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1560
      Width           =   2100
   End
   Begin VB.TextBox ctxtPassWord 
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
      Left            =   2415
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2115
      Width           =   2100
   End
   Begin VB.Image ccmdCancel 
      Height          =   300
      Left            =   3165
      Picture         =   "frmLogin.frx":0000
      Top             =   2745
      Width           =   945
   End
   Begin VB.Image ccmdOk 
      Height          =   300
      Left            =   1290
      Picture         =   "frmLogin.frx":27B9
      Top             =   2745
      Width           =   945
   End
   Begin VB.Label clblSysName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      TabIndex        =   5
      Top             =   885
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���¼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3960
      TabIndex        =   4
      Top             =   210
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    �"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1155
      TabIndex        =   3
      Top             =   2115
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���ţ�"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1185
      TabIndex        =   2
      Top             =   1620
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   -60
      Picture         =   "frmLogin.frx":4FC7
      Top             =   0
      Width           =   5610
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   -60
      Picture         =   "frmLogin.frx":8F65
      Stretch         =   -1  'True
      Top             =   645
      Width           =   5610
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
    Set pobjƽ̨�ṹ = Nothing
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    On Error GoTo errHandle
    'Dim lfrm������ As New frm������
    Dim lfrm������ As New frmMain
    Dim i As Long
    
    '��֤���
    If umfuncУ�����(Trim(ctxtUserNo.Text), ctxtPassWord.Text) Then
        '�Ϸ��û�
        '�޸ģ�2002-1-28���Ǽ����ñ���ǰ�û���ţ���
        On Error Resume Next
        sfsubSaveSetting "ϵͳ����", "��������", "�û����", Trim(ctxtUserNo.Text)
        
        On Error GoTo errHandle
        If func�Ƿ�ע��(Trim(ctxtUserNo.Text)) Then
            If Trim(um�û��������ұ��) = "" Then
                If Trim(ctxtUserNo.Text) = "0000" Then
                    '��ʾ������
                    Unload Me
                    lfrm������.Show
                    plngMainHwnd = lfrm������.hWnd
                Else
                    Call sffuncMsg("������û�����ÿ��ң�����ʹ�ñ�ϵͳ��", sf����)
                    Call ccmdCancel_Click
                End If
            Else
                Unload Me
                lfrm������.Show
                plngMainHwnd = lfrm������.hWnd
            End If
        Else
            If Trim(ctxtUserNo.Text) = "0000" Then
                Call sffuncMsg("�ù���վ��δע�ᣬֻ��ϵͳ����Ա����δע��Ĺ���վ�ϵ�¼�����Ƚ��빤��վ����ע�����ʹ�ñ�ϵͳ��", sf����)
                '��ʾ������
                Unload Me
                lfrm������.Show
                plngMainHwnd = lfrm������.hWnd
            Else
                Call sffuncMsg("�ù���վ��δע�ᣬ����ע����ʹ�ñ�ϵͳ��", sf����)
                Call ccmdCancel_Click
            End If
        End If
        
        Call oesubSave("�û���¼ϵͳ", "��¼")
    Else '�Ƿ��û���������
        Call sffuncMsg("�û���Ż�������", sf����)
        ctxtUserNo.SetFocus
    End If
errHandle:
    Set lfrm������ = Nothing
    If Err.Number = 0 Then Exit Sub
    Set pobjƽ̨�ṹ = Nothing
    Call sfsub������("������", "frmLogin", "ccmdOk_click", Err.Number, Err.Description, False)
End Sub

' �޸�˵�������жϹ���վ�Ƿ�ע�ᣬʼ�շ���Trueֵ��
' �޸��ˣ�  ����
' �޸�ʱ�䣺2001-8-6
Private Function func�Ƿ�ע��(ByVal para�û���� As String) As Boolean
    func�Ƿ�ע�� = True
    Exit Function
    If um����վ��� = "" Then
        func�Ƿ�ע�� = False
    End If
End Function

'�޸ģ�2002-1-28�������ע����л�ȡ�ϴε�¼���û���š�
Private Sub Form_Load()
    Dim lstrUser As String
    
    On Error Resume Next
    lstrUser = sffuncGetSetting("ϵͳ����", "��������", "�û����")
    If lstrUser = "" Then lstrUser = "0001"
    ctxtUserNo.Text = lstrUser
    'Label1 = "�������û���źͿ����Ե�¼" & pstrSysName
    
End Sub
