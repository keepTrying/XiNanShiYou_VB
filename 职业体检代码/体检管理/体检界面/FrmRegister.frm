VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "���Ǽ�"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10845
   ClipControls    =   0   'False
   Icon            =   "FrmRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8654.759
   ScaleMode       =   0  'User
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox cchkClear 
      Caption         =   "��������"
      Height          =   345
      Left            =   9000
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   120
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   6600
      Top             =   360
   End
   Begin VB.Frame cfram������Ϣ 
      Caption         =   "�Ǽǻ�����Ϣ���ǿ���¼��ʱ��ɫΪ��¼�����¼��ʱֻ������):"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   7095
      Left            =   60
      TabIndex        =   14
      Top             =   720
      Width           =   10620
      Begin VB.CheckBox cchk¼�뵥λ���� 
         Caption         =   "¼�뵥λ����"
         Height          =   255
         Left            =   8400
         TabIndex        =   36
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox ctxt���� 
         Height          =   270
         Left            =   360
         TabIndex        =   34
         Text            =   "1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox ctxt��쵥�� 
         Height          =   315
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox ctxtTubeNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox ccmb������� 
         Height          =   300
         ItemData        =   "FrmRegister.frx":0442
         Left            =   3960
         List            =   "FrmRegister.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox ctxtAge 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3480
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox ccmbSex 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         ItemData        =   "FrmRegister.frx":045C
         Left            =   2520
         List            =   "FrmRegister.frx":0466
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   840
      End
      Begin VB.TextBox ctxtName 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2130
      End
      Begin VB.Frame frmPhoto 
         Caption         =   "����"
         ClipControls    =   0   'False
         ForeColor       =   &H00800000&
         Height          =   4275
         Left            =   5520
         TabIndex        =   27
         Top             =   2760
         Width           =   5025
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3570
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   4725
            _ExtentX        =   8017
            _ExtentY        =   6297
            BackColor       =   0
            FontSize        =   9.75
            OriginalSize    =   -1  'True
         End
      End
      Begin VB.CommandButton ccmd��λ��λ 
         Caption         =   "��λ(&T)"
         Height          =   375
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   945
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   5760
         TabIndex        =   7
         Top             =   1080
         Width           =   3240
      End
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6600
         TabIndex        =   11
         Top             =   460
         Width           =   345
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   2520
         TabIndex        =   0
         Text            =   "ccmbTemplate"
         Top             =   480
         Width           =   3480
      End
      Begin ¼��ؼ�.ctlInputDictGrid c�ֵ�� 
         Height          =   3255
         Left            =   2280
         TabIndex        =   26
         Top             =   3360
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
         Cols            =   10
         Count           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker cdtpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   129171456
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin ¼��ؼ�.ctlInputFrame ciptBase 
         Height          =   4815
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   8493
         BackColor       =   15791081
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   ""
         Rows            =   7
         Cols            =   27
         DistanceofRow   =   0
         AutoSize        =   0   'False
         FormatString    =   "���֤��,1,0,12"
         Count           =   1
         titleInputBox0001=   "���֤��"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   12
         orderInputBox0001=   1
         valueInputBox0001=   ""
         datatypeInputBox0001=   3
         colInputBox0001 =   0
         rowInputBox0001 =   1
         PassWordCharInputBox0001=   0   'False
         ����InputBox0001=   0   'False
         ����������ֵInputBox0001=   0   'False
         ���������СֵInputBox0001=   0   'False
         �ֵ�����InputBox0001=   ""
         ��ʾ�ֵ��ֶ�InputBox0001=   ""
         �����ֵ��ֶ�InputBox0001=   ""
         ����InputBox0001=   "����� 1"
         ȱʡֵInputBox0001=   ""
         ����ȱʡֵInputBox0001=   ""
         ����InputBox0001=   0
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         �����ѡInputBox0001=   0   'False
         ErrColor        =   12648447
      End
      Begin VB.Label clbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥�ţ�"
         Height          =   180
         Index           =   7
         Left            =   4080
         TabIndex        =   31
         Top             =   840
         Width           =   900
      End
      Begin VB.Label clbl��������� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   30
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ�������ڣ�"
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   29
         Top             =   1560
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ͣ�"
         Height          =   180
         Index           =   3
         Left            =   3960
         TabIndex        =   28
         Top             =   1560
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2520
         TabIndex        =   25
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   720
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������뿴״̬��"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6960
         TabIndex        =   22
         Top             =   480
         Width           =   915
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6240
         TabIndex        =   21
         Top             =   465
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   900
      End
      Begin VB.Label clblSysNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
         Height          =   180
         Index           =   1
         Left            =   6240
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ڣ�"
         Height          =   180
         Index           =   2
         Left            =   8400
         TabIndex        =   17
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���ƣ�"
         Height          =   180
         Index           =   5
         Left            =   5760
         TabIndex        =   16
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   15
         Top             =   840
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   1680
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar cstbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   7860
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16510
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��˺�
'����޸ģ��
Public pstrϵͳ��� As String
Dim iPort As Integer
Dim m������״̬ As Boolean

Private mobj����� As Object                   '�����Ա������ε���졣
Private mobj��� As Object                     '���������ṩ��ȡϵͳ��ź��Թܱ�ţ�����Ǽ���Ϣ�ķ�����
Private mobj��켯 As Object                   '��켯��������λ��Ҫ���������Ա��Ϣ��
Private mobj����ģ�� As Object               '����ģ�壬��ȡ���еķǸ�������ģ�����ơ�
Private WithEvents mobjGUI As cls����ͨ�ö���  '����ͨ�ö���������ʼ��������������¼���ؼ���
Attribute mobjGUI.VB_VarHelpID = -1

'ҵ�����á�
Private mblnTakePhoto As Boolean               'ҵ�����á��Ƿ����࡯��
Private mbln����¼�� As Boolean

Private mcolTubeNo As New Collection           '��ǰ�����ѡ���Թ���ĸ��

Private mstr��λ������ As String             '��λ��λ�������š�
Private mblnInUse As Boolean

'��ѡ��������Ŀ���շ���Ŀ
Private mcol�����Ŀ As New Collection
Private mcol�շ���Ŀ As New Collection               'item:���,key����š�

Public pstrϵͳ������� As String '�޸ģ�2002-10-10�����Ϊ�ζ��������Ӹ����ԡ�

Private mobj����  As cls�û���������
Private mstrĬ������ As String


'���ܣ����ص�ǰ�����Ƿ��Ѽ��أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchkClear_Click()
    On Error Resume Next
    ctxtName.SetFocus
End Sub

Private Sub cchk¼�뵥λ����_Click()
    Dim lblnVisible As Boolean
    On Error Resume Next
    If cchk¼�뵥λ����.Value = 1 Then
        lblnVisible = True
    Else
        lblnVisible = False
    End If
    ccmbUnit.Visible = lblnVisible
    ccmd��λ��λ.Visible = lblnVisible
    Label2(5).Visible = lblnVisible
    ctxtName.SetFocus
End Sub

Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub
Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtAge.SetFocus
    End If
End Sub

Private Sub ccmbTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxtTubeNo.Visible Then
            ctxtTubeNo.SetFocus
        Else
            ctxt��쵥��.SetFocus
        End If
    End If
End Sub

'���ܣ����Ʋ��������������ƣ�ֻ��ѡ��
'������2002-11-28�������
Private Sub ccmbTemplate_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ccmbUnit_Click()
    On Error GoTo errHandler
    Dim i As Integer
    
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '���뵽�б����
        ccmbUnit.AddItem ccmbUnit.Text
        
        '���ص��������䲾�ļ���
        pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
    Else
        '�޸ģ�2001-8-23��
        On Error Resume Next
        mstr��λ������ = pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text)
        sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
    End If
    Exit Sub
errHandler:

End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxt����.Visible Then
            ctxt����.SetFocus
        Else
            If ciptBase.Visible Then
                ciptBase.SetFocus
            End If
        End If
    Else
        mstr��λ������ = ""
    End If
        
End Sub

Private Sub ccmb�������_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    If KeyCode = 13 Then
        ctxt����.SetFocus
    End If
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub


Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '�Զ����档
    If ctbMain.Buttons(6).Enabled Then
        ctxtName.SetFocus
        SendKeys "{F2}"
    End If
End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c�ֵ��" Then
        c�ֵ��.Visible = False
    End If

End Sub


Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt��쵥��.SetFocus
    End If

End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub

Private Sub ctxtTubeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        cdtpDate.SetFocus

    End If
End Sub



Private Sub ctxt����_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '��¼���û��¼����Ŀ����ֱ�ӱ��档
        If ciptBase.Visible Then
            ciptBase.SetFocus
            ciptBase.ItemSetFocus 0
        End If
    End If
End Sub

Private Sub ctxt��쵥��_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ccmbUnit.Visible Then
            ccmbUnit.SetFocus
        Else
            If ctxt����.Visible Then
                ctxt����.SetFocus
            Else
                If ciptBase.Visible Then
                    ciptBase.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnTakePhoto Then
        '���³�ʼ������ؼ���
        cctlCatchPhoto.funcInitVideo
    End If
    ctxtName.SetFocus
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    gfsubHideComboList ccmbUnit
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
   
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    MousePointer = 11
    
    '���治�ɲ�����
'    cfram������Ϣ.Enabled = False
    ctbMain.Enabled = False
    
    Set mcol�շ���Ŀ = New Collection
    Set mcol�����Ŀ = New Collection
    
    Set mobj����� = CreateObject("������.clsMedicalExam")
    
    Set mobj��� = CreateObject("������.clsMedicalExam")
    '�޸ģ�2002-10-10������ϵͳ������ƣ���
    If pstrϵͳ������� <> "" Then
        mobj���.ϵͳ������� = pstrϵͳ�������
    End If
    
    Set mobj��켯 = CreateObject("������.clsMedicalExamSet")
    Set mobj����ģ�� = CreateObject("������.ClsMedicalExamTemplate")
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    mobjGUI.pbln�Զ������ֵ�߶� = False
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    Dim lcol��������ť As New Collection           '�������ϵİ�ť��ʼ�����ϡ�
    With lcol��������ť
        .Add "���"
        .Add "|"
        .Add "�����Ŀ(&T)102"
        .Add "������Ƭ(&E)103"
        .Add "|"
        .Add "����"
        .Add "�޸�"
        .Add "|"
        .Add "����(&F5)106"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
        Set .c¼��� = ciptBase
        Set .c�ֵ�� = c�ֵ��
        Set .c״̬�� = cstbMain
        
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcol��������ť, ""
    End With
    pCardIdInit
    
    If m������״̬ = True Then
       ctbMain.Buttons(9).Enabled = True
    Else
       ctbMain.Buttons(9).Enabled = False
    End If
      
    '���
    subClear
    cdtpDate.Value = Date
    '�·���ϵͳ���
    clblSysNo.Caption = mobj���.Func����ϵͳ���
    mobj���.ϵͳ��� = clblSysNo.Caption
    
    cctlCatchPhoto.Visible = False
    cctlCatchPhoto.Visible = True
    
    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
        ctxtTubeNo.Visible = True
        ctxtTubeNo.TabIndex = 1
        clblTubeNo.Visible = False
        clblLetter.Visible = False
        cvscLetter.Visible = False
    Else
        ctxtTubeNo.Visible = False
        clblTubeNo.Visible = True
        clblLetter.Visible = True
        cvscLetter.Visible = True
    End If
    
    DoEvents
    
    'Ϊ�˼ӿ촰������ٶȣ����³�ʼ���������ڶ�ʱ������ɡ�
    Timer1.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "Form_Load", 6666, lstrError, False
    '�ָ����������á�
    ctbMain.Enabled = True
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
End Sub

'���ܣ����form_load���µĳ�ʼ��������
Private Sub Timer1_Timer()
    Dim lobj����ģ�弯 As Object  '����ģ�弯����ȡ���еķǸ�������ģ�����ơ�
    Dim lcolInfo As Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '��ʱ�����������á�
    Timer1.Enabled = False
    
    '�ӵ��չ������Ѳ��л�ȡ����¼����ĵ�λ���ơ�
    Set lcolInfo = pobjҵ�����.���չ������䲾.��λ���Ƽ�
    For i = 1 To lcolInfo.Count
        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    '�����еķǸ�������ģ����뵽���������б���С�
    Set lobj����ģ�弯 = CreateObject("������.ClsMedicalExamTemplateSet")
    lobj����ģ�弯.�������� = 3
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
    Next
    Set lobj����ģ�弯 = Nothing
    
    '����ҵ�������ж��Ƿ����ࡣ
    If pobjҵ�����.ҵ������("�Ƿ�����") = "��" Then
        mblnTakePhoto = True
    Else
        mblnTakePhoto = False
    End If
    
    If pobjҵ�����.ҵ������("�Ƿ���ٵǼ�") = "��" Then
        mbln����¼�� = True
    Else
        mbln����¼�� = False
    End If
    
    'ֻ�г��죬���ҿ��ٵǼǲſ��������Ǽǡ�
    If Not mbln����¼�� Or pstrϵͳ��� <> "" Then
        clbl����.Visible = False
        ctxt����.Visible = False
    End If
    
    ccmb�������.ListIndex = 0
    
    If ccmbTemplate.ListCount > 0 Then
        'ccmbTemplate.ListIndex = 0
        ccmbTemplate.Text = ccmbTemplate.List(0)
        subChangeTemplate
        
    End If
    
    '��Ҫ����ʱ��ʼ������ؼ���
    If mblnTakePhoto Then
        '��ʼ���ؼ���
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pstrϵͳ��� <> "" Then
        '�����Ǽǡ�
        '��ʾ�����Ա������Ϣ��
        SubGetPersonInfo pstrϵͳ���
    End If
    
    On Error Resume Next
    Set mobj���� = New cls�û���������
    mobj����.�û���� = "*"
    mobj����.ҵ���� = "������"
    mstrĬ������ = mobj����.������ֵ("�������")
'    If mstrĬ������ <> "" And ctxtAge = "" Then
'        ctxtAge = mstrĬ������
'    End If
    
    If mobj����.������ֵ("���Ǽ�ʱ¼�뵥λ����") = "" Or mobj����.������ֵ("���Ǽ�ʱ¼�뵥λ����") = "��" Then
        cchk¼�뵥λ����.Value = 1
    Else
        cchk¼�뵥λ����.Value = 0
    End If
    cfram������Ϣ.Enabled = True
    ctbMain.Enabled = True
    MousePointer = 0
    Exit Sub
    
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "Timer1_Timer", 6666, lstrError, False
    
    '�ָ�����ɲ�����
    cfram������Ϣ.Enabled = True
    ctbMain.Enabled = True
    MousePointer = 0

End Sub


Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    'ѡ������
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    MousePointer = 11
    
    subChangeTemplate
    
    ctbMain.Buttons(6).Enabled = True
'    If ctxtTubeNo.Visible Then
'        ctxtTubeNo.SetFocus
'    Else
        ctxtName.SetFocus
'    End If
    MousePointer = 0
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "ccmbTemplate_Click", 6666, lstrError, False
    
    Exit Sub
    Resume
End Sub

Private Sub subChangeTemplate()
    On Error GoTo errHandler
    'ѡ������
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    '��ȡ���Թܱ�š�
    If mobj���.����.������ <> ccmbTemplate.Text Then
        mobj���.����.������ = ccmbTemplate.Text

        '��������ģ���ȡ���������п��õ���ĸ��
        mobj����ģ��.������ = ccmbTemplate.Text

        If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
            '�Թܱ����ĸΪ��ʱcvscLetter����
            If mobj���.����.�Թܱ����ĸ = "" Then
                '����ĸ�����ŷֿ�������mcoltubeNo��
                lstrTubeNo = mobj����ģ��.�Թ���ĸ���
                If Right(lstrTubeNo, 1) <> "," Then lstrTubeNo = lstrTubeNo & ","
                lstrTemp = ""
                Set mcolTubeNo = New Collection
                For i = 1 To Len(lstrTubeNo)
                    lstrTemp = lstrTemp & Mid(lstrTubeNo, i, 1)
                    If Mid(lstrTubeNo, i, 1) = "," Then
                        If Left(lstrTemp, Len(lstrTemp) - 1) <> "" Then
                            mcolTubeNo.Add Left(lstrTemp, Len(lstrTemp) - 1)
                        End If
                        lstrTemp = ""
                    End If
                Next i
                If mcolTubeNo.Count > 0 Then
                    '�Թ���ĸ�ı��ˣ�������ʾ��
                    If clblLetter.Caption <> "" And clblLetter.Caption <> mcolTubeNo(1) Then
                        sffuncMsg "��ע�⣬������ѡ�������ʹ�õ��Թ���ĸ��ǰһ����" & clblLetter.Caption & "����ͬ�ˡ�"
                    End If

                    '��ֵ��clblLetter
                    clblLetter.Caption = mcolTubeNo(1)
                    cvscLetter.Enabled = True
                    cvscLetter.Min = 1
                    cvscLetter.Max = mcolTubeNo.Count
                    cvscLetter.Value = 1
                Else
                    ctbMain.Buttons(6).Enabled = False
                    '��ʾ�������޿��õ���ĸ��
                    Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
                End If
            Else
                '����ĸ������ѡ����ĸ��
                clblLetter.Caption = mobj���.����.�Թܱ����ĸ
                cvscLetter.Enabled = False
            End If
        Else
            clblLetter.Caption = mobj����ģ��.�Թ���ĸ���
        End If
        
        '��ʼ��������Ϣ��
        On Error Resume Next
        mobjGUI.sub��ʼ��¼��� ccmbTemplate.Text
        
        '�޸ģ�2001-8-23����ʾ��λ���ԣ���
        If mstr��λ������ <> "" Then
            sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
        End If

        '������д������Ϣֵ��
        If mobj����ģ��.����������Ŀ��.Count > 0 Then
            Set lcolInfo = mobj�����.����.������Ϣ
            If lcolInfo.Count > 0 Then
                sub��¼���ֵ ciptBase, mobjGUI, lcolInfo
            End If
        End If

        '�޸ģ�2002-7-26��������ݡ��Ƿ�����ѡ���������͡�
        If mobj����ģ��.�Ƿ����� Then
            ccmb�������.ListIndex = 1
        Else
            ccmb�������.ListIndex = 0
        End If

        '�޸ģ�2002-10-10������ζ����ƣ���ʾ����
        On Error Resume Next
        ciptBase.Box1("�����").Text = mobj����ģ��.�շѱ�׼���
'

    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "subChangeTemplate", 6666, lstrError, True
    
    Exit Sub
    Resume
End Sub

'�Զ������б��
Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
'    gfsubShowComboList ccmbUnit
    Exit Sub
errHandler:
    'sfsub������ "�����沿��", "FrmRegisterAnnual", "ccmbUnit_GotFocus", Err.Number, Err.Description, False
End Sub

Private Sub ccmbUnit_LostFocus()
    On Error GoTo errHandler
    Dim i As Integer
    If Trim(ccmbUnit.Text) = "" Then Exit Sub
    
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '���뵽�б����
        ccmbUnit.AddItem ccmbUnit.Text
        
        '���ص��������䲾�ļ���
        pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
    Else
        '�޸ģ�2001-8-26������λ�����Ų�ͬ���޸Ĺ������䲾����
        If mstr��λ������ <> pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text) And mstr��λ������ <> "" Then
            pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
        End If
    End If
    Exit Sub
errHandler:
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
    Case vbKeyF8
        If mblnTakePhoto Then
            If cctlCatchPhoto.VideoIsOk Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
    
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

'���ܣ������ʼ����

'���õ�λ��λ
Private Sub ccmd��λ��λ_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��

    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
            mstr��λ������ = lobjRec!������
            
            If mstr��λ������ <> "" Then
                '�޸ģ�2001-8-23����ʾ��λ���ԣ���
                On Error Resume Next
                sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
            End If
        End If
    End If
    
    '�ѽ���ص���λ¼��򡣱����ܱ����µ�λ��λ��Ϣ��
    ccmbUnit.SetFocus
    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub


Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal ���� As String, ByVal ���� As String, ByVal �������� As String, ByVal IsError As Boolean)
    On Error GoTo errHandler
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    

    ldatBirth = ""
    Select Case ����
    Case "���֤��"
        lstrIDCard = ciptBase.ItemText(Index)
        If lstrIDCard <> "" Then
            '��ȷʱ�����֤���л�ȡ�������ڡ�
            sub���ݹ�����ݺ����ȡ���պ��Ա� lstrIDCard, ldatBirth, lstrSex
            If Not IsDate(ldatBirth) Then
                Err.Raise 6666, , "���֤�Ų��Ϸ���"
            End If
            
            '�����Ƿ���Ҫ¼��������ڣ���Ҫʱ�Զ��������֤����д��������
            On Error Resume Next
            If IsDate(ldatBirth) Then
                ciptBase.Box1("��������").Text = ldatBirth
                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
            End If
        End If
    Case "��������"
        Dim lstrItemText  As String
        '������ҵ���¼�����ֵ䡣
        For i = 1 To ciptBase.InfoCollection.Count
            If ciptBase.InfoCollection(i).Title = "��ҵ���" Then
                If Not ciptBase.InfoCollection(Index + 1).DictRecordSet Is Nothing Then
                    If ciptBase.InfoCollection(Index + 1).DictRecordSet.EOF Then
                    Else
                        mobjGUI.sub��ʼ���ֵ�� i, "Parent=" & ciptBase.InfoCollection(Index + 1).DictRecordSet("InnerId")
                    End If
                End If
                ciptBase.pblnTemp = True
                lstrItemText = ciptBase.Box1(i - 1).Text
                ciptBase.Box1(i - 1).Text = ""
                ciptBase.Box1(i - 1).Text = lstrItemText
                ciptBase.pblnTemp = False
                
                Exit For
            End If
        Next
    Case "����"
        '��Ч���жϡ�
        If ���� <> "" Then
            If Val(����) > 100 Then
                Err.Raise 6666, , "���䲻�ܴ���100��"
            End If
            If Val(����) >= Val(ctxtAge.Text) Then
                Err.Raise 6666, , "����>=���䣬���ǷǷ������ݣ�"
            End If
        End If
        
    End Select
    Exit Sub

errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemBox(Index).Text = ""
    ciptBase.ItemSetFocus Index
End Sub





Private Sub cvscLetter_Change()
    On Error Resume Next
    '����������������Ӧ����ĸ��
    If mcolTubeNo.Count > 0 Then
        clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
    End If
End Sub

'���ܣ���ս��档
Private Sub subClear()
    
    On Error Resume Next
    'cdtpDate.Value = Date
    ctxtName.Text = ""
    'ctxtAge = mstrĬ������
    ctxtAge = ""

    ccmbUnit.Text = ""
    
    ctxtTubeNo = ""
    ctxt��쵥�� = ""
    
    '�޸ģ�2002-10-10������ζ����ƣ�������ա�
    Dim ldbl����� As Double
    ldbl����� = ciptBase.Box1("�����").Text
    ciptBase.ClearContent
    ciptBase.Box1("�����").Text = ldbl�����
    
    clbl���������.Caption = ""
    Label2(4).Visible = False
    clbl���������.Visible = False
    Set cctlCatchPhoto.Photo = Nothing
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '����������¼û�б��棬�˻�ϵͳ��š�
    If Not mobj��� Is Nothing Then
        If mobj���.ϵͳ��� <> "" And Not mobj���.�Ƿ��Ѵ��� Then
            '�˻�ϵͳ��š�
            mobj���.sub�˻�ϵͳ��� mobj���.ϵͳ���
        End If
    End If
    mobj����.sub���Ǽ���ֵ "���Ǽ�ʱ¼�뵥λ����", IIf(cchk¼�뵥λ����.Value = 1, "��", "��")
     
    Set mobj��� = Nothing
    Set mobj��켯 = Nothing
    Set mobj����ģ�� = Nothing
    '�ر������
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    mblnInUse = False
    pstrϵͳ������� = ""
End Sub


'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    Dim lstr��ˮ�� As String
    Dim lstrϵͳ��� As String
    Dim lcolԭ�����Ŀ As Collection
    
    On Error GoTo errHandler
    
    Select Case Operate
    
    Case "���"
        subClear
        '��ս���������ţ���ʾ��¼��������Ա��
        mobj���.�����Ա.����������� = ""
        
        Cancel = True
    
    Case "����"
        '�ж��Ƿ���Ҫ���ࡣ
        If mblnTakePhoto = True Then
            '�ж��Ƿ�����
            If cctlCatchPhoto.Photo Is Nothing Then
                Err.Raise 6666, , "���ڡ�ҵ�����á���������Ҫ���񣬵���������û�����࣬�޷����档����취��" & Chr(13) & Chr(10) & "��1�� �밴��ȡ�񡱰�ť����󱣴棡" & Chr(13) & Chr(10) & "��2�����㲻׼�����࣬���Ƚ��롰ҵ�����á����ò����ࡣ"
            End If
        End If
        
        '�����ǿ���¼�룬���¼���Ƿ��д���
        If mobj���.����.������Ϣ.Count > 0 Then
            '�޸ģ�2001-9-12�������
            On Error Resume Next
            ciptBase.Box1(ciptBase.ActiveInputBoxIndex).LostFocus
            On Error GoTo errHandler
            
            If ciptBase.ItemsError.Count > 0 And Not mbln����¼�� Then
                Err.Raise 6666, , "�������ɫ¼������ݣ�"
            End If
        End If
        '��ֹ˫����������¿հ׼�¼
        ctbMain.Buttons(6).Enabled = False
        MousePointer = 11
        
        '�����Թܱ�Ų�����
        With mobj���
            If .����.������ <> ccmbTemplate.Text Then
                .����.������ = ccmbTemplate.Text
            End If
            '�޸ģ�2004-1-9���Թܱ�ſ������룩
            If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
                If .����.�Թܱ����ĸ <> clblLetter.Caption Then
                    .����.�Թܱ����ĸ = clblLetter.Caption
                End If
            Else
                .����.�Թܱ����ĸ = clblLetter.Caption
                .�Թܱ�� = ctxtTubeNo.Text
            End If
            
            .�����Ա.���� = ctxtName
            .�����Ա.�Ա� = ccmbSex.Text
            .�����Ա.��λ���� = ccmbUnit.Text
            
            If mblnTakePhoto Then
                .�����Ա.��Ƭ = cctlCatchPhoto.Photo
'                .�����Ա.��Ƭѹ�� = cctlCatchPhoto.Photo
            End If
            If Val(ctxtAge.Text) > 0 Then
'                If Val(ctxtAge.Text) > 200 Then
'                    Err.Raise 6666, , "���䳬��ϵͳ������������200��"
'                End If
                .�����Ա.�������� = DateAdd("yyyy", -Val(ctxtAge.Text), Date)
            Else
                '��������ַ������������䡣
                mobj����.sub���Ǽ���ֵ "�������", ctxtAge.Text
                mstrĬ������ = ctxtAge.Text
            End If
            .�����Ա.���� = ctxtAge.Text
            
            On Error Resume Next
            .�����Ա.������ݺ��� = ciptBase.Box1("���֤��").Text
            .�����Ա.�������� = ciptBase.Box1("��������").TrueText
            .�����Ա.��ҵ��� = ciptBase.Box1("��ҵ���").TrueText
            .�����Ա.Ƭ�� = ciptBase.Box1("Ƭ��").TrueText
            
            If ccmbUnit.Text = "" Then
                .�����Ա.��λ������ = ""
            Else
                If .�����Ա.��λ������ <> mstr��λ������ Then
                    '����λ������¸�ֵ���������»�ȡ���������ࡢ��ҵ���Ƭ����
                    .�����Ա.��λ������ = mstr��λ������
                End If
            End If
            
            '���渽����Ϣ
            For i = 1 To ciptBase.ItemCount
                'If ciptBase.Box1(i - 1).TrueText <> ciptBase.Box1(i - 1).Text And ciptBase.Box1(i - 1).Text <> "" Then
                If ciptBase.InfoCollection(i).�ֵ����� <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
                    .����.Sub�����Ϣֵ ciptBase.InfoCollection(i).����, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
                Else
                    .����.Sub�����Ϣֵ ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
                End If
            Next i
            
            '����Ϊ������
            If ccmb�������.Text = "����" Then
                .������ = P_EXAM_FIRST
            Else
                .������ = P_EXAM_ANNUAL
            End If
            .������� = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            '�޸ģ�2004-1-9��������쵥�ţ�
            .��쵥�� = ctxt��쵥��.Text

        End With
        
        On Error GoTo errHandler
        '�޸�����
        If mcol�����Ŀ.Count > 0 Then
            '��ȡ���������е������Ŀ��
            Set lcolԭ�����Ŀ = mobj���.����.�����Ŀ��("")
            'ɾ��ȥ���ġ�
            For i = 1 To lcolԭ�����Ŀ.Count
                If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(mcol�����Ŀ, lcolԭ�����Ŀ(i).�����Ŀ���) Then
                    mobj���.����.Subɾ�������Ŀ lcolԭ�����Ŀ(i).�����Ŀ���
                End If
            Next
            '���������Ŀ
            For i = 1 To mcol�����Ŀ.Count
                mobj���.����.Sub��������Ŀ mcol�����Ŀ(i)("����")
            Next
            
        End If
        
        If mcol�շ���Ŀ.Count > 0 Then
            pobjҵ�����.Sub���Ǽ� mobj���, , , mcol�շ���Ŀ, Val(ctxt����)
        Else
            pobjҵ�����.Sub���Ǽ� mobj���, , , , Val(ctxt����)
        End If
        cstbMain.Panels(1) = "�ϴα�������ϵͳ��ţ�" & mobj���.ϵͳ��� & "���Թܱ�ţ�" & mobj���.�Թܱ��
        If mobj���.�շ����� <> "" Then
            cstbMain.Panels(1) = cstbMain.Panels(1) & "���շ����ţ�" & mobj���.�շ�����
        End If
        
        
        If cchkClear = 1 Then
            subClear
        End If
        
        '�����µ�ϵͳ��š�
        clblSysNo.Caption = mobj���.Func����ϵͳ���
        mobj���.ϵͳ��� = clblSysNo.Caption
        mobj���.����.������ = ccmbTemplate.Text
        
        Set mcol�����Ŀ = New Collection
        Set mcol�շ���Ŀ = New Collection
        '�ָ����ࡣ
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "�ָ�" Then
                cctlCatchPhoto.subת��״̬
            End If
            
        End If
        
        '�Թ���ĸ������ѡ��
        cvscLetter.Enabled = False
        ctxtName.SetFocus

        frmRegisterManage.sub��ѯ����ʾ
        
        ctbMain.Buttons(6).Enabled = True
        Cancel = True
        MousePointer = 0
    
    Case "�����Ŀ"
        Dim lobj���ģ�� As Object
        
        '��ȡ���������е������Ŀ��
        Set lcolԭ�����Ŀ = mobj���.����.�����Ŀ��("")
        
        '����ѡ����Ŀ��������ԡ�
        frmSelectItem.pstr�������� = ccmbTemplate.Text
        Set frmSelectItem.pcol������Ŀ = lcolԭ�����Ŀ
        Set frmSelectItem.pcol�շ���Ŀ = mcol�շ���Ŀ
        '����ѡ����Ŀ���档
        frmSelectItem.Show 1
        If frmSelectItem.pblnOk Then
            '��ȡѡ�еĸ�����Ŀ��
            Set mcol�����Ŀ = frmSelectItem.pcol������Ŀ
            '��ȡ���õ��շ���Ŀ��
            Set mcol�շ���Ŀ = frmSelectItem.pcol�շ���Ŀ
            
            '�޸ģ�2002-10-14������ζ����ƣ���ʾ�շѽ�
            Dim ldblTotal As Double
            For i = 1 To mcol�շ���Ŀ.Count
                ldblTotal = Format(ldblTotal + mcol�շ���Ŀ(i)("����"), "0.00")
            Next
            On Error Resume Next
            If sffunc�жϼ��ϼ�ֵ�Ƿ����(mobj���.����.������Ϣ, "�����") Then
                ciptBase.Box1("�����").Text = ldblTotal
                mobj���.����.Sub�����Ϣֵ "�����", ldblTotal
            End If
            
        End If
    Case "������Ƭ"
        Dim lstrFile As String
        ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ", vbDirectory) <> "" Then
            ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ"
        End If
        ccmdFile.FileName = clblSysNo.Caption
        ccmdFile.ShowOpen
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            If InStr(lstrFile, ".") > 0 Then
                Set cctlCatchPhoto.Photo = LoadPicture(lstrFile)
                mblnTakePhoto = True
            End If
        End If
    Case "�޸�"
        Dim lobjRec As Object
        '��ȡ�������ĺš�
        If Val(Right(clblSysNo.Caption, Len(clblSysNo) - Len(mobj���.ϵͳ��Ź̶�����))) > 1 Then
            FrmEditRegister.ϵͳ��� = mobj���.ϵͳ��Ź̶����� & Format(Val(Right(clblSysNo.Caption, Len(clblSysNo) - Len(mobj���.ϵͳ��Ź̶�����))) - 1, String(Len(clblSysNo) - Len(mobj���.ϵͳ��Ź̶�����), "0"))
        Else
            FrmEditRegister.ϵͳ��� = ""
        End If
        FrmEditRegister.Show 1, Me
    Case "����"
        sub������Ϣ
    End Select
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    ctbMain.Buttons(6).Enabled = True
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
    Cancel = True
    Exit Sub
    Resume
    Exit Sub

End Sub


'���ܣ���ʾָ��ϵͳ��ŵ������Ա����Ϣ�ڽ����ϡ�
Private Sub SubGetPersonInfo(ByVal paraϵͳ��� As String)
    Dim lcolInfo As New Collection
    Dim i As Integer
    Dim j As Integer
    Dim lstrTemp As String
    Dim lstrTubeNo As String
    Dim lstrSysNo As String
    
    
    On Error GoTo errHandler
    MousePointer = 11
    
    '������ʱ���ɲ�����
    ctbMain.Enabled = False
    
    '���˻ؾ�ϵͳ��š�
    If Not mobj���.�Ƿ��Ѵ��� And mobj���.ϵͳ��� <> "" Then
        mobj���.sub�˻�ϵͳ��� mobj���.ϵͳ���
    End If
    
    '������������
    Set mobj����� = CreateObject("������.clsMedicalExam")
    mobj�����.ϵͳ��� = paraϵͳ���
    
    '����ϴ���������
    If ccmbTemplate.Text <> mobj�����.����.������ Then
        ccmbTemplate.Text = mobj�����.����.������
    
        '���³�ʼ��¼��塣
        On Error Resume Next
        mobjGUI.sub��ʼ��¼��� mobj�����.����.������
        On Error GoTo errHandler
    End If
    
    '��ȡ������¼�ĸ�����Ϣ��
    Set lcolInfo = mobj�����.����.������Ϣ
    
    '��д������Ϣֵ
    sub��¼���ֵ ciptBase, mobjGUI, lcolInfo
    
    '��ʾ������Ϣ��
    With mobj�����.�����Ա
        ctxtName.Text = .����
        ccmbSex.Text = .�Ա�
        ctxtAge.Text = .����
        ccmbUnit.Text = .��λ����
        ccmbUnit_LostFocus
        
        '��Ƭ
        '��ò���ʾ��Ƭ��
        If Not .��Ƭ Is Nothing Then
            Set cctlCatchPhoto.Photo = .��Ƭ
        Else
            cctlCatchPhoto.subClear
        End If
        
        '�޸ģ�2001-8-23��
        On Error Resume Next
        mstr��λ������ = .��λ������
        
        On Error GoTo errHandler
    End With
    
    '�޸ģ�2001-12-30����ʾ�ϴ�������ڣ���
    Label2(4).Visible = True
    clbl���������.Visible = True
    clbl���������.Caption = mobj�����.�������
    
    '�޸ģ�2002-1-6����ʱ��������18���£��Զ�����Ϊ���죩��
    If IsDate(clbl���������.Caption) Then
        If DateDiff("m", clbl���������.Caption, Now) >= 18 Then
            ccmb�������.ListIndex = 0
        Else
            '����18���£��Զ�����Ϊ��졣
            ccmb�������.ListIndex = 1
        End If
    End If
    '�����µ�ϵͳ���
    lstrSysNo = mobj���.Func����ϵͳ���
    mobj���.ϵͳ��� = lstrSysNo
    clblSysNo.Caption = lstrSysNo
    
    '�����������䡣
    mobj���.�����Ա.����������� = mobj�����.�����Ա.�����������
    
    
    '�����������������Ӷ���ȡ���Թܱ�š�
    mobj���.����.������ = ccmbTemplate.Text
    
    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
        '��ȡ������ĵ�����ʹ�õ��Թܱ����ĸ��
        clblLetter.Caption = mobj���.����.�Թܱ����ĸ
        If clblLetter.Caption = "" Then
            
            '�ô����Ǽ��ǵ���ĵ�һ����������ģ������л�ȡ���п�ѡ����Ļ��
            mobj����ģ��.������ = ccmbTemplate.Text
            lstrTubeNo = mobj����ģ��.�Թ���ĸ���
            
            '����ĸ�����ŷֿ�������mcoltubeNo�С�
            If Right(lstrTubeNo, 1) <> "," Then lstrTubeNo = lstrTubeNo & ","
            lstrTemp = ""
            Set mcolTubeNo = New Collection
            For i = 1 To Len(lstrTubeNo)
                lstrTemp = lstrTemp & Mid(lstrTubeNo, i, 1)
                If Mid(lstrTubeNo, i, 1) = "," Then
                    If Left(lstrTemp, Len(lstrTemp) - 1) <> "" Then
                        mcolTubeNo.Add Left(lstrTemp, Len(lstrTemp) - 1)
                    End If
                    lstrTemp = ""
                End If
            Next i
            If mcolTubeNo.Count > 0 Then
                '�Թ���ĸ�ı��ˣ�������ʾ��
                If clblLetter.Caption <> "" And clblLetter.Caption <> mcolTubeNo(1) Then
                    sffuncMsg "��ע�⣬������ѡ�������ʹ�õ��Թ���ĸ��ǰһ����" & clblLetter.Caption & "����ͬ�ˡ�"
                End If
            
                '��ֵ��clblLetter��
                clblLetter.Caption = mcolTubeNo(1)
                '��ĸ����ѡ��
                cvscLetter.Enabled = True
                cvscLetter.Min = 1
                cvscLetter.Max = mcolTubeNo.Count
                cvscLetter.Value = 1
            Else
                ctbMain.Buttons(6).Enabled = False
                '��ʾ�������޿��õ���ĸ��
                Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
            End If
        Else
            '����ĸ������ѡ����ĸ��
            cvscLetter.Enabled = False
        End If
    Else
        ctxtTubeNo = mobj���.�Թܱ��
    End If
    '���水ť���á�
    ctbMain.Buttons(6).Enabled = True
    Err.Clear
    
errHandler:
    '�ָ�����ɲ�����
    ctbMain.Enabled = True
    MousePointer = 0
    If Err <> 0 Then
        sfsub������ "�����沿��", "FrmRegisterAnnual", "SubGetPersonInfo", Err.Number, Err.Description, True
    End If
    
    Exit Sub
    Resume
End Sub
    
    
'��ʼ������֤�����豸 �켽�� 2012-08-28
Private Sub sub������Ϣ()
    
On Error GoTo errHandler
    Dim nLen As Long
    Dim nRet As Long
    Dim utmp(1264) As Byte
    Dim uName(32) As Byte
    Dim uSex(4) As Byte
    Dim uNation(6) As Byte
    Dim uBorn(18) As Byte
    Dim uAddress(72) As Byte
    Dim uIDCardNo(38) As Byte
    Dim uGrantDept(32) As Byte
    Dim uUserLife(18) As Byte
    Dim uUserLifeEnd(18) As Byte
    Dim uPhotoFileName(255) As Byte
    Dim i, j, k, l, m, n, o, p, q, r As Integer
    Dim strMsg, tmpstr, tmpstr2 As String
    Dim pucIIN(4) As Byte
    Dim pucSN(8) As Byte

    nRet = Syn_OpenPort(iPort)
    If nRet = 0 Then
        nRet = Syn_GetSAMStatus(iPort, 0)
        nRet = Syn_StartFindIDCard(iPort, pucIIN(0), 0)
        nRet = Syn_SelectIDCard(iPort, pucSN(0), 0)
        nLen = Syn_ReadMsg(iPort, 0, utmp(0))
        If nLen = 0 Then
            For i = 0 To 32
                uName(i) = utmp(i)
            Next i
            For j = 0 To 4
                uSex(j) = utmp(32 + j)
            Next j
            For k = 0 To 6
                uNation(k) = utmp(36 + k)
            Next k
            For l = 0 To 18
                uBorn(l) = utmp(42 + l)
            Next l
            For m = 0 To 72
                uAddress(m) = utmp(60 + m)
            Next m
            For n = 0 To 38
                uIDCardNo(n) = utmp(132 + n)
            Next n
            For o = 0 To 32
                uGrantDept(o) = utmp(170 + o)
            Next o
            For p = 0 To 18
                uUserLife(p) = utmp(202 + p)
            Next p
            For r = 0 To 18
                uUserLifeEnd(r) = utmp(220 + r)
            Next r
           For q = 0 To 255
                uPhotoFileName(q) = utmp(276 + q)
            Next q
            strMsg = "������" + StrConv(uName, vbUnicode)
            
            '��ȡ����
            ctxtName.Text = StrConv(uName, vbUnicode)
            ctxtName.Text = Trim(ctxtName.Text)
            strMsg = "�Ա�" + StrConv(uSex, vbUnicode)
            
            
            
            '��ȡ�Ա�
            If StrConv(uSex, vbUnicode) = 1 Then
                ccmbSex.ListIndex = 0
            Else
                ccmbSex.ListIndex = 1
            End If
            
            
            '��ȡ����
            ctxtAge.Text = Year(Now()) - Left(StrConv(uBorn, vbUnicode), 4)
            On Error Resume Next
             '��ȡ���֤��
            ciptBase.Box1("���֤��").Text = StrConv(uIDCardNo, vbUnicode)
         
            '��ȡͼƬ
            Set cctlCatchPhoto.Photo = LoadPicture(StrConv(uPhotoFileName, vbUnicode))
         
         
            
           
            
           
        Else
            sffuncMsg "������֤ʧ�ܣ��п�����û�з����֤��"
        End If
    Else
       sffuncMsg "�򿪶�����ʧ��"
        m������״̬ = False
    End If
    Syn_ClosePort (iPort)
    
    
    
    
Exit Sub
errHandler:

    If Err <> 0 Then
        sfsub������ "�����沿��", "FrmRegister", "sub������Ϣ", Err.Number, Err.Description, True
    End If
    
    Exit Sub
End Sub

Private Sub pCardIdInit()
    
On Error GoTo errHandler

    Dim i, j As Long                '����ѭ������
    Dim sMsg As String              '������Ϣ�����ִ�
    Dim ret As Long                 '����״̬��������
    
    '��ʼ����ʾ����
     cstbMain.Panels(2).Text = ""
     m������״̬ = False
     
     cstbMain.Panels(2).Text = "���֤����������ʼ��ʧ�ܣ�"
    
    
    '�ֱ��1001-1017���˿ڣ���Ѱ�豸�˿�
    For j = 1001 To 1017
        If Syn_OpenPort(j) = 0 Then
            If Syn_GetSAMStatus(j, 0) = 0 Then
                iPort = j
                Syn_ClosePort (j)
                sMsg = "������������" + Str(j) + "USB�˿���"
                
                 cstbMain.Panels(2).Text = "���֤����������ʼ���ɹ���"
                
                 m������״̬ = True
                GoTo FIND
            End If
            Syn_ClosePort (j)
        End If
    Next j
    
    '��ʼ�����򿪣��˿��豸�ϵĶ����豸
    For i = 1 To 17
        If Syn_OpenPort(i) = 0 Then
            If Syn_GetSAMStatus(i, 0) = 0 Then
                iPort = i
                Syn_ClosePort (i)
                sMsg = "�����������ڴ���" + Str(i) + "��"
               
                cstbMain.Panels(2).Text = "���֤����������ʼ���ɹ���"
              
                 m������״̬ = True
                GoTo FIND
            End If
            Syn_ClosePort (i)
        End If
    Next i
    
FIND:
Exit Sub
errHandler:

    If Err <> 0 Then
        sfsub������ "�����沿��", "FrmRegister", "pCardIdInit", Err.Number, Err.Description, True
    End If
    
    Exit Sub
End Sub
