VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "¼��ؼ�.ocx"
Begin VB.Form FrmRegisterFirst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ǽ�"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9363.768
   ScaleMode       =   0  'User
   ScaleWidth      =   10560
   StartUpPosition =   3  '����ȱʡ
   Begin ¼��ؼ�.ctlInputDictGrid c�ֵ�� 
      Height          =   4455
      Left            =   5825
      TabIndex        =   25
      Top             =   2552
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7858
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   2160
      Top             =   600
   End
   Begin ¼��ؼ�.ctlInputFrame ciptBase 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   7858
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      BorderStyle     =   0
      Caption         =   "Frame1"
      Rows            =   6
      Cols            =   6
      DistanceofRow   =   0
      BorderStyle     =   0
      FormatString    =   "���֤��,1,0,3"
      Count           =   1
      titleInputBox0001=   "���֤��"
      statusinfoInputBox0001=   ""
      lengthInputBox0001=   3
      orderInputBox0001=   1
      valueInputBox0001=   ""
      datatypeInputBox0001=   2
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
      PermitNullInputBox0001=   0   'False
      TriggerstrInputBox0001=   ""
      �����ѡInputBox0001=   0   'False
      ErrColor        =   12648447
   End
   Begin VB.Frame cframJBXX 
      BackColor       =   &H80000013&
      Caption         =   "�Ǽǻ�����Ϣ��"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1485
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   10215
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6030
         TabIndex        =   26
         Top             =   420
         Width           =   255
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         ItemData        =   "FrmRegisterFirst.frx":0000
         Left            =   4440
         List            =   "FrmRegisterFirst.frx":0002
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker cdtpDate 
         Height          =   300
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
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
         Format          =   23592961
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.TextBox ctxtName 
         Height          =   300
         Left            =   240
         TabIndex        =   0
         Top             =   1080
         Width           =   2010
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "FrmRegisterFirst.frx":0004
         Left            =   2520
         List            =   "FrmRegisterFirst.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   840
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "��λ(&L)"
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox ctxtAge 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   15
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   7
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label clblTubeNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������뿴״̬��"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6285
         TabIndex        =   22
         Top             =   450
         Width           =   1650
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5520
         TabIndex        =   21
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   270
         Width           =   900
      End
      Begin VB.Label clblSysNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
         Height          =   180
         Index           =   1
         Left            =   5520
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
         Left            =   8280
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
         Left            =   4440
         TabIndex        =   14
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   3600
         TabIndex        =   13
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame frmPhoto 
      Caption         =   "����"
      ForeColor       =   &H00800000&
      Height          =   4365
      Left            =   5800
      TabIndex        =   23
      Top             =   2640
      Width           =   4575
      Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
         Height          =   3720
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   6562
         BackColor       =   0
         FontSize        =   9.75
         OriginalSize    =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7080
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15849
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctlbTool 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1032
      ButtonWidth     =   820
      ButtonHeight    =   926
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkClear 
         Caption         =   "��������"
         Height          =   435
         Left            =   6885
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   3480
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "FrmRegisterFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��˺�
'����޸ģ��

Private mobj��� As Object                   '������
Private mobj����ģ�� As Object             '����ģ�����
Private mobj���� As Object                 '�������

'ҵ�����á�
Private mblnTakePhoto As Boolean             '�Ƿ���Ҫ����
Private mbln����¼�� As Boolean

Private mblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö��� '����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mcolTubeNo As New Collection          '�Թܱ�ż�

Private mstr��λ������ As String

Private mblnSys As Boolean

Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub ccmbUnit_Click()
    Dim i As Integer
    Dim lcolInfo As New Collection
    
    On Error GoTo errHandler
    If ccmbUnit.Text = "" Then Exit Sub
    
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

Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '�Զ����档
    If ctlbTool.Buttons(4).Enabled Then
'        ctxtName.SetFocus
'        SendKeys "{F2}"
        mobjGUI_BeforeOperate "����", blnCancel
    End If
End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c�ֵ��" Then
        c�ֵ��.Visible = False
    End If

End Sub


Private Sub ctxtAge_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    '�ж��Ƿ�Ϊ����
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = 13 Then
    Else
        sffuncMsg "����ֻ��Ϊ���֣�������¼�롣", sf����
        KeyAscii = 0
    End If
End Sub

Private Sub ctxtAge_LostFocus()
    On Error GoTo errHandler
    '�ƶ����㡣
    If ctxtAge <> "" Then
        If Val(ctxtAge) > 150 Or Val(ctxtAge) <= 0 Then
            sffuncMsg "�������>0�����Ҳ����ܳ���150���������������䡣", sf����
            ctxtAge.SetFocus
        End If
    End If
    Exit Sub

errHandler:
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnTakePhoto Then
        '���³�ʼ������ؼ���
        cctlCatchPhoto.funcInitVideo
    End If
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    gfsubHideComboList ccmbUnit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '�Ƚ�����
    If KeyCode = vbKeyF8 And mblnTakePhoto Then
        If cctlCatchPhoto.VideoIsOk Then
            cctlCatchPhoto.subת��״̬
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

'���ܣ������ʼ����
Private Sub Form_Load()
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then
        Exit Sub
    End If
    MousePointer = 11
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    csbMain.Panels(1) = "�������ڳ�ʼ�������Ժ�..."
    
    '���治�ɲ�����
    cframJBXX.Enabled = False
    ciptBase.Enabled = False
    frmPhoto.Enabled = False
    ctlbTool.Enabled = False
       
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    mobjGUI.pbln�Զ������ֵ�߶� = False
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    Dim lcol��������ť As New Collection           '�������ϵİ�ť��ʼ�����ϡ�
    With lcol��������ť
        .Add "���"
        .Add "�޸�"
        .Add "|"
        .Add "����"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlbTool
        Set .c¼��� = ciptBase
        Set .c�ֵ�� = c�ֵ��
        
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcol��������ť, ""
    End With
    
    Set mobj��� = CreateObject("�����󲿼�.clsMedicalExam")
    Set mobj����ģ�� = CreateObject("�����󲿼�.ClsMedicalExamTemplate")
    
    '���
    subClear

    'Ϊ�˼ӿ촰������ٶȣ����³�ʼ���������ڶ�ʱ������ɡ�
    Timer1.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterFirst", "Form_Load", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = "�����ʼ��ʧ�ܣ�"
    ctlbTool.Enabled = True
    Exit Sub
    Resume
End Sub


'���ܣ�Ϊ�˼ӿ촰������ٶȣ�������³�ʼ��������
Private Sub Timer1_Timer()
    Dim lobj����ģ�弯 As Object           '����ģ�弯����
    Dim lcol����ģ�弯 As Collection
    Dim lcol��λ���Ƽ� As Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    Timer1.Enabled = False
    
    '���뵥λ����
    Set lcol��λ���Ƽ� = pobjҵ�����.���չ������䲾.��λ���Ƽ�
    ccmbUnit.Clear
    For i = 1 To lcol��λ���Ƽ�.Count
        ccmbUnit.AddItem lcol��λ���Ƽ�(i)
    Next i
    Set lcol��λ���Ƽ� = Nothing
    
    '�����еķǸ�������ģ����뵽ccmb������
    Set lobj����ģ�弯 = CreateObject("�����󲿼�.ClsMedicalExamTemplateSet")
    lobj����ģ�弯.�������� = 3
    Set lcol����ģ�弯 = lobj����ģ�弯.Ԫ�ؼ�
    For i = 1 To lcol����ģ�弯.Count
        ccmbTemplate.AddItem lcol����ģ�弯(i)
    Next i
    Set lcol����ģ�弯 = Nothing
    Set lobj����ģ�弯 = Nothing
    
    '���õ�һ������Ϊȱʡ������
    mblnSys = True
    If ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.ListIndex = 0
    Else
        sffuncMsg "������������ʱ�޷����У����Ƚ��롰�������á��������棬���ø�����������ݣ�", sf����
    End If
    mblnSys = False
    
    '��ȡҵ�����á�
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
    
    '�������������ø�����Ϣ¼��塣
    sub��������
    
    '��ȡ��ǰ�����Ѿ�ʹ�õ��Թܱ����ĸ��
    If mobj���.����.�Թܱ����ĸ <> "" Then
        '��ǰ�����ѹ̶�ʹ����ĳ����ĸ��
        clblLetter.Caption = mobj���.����.�Թܱ����ĸ
        cvscLetter.Enabled = False
    Else
        mobj���.����.�Թܱ����ĸ = clblLetter.Caption
    End If
    
    '��Ҫ����ʱ��ʼ������ؼ�
    If mblnTakePhoto Then
        '��ʼ���ؼ�
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    '����ϵͳ���
    clblSysNo.Caption = mobj���.Func����ϵͳ���
    
    '�ָ�����ɲ�����
    cframJBXX.Enabled = True
    ciptBase.Enabled = True
    frmPhoto.Enabled = True
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmRegisterFirst", "Form_Load", 6666, lstrError, False
    Else
        ctxtName.SetFocus
    End If
    ctlbTool.Enabled = True
    mblnSys = False
    csbMain.Panels(1) = ""
    
    '�жϳ�ʼ���Ƿ�ɹ���
    If mblnTakePhoto Then
        If Not cctlCatchPhoto.VideoIsOk Then
            csbMain.Panels(1) = "�����豸��ʼ��ʧ�ܣ�����ԭ���ҵ�����������ò��������࣡"
        End If
    End If
    MousePointer = 0
    
    Exit Sub
    Resume
End Sub


Private Sub cvscLetter_Change()
    On Error GoTo errHandler
    
    '����������������Ӧ����ĸ��
    clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
        
    Exit Sub
errHandler:
    'sfsub������ "�����沿��", "FrmRegisterFirst", "cvscLetter_Scroll", Err.Number, Err.Description, False
End Sub

Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '�ƶ����㡣
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If

End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal ���� As String, ByVal ���� As String, ByVal �������� As String, ByVal IsError As Boolean)
    Dim lstrIDCard As String
    Dim ldatBirth As String
    Dim lstrSex As String
    Dim i As Integer
    
    On Error GoTo errHandler
    ldatBirth = ""
    Select Case ����
    Case "���֤��"
        lstrIDCard = ����
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
                ctxtAge = DateDiff("yyyy", ldatBirth, Date)
            End If
            If lstrSex <> "" Then
                ccmbSex.Text = lstrSex
            End If
        End If
    Case "��������"
        Dim lstrItemText As String
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
            If Val(����) >= Val(ctxtAge) Then
                Err.Raise 6666, , "����>=���䣬���ǷǷ������ݣ�"
            End If
        End If
        
    Case Else
    End Select
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterFirst", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    ciptBase.ItemSetFocus Index
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj��� = Nothing
    Set mobj����ģ�� = Nothing
    '�ر������
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    mblnInUse = False

End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    '�ƶ����㡣
    If KeyCode = 13 Then
        ctxtAge.SetFocus
    End If
    Exit Sub
errHandler:
End Sub

'���ܣ�ѡ������
Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    
    If mobj���.����.������ = ccmbTemplate.Text Or mblnSys Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1).Text = "���ڻ�ȡ����ģ����Ϣ�����Ժ�..."
    
    sub��������
    
    '�޸ģ�2001-8-23����ʾ��λ���ԣ���
    On Error Resume Next
    If mstr��λ������ <> "" Then
        sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
    End If
    
    ctxtName.SetFocus
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    Err.Clear
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterFirst", "ccmbTemplate_Click", 6666, lstrError, False
    
    csbMain.Panels(1).Text = ""
    MousePointer = 0
    Exit Sub
    Resume
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbUnit
    csbMain.Panels(1) = "Ҫ��ձ��ؼ�¼�ĵ�λ�����嵥����ɾ���ļ���c:\temp\������칤�����䲾.ini����"
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    '�ƶ����㡣
    If KeyCode = 13 Then
        '��¼���û��¼����Ŀ����ֱ�ӱ��档
        If mobj����ģ��.����������Ŀ��.Count > 0 Then
            ciptBase.ItemSetFocus 0
        Else
            ciptBase_LastLostFocus
        End If
    Else
        mstr��λ������ = ""
    End If
    Exit Sub
errHandler:
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
        If mstr��λ������ <> pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text) Then
            pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
        End If
    End If

    Exit Sub
errHandler:
End Sub

Private Sub ccmdLocateUnit_Click()
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��
    Dim lcolInfo As Collection
    
    On Error GoTo errHandler
    
    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ccmbUnit.Text = lobjRec("��λ����")
            mstr��λ������ = lobjRec!������
            
            '��ʾ�������Ա���ڵ�λ���������ԡ�
            '�޸ģ�2001-8-23�������������
            On Error Resume Next
            sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
        End If
    End If
    
    '�ѽ���ص���λ¼���
    ccmbUnit.SetFocus
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterFirst", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub

Private Sub ctxtAge_GotFocus()
    On Error GoTo errHandler
    With ctxtAge
        .SelStart = 0
        .SelLength = Len(Trim(ctxtAge.Text))
    End With
    Exit Sub
errHandler:
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '�ƶ����㡣
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub ctxtName_GotFocus()
    On Error GoTo errHandler
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
    Exit Sub
errHandler:

End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '�ƶ����㡣
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub
Private Sub subClear()
    On Error Resume Next

    'clblTubeNo.Caption = ""
    ctxtName.Text = ""
    ctxtAge.Text = ""
    cdtpDate.Value = Date
    ccmbUnit.Text = ""
    ciptBase.ClearContent
    Set cctlCatchPhoto.Photo = Nothing
    
End Sub

'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    On Error GoTo errHandler
    
    Select Case Operate
    Case "���"
        '��ս��档
        subClear
        
        '���Խ���ͨ�ö������µĴ���
        Cancel = True
        
    Case "�޸�"
        Dim lstr��һ���� As String
        
        '�������¼���ϵͳ��š�
        If mobj���.ϵͳ��� = "" Then
            On Error Resume Next
            lstr��һ���� = mobj���.func��ȡϵͳ��ŵ�ǰһ����(clblSysNo.Caption)
            On Error GoTo errHandler
        Else
            lstr��һ���� = mobj���.ϵͳ���
        End If
        '�ر������
        If mblnTakePhoto Then
            cctlCatchPhoto.subDisconnect
        End If
        
        FrmEditRegister.ϵͳ��� = lstr��һ����
        '�����޸Ľ��档
        FrmEditRegister.Move Me.Left, Me.Top
        FrmEditRegister.Show 1
        
        '���¿��������
        If mblnTakePhoto Then
            cctlCatchPhoto.funcInitVideo
        End If
        
        '���Խ���ͨ�ö������µĴ���
        Cancel = True
        
    Case "����"
        '���¼���Ƿ��д���
        If mobj����ģ��.����������Ŀ��.Count > 0 Then
            '�޸ģ�2001-9-12�������
            On Error Resume Next
            ciptBase.Box1(ciptBase.ActiveInputBoxIndex).LostFocus
            On Error GoTo errHandler
            
            If ciptBase.ItemsError.Count > 0 And Not mbln����¼�� Then
                sffuncMsg "�������ɫ¼������ݣ�", sf����
                ciptBase.SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
         '����150�걨��
        If Val(ctxtAge.Text) > 150 Or (ctxtAge.Text <> "" And Val(ctxtAge.Text) <= 0) Then
            sffuncMsg "��������䲻�Ϸ�������������Ϸ������䣨<150����", sf����
            With ctxtAge
                .SelStart = 0
                .SetFocus
            End With
            Cancel = True
            Exit Sub
        End If
        '�ж��Ƿ���Ҫ���ࡣ
        If mblnTakePhoto = True Then
            '�ж��Ƿ�����
            If cctlCatchPhoto.PhotoIsOk = False Then
                sffuncMsg "û�����࣬����������󱣴棡", sf����
                Cancel = True
                Exit Sub
            End If
        End If
        MousePointer = 11
        csbMain.Panels(1) = "���ڱ������Ǽ���Ϣ�����Ժ�..."
        
        '������ʱ���ɲ�����
        ctlbTool.Enabled = False
        cframJBXX.Enabled = False
        ciptBase.Enabled = False
        frmPhoto.Enabled = False
        
        '�������������ԡ�
        With mobj���
            .ϵͳ��� = clblSysNo.Caption
            .����.������ = ccmbTemplate.Text
            .����.�Թܱ����ĸ = clblLetter.Caption
            .�����Ա.���� = Trim(ctxtName.Text)
            .�����Ա.�Ա� = ccmbSex.Text
            .�����Ա.��λ���� = Trim(ccmbUnit.Text)
            
            
            .������� = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            If mblnTakePhoto Then
                .�����Ա.��Ƭ = cctlCatchPhoto.Photo
            End If
            If Val(ctxtAge) > 0 Then
                .�����Ա.�������� = DateAdd("yyyy", -Val(ctxtAge), Date)
            End If
            On Error Resume Next
            .�����Ա.������ݺ��� = ciptBase.Box1("���֤��").Text
            .�����Ա.�������� = ciptBase.Box1("��������").TrueText
            .�����Ա.Ƭ�� = ciptBase.Box1("Ƭ��").TrueText
            .�����Ա.��ҵ��� = ciptBase.Box1("��ҵ���").TrueText
            
            If .�����Ա.��λ������ <> mstr��λ������ Then
                .�����Ա.��λ������ = mstr��λ������
            End If
            On Error GoTo errHandler
            
            '���ø�����Ŀ�����
            For i = 1 To ciptBase.ItemCount
                '�����ֵ�¼�룬��ֵ������Ŀֵ��š�
                If ciptBase.InfoCollection(i).�ֵ����� <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
                    .����.Sub�����Ϣֵ ciptBase.InfoCollection(i).����, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
                Else
                    .����.Sub�����Ϣֵ ciptBase.InfoCollection(i).����, ciptBase.Box1(i - 1).Text
                End If
            Next
        End With
        
        'ͨ��������ҵ�����ִ�б��档
        pobjҵ�����.Sub���Ǽ� mobj���
        
        csbMain.Panels(2) = "�ϴα�������ϵͳ��ţ�" & mobj���.ϵͳ��� & "���Թܱ�ţ�" & mobj���.�Թܱ��
        If mobj���.�շ����� <> "" Then
            csbMain.Panels(2) = csbMain.Panels(2) & "���շ����ţ�" & mobj���.�շ�����
        End If
        
        If cchkClear = 1 Then
            subClear
        End If
        
        '�ָ����ࡣ
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "�ָ�" Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
        
        '�·���ϵͳ���
        clblSysNo.Caption = mobj���.Func����ϵͳ���
        
        '�ָ�����ɲ�����
        ctlbTool.Enabled = True
        cframJBXX.Enabled = True
        ciptBase.Enabled = True
        frmPhoto.Enabled = True
        
        '�Թ���ĸ������ѡ��
        cvscLetter.Enabled = False
        
        ctxtName.SetFocus
        
        MousePointer = 0
        csbMain.Panels(1) = ""
        
        '���Խ���ͨ�ö������µĴ���
        Cancel = True
    Case "�˳�"
        If Trim(ctxtName) <> "" Then
            If Not sffuncMsg("��ȷ��Ҫ�˳������棬���Ҳ����浱ǰ��¼��������Ա�Ǽ���Ϣ��", sfѯ��) Then
                Cancel = True
                Exit Sub
            End If
        End If
        If clblSysNo.Caption <> "" Then
            '�˻�ϵͳ��š�
            mobj���.sub�˻�ϵͳ��� clblSysNo.Caption
        End If
    End Select
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterFirst", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    If Operate = "����" Then
        '�ָ�����ɲ�����
        ctlbTool.Enabled = True
        cframJBXX.Enabled = True
        ciptBase.Enabled = True
        frmPhoto.Enabled = True
    End If
    Cancel = True
    Exit Sub
    Resume

End Sub


'���ܣ��������������б��ǰ�������������������������ԣ�
'      ���ݴ��ж��Ƿ���������Թܱ��,��ʼ��������Ϣ��
Private Sub sub��������()
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer
    
    On Error GoTo errHandler
    '��ȡ���Թܱ�š�
    mobj���.����.������ = ccmbTemplate.Text
    
    '��������ģ���ȡ���������п��õ���ĸ��
    If mobj����ģ��.������ <> ccmbTemplate.Text Then
        mobj����ģ��.������ = ccmbTemplate.Text
    End If
    
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
            
            '��ֵ��clblLetter��
            clblLetter.Caption = mcolTubeNo(1)
            cvscLetter.Min = 1
            cvscLetter.Max = mcolTubeNo.Count
            cvscLetter.Enabled = True
            cvscLetter.Value = 1
            
        Else
            ctlbTool.Buttons(4).Enabled = False
            If ccmbTemplate.Text <> "" Then
                '��ʾ�������޿��õ���ĸ��
                Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
            Else
                Exit Sub
            End If
        End If
    Else
        '����ĸ������ѡ����ĸ��
        clblLetter.Caption = mobj���.����.�Թܱ����ĸ
        cvscLetter.Enabled = False
        
    End If
    
    '��ʼ��������Ϣ��
    On Error Resume Next
    mobjGUI.sub��ʼ��¼��� ccmbTemplate.Text
    
    ctlbTool.Buttons(4).Enabled = True
    Exit Sub
    
errHandler:
    
    sfsub������ "�����沿��", "FrmRegisterFirst", "sub��������", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

