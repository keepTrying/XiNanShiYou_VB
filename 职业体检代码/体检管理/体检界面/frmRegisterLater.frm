VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "¼��ؼ�.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmRegisterLater 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��¼���Ǽ���Ϣ"
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
   ScaleHeight     =   7455
   ScaleWidth      =   10470
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   3120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7080
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18415
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkˢ���� 
         Caption         =   "ˢ����"
         Height          =   375
         Left            =   6840
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox cchkClear 
         Caption         =   "��������"
         Height          =   375
         Left            =   8520
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Value           =   1  'Checked
         Width           =   1530
      End
   End
   Begin VB.Frame cframJBXX 
      BackColor       =   &H80000013&
      Caption         =   "�Ǽǻ�����Ϣ:"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   6195
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   10400
      Begin VB.TextBox ctxt��쵥�� 
         Height          =   315
         Left            =   8760
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox ctxtTubeNo 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame frmPhoto 
         Caption         =   "����"
         ForeColor       =   &H00800000&
         Height          =   4305
         Left            =   5640
         TabIndex        =   29
         Top             =   1560
         Width           =   4695
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3570
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   6297
            BackColor       =   0
            FontSize        =   9.75
            OriginalSize    =   -1  'True
         End
      End
      Begin VB.TextBox ctxtSysNo 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2100
      End
      Begin ¼��ؼ�.ctlInputDictGrid c�ֵ�� 
         Height          =   4335
         Left            =   5640
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   7646
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
      Begin ¼��ؼ�.ctlInputFrame ciptBase 
         Height          =   4455
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   7858
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
         BorderStyle     =   0
         Caption         =   "Frame1"
         Rows            =   6
         Cols            =   27
         DistanceofRow   =   0
         BorderStyle     =   0
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
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   4200
         TabIndex        =   5
         Top             =   1080
         Width           =   3360
      End
      Begin VB.TextBox ctxtAge 
         Height          =   300
         Left            =   3480
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "��λ(&L)"
         Height          =   375
         Left            =   7560
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1020
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "frmRegisterLater.frx":0000
         Left            =   2280
         List            =   "frmRegisterLater.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   960
      End
      Begin VB.TextBox ctxtName 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2100
      End
      Begin VB.Frame Frame2 
         Height          =   75
         Left            =   30
         TabIndex        =   16
         Top             =   1440
         Width           =   10335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥�ţ�"
         Height          =   180
         Index           =   9
         Left            =   8760
         TabIndex        =   34
         Top             =   840
         Width           =   900
      End
      Begin VB.Label clbl������� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8760
         TabIndex        =   32
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ͣ�"
         Height          =   180
         Index           =   8
         Left            =   8760
         TabIndex        =   31
         Top             =   240
         Width           =   900
      End
      Begin VB.Label clblDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7200
         TabIndex        =   28
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label clblTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2280
         TabIndex        =   26
         Top             =   480
         Width           =   3165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   7
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   3480
         TabIndex        =   22
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���ƣ�"
         Height          =   180
         Index           =   5
         Left            =   4200
         TabIndex        =   21
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ڣ�"
         Height          =   180
         Index           =   2
         Left            =   7200
         TabIndex        =   20
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
         Height          =   180
         Index           =   1
         Left            =   5640
         TabIndex        =   19
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   4
         Left            =   2280
         TabIndex        =   17
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton ccmdPre 
      Appearance      =   0  'Flat
      Caption         =   "<"
      Height          =   450
      Left            =   945
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin VB.CommandButton ccmdFirst 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Height          =   450
      Left            =   450
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin VB.CommandButton ccmdNext 
      Appearance      =   0  'Flat
      Caption         =   ">"
      Height          =   450
      Left            =   1455
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin VB.CommandButton ccmdLast 
      Appearance      =   0  'Flat
      Caption         =   ">>"
      Height          =   450
      Left            =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4125
      Width           =   510
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   1560
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRegisterLater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��
'����޸ģ��

Private WithEvents mobjGUI As cls����ͨ�ö��� '����ͨ�ö���������ʼ��������������¼��塣
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj��� As Object                    '�����������������Ǽ���Ϣ��

Private mstrϵͳ��Ź̶����� As String
Private mbln����¼�� As Boolean              'ҵ�����ã��Ƿ����¼�루���ǿ���¼�룬����¼��Ϸ��Լ�飩��
Private mblnTakePhoto As Boolean             'ҵ�����ã��Ƿ����ࡣ
Private mblnChangedPhoto As Boolean

Private mstr��λ������ As String

Private mblnInUse As Boolean

'�޸ģ�2003-4-15�����Ӹ�ģ�鼶ȫ�ֶ���Ϊ�˻�ȡˢ�����������ֵ����
Private mobj����  As cls�û���������

'���ܣ���ȡ�������Ƿ��������ı�־��
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchkˢ����_Click()
    On Error GoTo errhandler
    
    If cchkˢ����.Value = 1 Then
        '���ϵͳ��������
        ctxtSysNo = ""
    Else
        ctxtSysNo = mstrϵͳ��Ź̶�����
    End If
    mobj���.ϵͳ��� = ctxtSysNo.Text
    ctxtSysNo.SelStart = Len(ctxtSysNo)
    ctxtSysNo.SelLength = 0
    ctxtSysNo.SetFocus
    Exit Sub
errhandler:
End Sub

Private Sub ccmbUnit_Click()
    On Error GoTo errhandler
    Dim i As Integer
    If ccmbUnit.Text = "" Then Exit Sub  'Ϊ��ʱ�������б�
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '���뵽�б����
        ccmbUnit.AddItem ccmbUnit.Text
        
        '���ص��������䲾�ļ���
        pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
    Else
        '�޸ģ�2001-8-23(��ʾ��λ����)��
        On Error Resume Next
        mstr��λ������ = pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text)
        sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
    
    End If
    Exit Sub
errhandler:
End Sub

Private Sub ciptBase_ItemLostFocus(Index As Integer)
    On Error Resume Next
    If ciptBase.InfoCollection(Index + 1).Title = "��������" Then
        'mobjGUI_ItemLostFocus Index, "��������", ciptBase.ItemText(Index), ciptBase.ItemTrueText(Index), False
    End If

End Sub

Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '�Զ����档
    If ctbMain.Buttons(3).Enabled Then
'        ctxtSysNo.SetFocus
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

Private Sub ctxtTubeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub ctxt��쵥��_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '��¼���û��¼����Ŀ����ֱ�ӱ��档
        If mobj���.����.������Ϣ.Count > 0 Then
            ciptBase.SetFocus
        Else
            ciptBase_LastLostFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtSysNo.SetFocus
    ctxtSysNo.SelStart = Len(ctxtSysNo)
    ctxtSysNo.SelLength = 0
    
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


Private Sub Form_Load()
    Dim lcolInfo As Collection
    Dim i As Integer
    
    On Error GoTo errhandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1) = "�������ڳ�ʼ�������Ժ�..."
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '������ʱ���ɲ�����
    cframJBXX.Enabled = False
    ctbMain.Enabled = False
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    mobjGUI.pbln�Զ������ֵ�߶� = False
    
    Set lcolInfo = New Collection
    With lcolInfo
        .Add "��ѯ(&Q)105"
        .Add "|"
        .Add "����"
        .Add "�����Ƭ(&A)111"
        .Add "������Ƭ(&E)103"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
        Set .c¼��� = ciptBase
        Set .c�ֵ�� = c�ֵ��
        
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcolInfo, ""
    End With

    '���뵥λ����
    Set lcolInfo = pobjҵ�����.���չ������䲾.��λ���Ƽ�
    For i = 1 To lcolInfo.Count
        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    '����������
    Set mobj��� = CreateObject("������.clsMedicalExam")
    mstrϵͳ��Ź̶����� = mobj���.ϵͳ��Ź̶�����
    mobj���.ϵͳ��� = mstrϵͳ��Ź̶�����
    
    ctxtSysNo = mstrϵͳ��Ź̶�����
    ctxtSysNo.SelLength = Len(ctxtSysNo)
    ctxtSysNo.SelStart = 0
    
    ctbMain.Buttons(3).Enabled = False
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    If pobjҵ�����.ҵ������("�Ƿ���ٵǼ�") = "��" Then
        mbln����¼�� = True
    Else
        mbln����¼�� = False
    End If

    If pobjҵ�����.ҵ������("�Ƿ�����") = "��" Then
        mblnTakePhoto = True
    Else
        mblnTakePhoto = False
    End If
    
    
    '��Ҫ����ʱ��ʼ������ؼ�
    If mblnTakePhoto Then
        '��ʼ���ؼ�
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
        ctxtTubeNo.Enabled = True
    Else
        ctxtTubeNo.Enabled = False
    End If
        
    
    '�����������
    On Error Resume Next
    Set mobj���� = New cls�û���������
    mobj����.�û���� = um�û����
    mobj����.ҵ���� = "��첹¼�Ǽ�"
    Dim lstrOption As String
    lstrOption = mobj����.������ֵ("ˢ����")
    If lstrOption = "��" Then
        cchkˢ����.Value = 1
    End If
    
    Err.Clear
    
errhandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "frmRegisterLater", "Form_Load", 6666, lstrError, False
    End If
    '����ɲ�����
    cframJBXX.Enabled = True
    ctbMain.Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    
    Exit Sub
    Resume
End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtAge.SetFocus
    End If
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error GoTo errhandler
    '��ý���ʱ����������
    gfsubShowComboList ccmbUnit
    Exit Sub
errhandler:
End Sub

Private Sub ccmbUnit_LostFocus()
    On Error GoTo errhandler
    Dim i As Integer
    If Trim(ccmbUnit.Text) = "" Then Exit Sub  'Ϊ��ʱ�������б�
    
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
errhandler:
    
End Sub

Private Sub ccmdLocateUnit_Click()
    On Error GoTo errhandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��

    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ccmbUnit.Text = lobjRec("��λ����")
            mstr��λ������ = lobjRec!������
            
            '�޸ģ�2001-8-23����ʾ��λ���ԣ���
            On Error Resume Next
            sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
            
        End If
    End If
    
    '�ѽ���ص���λ¼���
    ccmbUnit.SetFocus
    SendKeys vbTab
    Exit Sub
errhandler:
    'sfsub������ "�����沿��", "frmRegisterLater", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub



Private Sub ctxtAge_GotFocus()
    On Error GoTo errhandler
    With ctxtAge
        .SelStart = 0
        .SelLength = Len(Trim(ctxtAge.Text))
    End With
    Exit Sub
errhandler:
    'sfsub������ "�����沿��", "frmRegisterLater", "ctxtAge_GotFocus", Err.Number, Err.Description, False
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub ctxtAge_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler
    '�ж��Ƿ�Ϊ����
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = 13 Then
    Else
        sffuncMsg "����ֻ��Ϊ���֣�������¼�롣", sf����
        KeyAscii = 0
    End If
    Exit Sub
    
errhandler:
End Sub

Private Sub ctxtAge_LostFocus()
    On Error GoTo errhandler
    '�ƶ����㡣
    If ctxtAge <> "" Then
        If Val(ctxtAge) > 150 Or Val(ctxtAge) <= 0 Then
            sffuncMsg "�������>0�����Ҳ����ܳ���150���������������䡣", sf����
            ctxtAge.SetFocus
        End If
    End If
    Exit Sub
errhandler:
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub

Private Sub ctxtSysNo_GotFocus()
    On Error Resume Next
    With ctxtSysNo
        '��ʾϵͳ��ŵĹ̶����֡�
        If ctxtSysNo.Text = "" And cchkˢ����.Value = 0 Then
            ctxtSysNo.Text = mstrϵͳ��Ź̶�����
        End If
        .SelStart = Len(Trim(ctxtSysNo.Text))
        .SelLength = 0
    End With
End Sub

Private Sub ctxtSysNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '��ʾ�����Ա�Ļ�����Ϣ���ƶ����㡣
    If KeyCode = 13 Then
        If ctxtTubeNo.Enabled Then
            ctxtTubeNo.SetFocus
        Else
            ctxtName.SetFocus
        End If
    End If
End Sub

Private Sub ctxtSysNo_LostFocus()
    Dim lcol��츽����Ŀ As New Collection
    Dim lstr�Թܱ�� As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo errhandler
    
    '�����δ�仯����������
    If mobj���.ϵͳ��� = ctxtSysNo.Text Or (Len(ctxtSysNo.Text) = Len(mstrϵͳ��Ź̶�����) And mstrϵͳ��Ź̶����� <> "") Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1) = "�������������Ա��Ϣ�����Ժ�..."
    
    ctbMain.Buttons(3).Enabled = False
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    mobj���.ϵͳ��� = Trim(ctxtSysNo.Text)
    ctxtName.Text = ""
    cctlCatchPhoto.subClear
    mblnChangedPhoto = False
    
    If Not mobj���.�Ƿ��Ѵ��� Then
        '����¼������ʱ������ʾ
        Err.Raise 6666, , "������Ų����ڣ���������������š�"
    Else
        '�ж��Ƿ����������ۡ�
        If mobj���.���״̬ = P_ENDED_STATUS Then
            Err.Raise 6666, , "���������ۣ�������¼���Ǽ���Ϣ��"
        End If
        '�ж��Ƿ񸴲�����¼��
        If mobj���.������ = P_EXAM_AGAIN Then
            Err.Raise 6666, , "�����¼����������¼��"
        End If
        
        '���ҵ���������д�Ǽ���Ϣ�ڽ����ϡ�
        With mobj���
            ctxt��쵥�� = .��쵥��
            
            ctxtName.Text = .�����Ա.����
            i = gffuncItemIsInComboBox(ccmbSex, .�����Ա.�Ա�)
            ccmbSex.ListIndex = i
            
            If IsDate(.�����Ա.��������) Then
                ctxtAge.Text = DateDiff("yyyy", .�����Ա.��������, Date)
            Else
                ctxtAge.Text = ""
            End If
            ccmbUnit.Text = .�����Ա.��λ����
            If clblTemplate.Caption <> .����.������ Then
                clblTemplate.Caption = .����.������
                '���³�ʼ��������Ϣ¼��塣
                On Error Resume Next
                mobjGUI.sub��ʼ��¼��� clblTemplate.Caption
                On Error GoTo errhandler
            End If
            ctxtTubeNo.Text = .�Թܱ��
            
            If IsDate(.�������) Then
                clblDate.Caption = Format(.�������, "yyyy-mm-dd")
            Else
                clblDate.Caption = .�������
            End If
            '�޸ģ�2001-8-23��
            mstr��λ������ = .�����Ա.��λ������
            
            '�޸ģ�2001-12-29����ʾ������ͣ���
            If .������ = P_EXAM_ANNUAL Then
                clbl�������.Caption = "���"
            Else
                clbl�������.Caption = "����"
            End If
        End With
        
                
        '��ȡ���и�����Ŀ��������
        Set lcol��츽����Ŀ = mobj���.����.������Ϣ
        
        '��д������Ϣ�����
        sub��¼���ֵ ciptBase, mobjGUI, lcol��츽����Ŀ
        
        '��ò���ʾ��Ƭ��
        If Not mobj���.�����Ա.��Ƭ Is Nothing Then
            Set cctlCatchPhoto.Photo = mobj���.�����Ա.��Ƭ
        Else
            cctlCatchPhoto.subClear
        End If
        
        If ctxtTubeNo.Enabled Then
            ctxtTubeNo.SetFocus
        Else
            ctxtName.SetFocus
        End If
    End If
    ctbMain.Buttons(3).Enabled = True
    ctbMain.Buttons(4).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    
errhandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmRegisterLater", "ctxtSysNo_LostFocus", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    If ctxtSysNo.Enabled Then
        ctxtSysNo.SetFocus
    End If
    ctxtSysNo = mstrϵͳ��Ź̶�����
    ctxtSysNo.SelStart = Len(ctxtSysNo)
    ctxtSysNo.SelLength = 0
    Exit Sub
    Resume
End Sub
Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt��쵥��.SetFocus
    Else
        mstr��λ������ = ""
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj��� = Nothing
    Set mobjGUI = Nothing
    
    '�ر������
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    
    '�޸ģ�2003-4-15������������䣩��
    mobj����.sub���Ǽ���ֵ "ˢ����", IIf(cchkˢ����.Value = 1, "��", "��")
    
    Set mobj���� = Nothing
    
    mblnInUse = False
End Sub

'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Dim i As Integer
    Dim lstr��ˮ�� As String
    Dim lstrFile As String
    
    On Error GoTo errhandler
    
    Select Case Operate
    
    Case "��ѯ"
        '��ʾ�������Ա�Ļ�����Ϣ��
        Dim lstrϵͳ��� As String
        frm������Ա.pstr������� = "����"
        frm������Ա.Show 1, Me
        lstrϵͳ��� = frm������Ա.pstrϵͳ���
        If lstrϵͳ��� <> "" Then
            ctxtSysNo.Text = lstrϵͳ���
            '��ʾ�����Ա������Ϣ��
            ctxtSysNo_LostFocus
        End If
    
    
    Case "����"
         '����150�걨��
        If Val(ctxtAge.Text) > 150 Or (ctxtAge.Text <> "" And Val(ctxtAge.Text) <= 0) Then
            sffuncMsg "��������䲻�Ϸ������������롣", sf����
            ctxtAge.SetFocus
            Cancel = True
            Exit Sub
        End If
        If Trim(ctxtTubeNo.Text) = "" Then
            sffuncMsg "�Թܱ�ű������룡", sf����
            ctxtTubeNo.SetFocus
            Cancel = True
            Exit Sub
        End If
        
        '���¼���Ƿ��д���
        If mobj���.����.������Ϣ.Count > 0 Then
            '�޸ģ�2001-9-12�������
            On Error Resume Next
            ciptBase.Box1(ciptBase.ActiveInputBoxIndex).LostFocus
            On Error GoTo errhandler
            
            If ciptBase.ItemsError.Count > 0 And Not mbln����¼�� Then
                sffuncMsg "�������ɫ¼������ݣ�", sf����
                Cancel = True
                Exit Sub
            End If
        End If
        MousePointer = 11
        csbMain.Panels(1) = "���ڱ������Ǽ���Ϣ�����Ժ�..."
        
        '�����Թܱ�Ų�����
        With mobj���
            .�����Ա.���� = Trim(ctxtName.Text)
            .�����Ա.�Ա� = ccmbSex.Text
            .�����Ա.��λ���� = Trim(ccmbUnit.Text)
            
            If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
                .�Թܱ�� = ctxtTubeNo.Text
            End If
            .��쵥�� = ctxt��쵥��.Text
            
            If Val(ctxtAge) > 0 Then
                .�����Ա.�������� = DateAdd("yyyy", -Val(ctxtAge), Date)
            End If
            If mblnTakePhoto Or mblnChangedPhoto Then
                .�����Ա.��Ƭ = cctlCatchPhoto.Photo
            End If
            On Error Resume Next
            .�����Ա.������ݺ��� = ciptBase.Box1("���֤��").Text
            .�����Ա.�������� = ciptBase.Box1("��������").TrueText
            .�����Ա.Ƭ�� = ciptBase.Box1("Ƭ��").TrueText
            .�����Ա.��ҵ��� = ciptBase.Box1("��ҵ���").TrueText
            
            If .�����Ա.��λ������ <> mstr��λ������ Then
                .�����Ա.��λ������ = mstr��λ������
            End If
            
            On Error GoTo errhandler
            '���渽����Ϣ
            For i = 1 To ciptBase.ItemCount
                'If ciptBase.Box1(i - 1).TrueText <> ciptBase.Box1(i - 1).Text And ciptBase.Box1(i - 1).Text <> "" Then
                If ciptBase.InfoCollection(i).�ֵ����� <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
                    .����.Sub�����Ϣֵ ciptBase.InfoCollection(i).����, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
                Else
                    .����.Sub�����Ϣֵ ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
                End If
            Next i
            
            .�Թܱ�� = ctxtTubeNo.Text
        End With
        
        pobjҵ�����.Sub���Ǽ� mobj���
        
        If cchkClear = 1 Then
            subClear
            mobj���.ϵͳ��� = ""
            ctbMain.Buttons(3).Enabled = False
            ctbMain.Buttons(4).Enabled = False
            ctbMain.Buttons(5).Enabled = False
            
        End If
        
        '�ָ����ࡣ
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "�ָ�" Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
        
        mblnChangedPhoto = False
        
        ctxtSysNo.SelStart = Len(ctxtSysNo)
        ctxtSysNo.SelLength = 0
        ctxtSysNo.SetFocus
        
        MousePointer = 0
        csbMain.Panels(1) = ""
        Cancel = True
        
    Case "�����Ƭ"
        ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ", vbDirectory) <> "" Then
            ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ"
        End If
        ccmdFile.FileName = ctxtSysNo.Text
        ccmdFile.ShowOpen
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            SavePicture cctlCatchPhoto.Photo, lstrFile
        End If
    Case "������Ƭ"
        ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ", vbDirectory) <> "" Then
            ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ"
        End If
        ccmdFile.FileName = ctxtSysNo.Text
        ccmdFile.ShowOpen
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            If InStr(lstrFile, ".") > 0 Then
                Set cctlCatchPhoto.Photo = LoadPicture(lstrFile)
                mblnChangedPhoto = True
            End If
        End If
    End Select
    
    Exit Sub
errhandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "frmRegisterLater", "mobjGUI_BeforeOperate", 6666, lstrError, False
    End If
    MousePointer = 0
    csbMain.Panels(1) = ""
    Cancel = True
    Exit Sub
    Resume
    Exit Sub
End Sub
'���
Private Sub subClear()
    On Error Resume Next
    If cchkˢ����.Value = 1 Then
        ctxtSysNo.Text = ""
    Else
        ctxtSysNo.Text = mstrϵͳ��Ź̶�����
    End If
    clblTemplate.Caption = ""
    ctxtTubeNo.Text = ""
    ctxtName.Text = ""
    If ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
    ctxtAge.Text = ""
    ccmbUnit.Text = ""
    cctlCatchPhoto.subClear
    ciptBase.ClearContent
    
End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal ���� As String, ByVal ���� As String, ByVal �������� As String, ByVal IsError As Boolean)
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    
    On Error GoTo errhandler
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
                ctxtAge = DateDiff("yyyy", ldatBirth, Date)
            End If
            If lstrSex <> "" Then
                ccmbSex.Text = lstrSex
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
            If Val(����) >= Val(ctxtAge) Then
                Err.Raise 6666, , "����>=���䣬���ǷǷ������ݣ�"
            End If
        End If
        
    End Select
    Exit Sub
errhandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterLater", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemBox(Index).Text = ""
    ciptBase.ItemSetFocus Index
    Exit Sub
    Resume

End Sub
