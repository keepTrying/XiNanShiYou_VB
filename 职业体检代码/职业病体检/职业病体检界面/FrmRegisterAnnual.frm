VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "¼��ؼ�.ocx"
Begin VB.Form FrmRegisterAnnual 
   BorderStyle     =   0  'None
   Caption         =   "���Ǽ�--���Ǽ�"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8229.114
   ScaleMode       =   0  'User
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   3600
      Top             =   480
   End
   Begin VB.ListBox clstPersonList 
      Height          =   1680
      Left            =   5280
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Frame cfram������Ϣ 
      BackColor       =   &H80000013&
      Caption         =   "�Ǽǻ�����Ϣ:"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   5595
      Left            =   60
      TabIndex        =   21
      Top             =   1920
      Width           =   10380
      Begin VB.Frame frmPhoto 
         Caption         =   "����"
         ClipControls    =   0   'False
         ForeColor       =   &H00800000&
         Height          =   4035
         Left            =   5760
         TabIndex        =   40
         Top             =   1560
         Width           =   4575
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3720
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   6562
            BackColor       =   0
            FontSize        =   9.75
            OriginalSize    =   -1  'True
         End
      End
      Begin VB.CommandButton ccmd��λ��λ 
         Caption         =   "��λ(&T)"
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   945
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Top             =   1200
         Width           =   3120
      End
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6480
         TabIndex        =   10
         Top             =   600
         Width           =   345
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   3120
      End
      Begin ¼��ؼ�.ctlInputDictGrid c�ֵ�� 
         Height          =   3735
         Left            =   5640
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   6588
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
         Height          =   315
         Left            =   8760
         TabIndex        =   39
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   72024065
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin ¼��ؼ�.ctlInputFrame ciptBase 
         Height          =   3780
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6668
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
         Rows            =   5
         Cols            =   27
         DistanceofRow   =   0
         BorderStyle     =   0
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
      Begin VB.Label clblAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1920
         TabIndex        =   37
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label clblSex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label clblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   1320
         TabIndex        =   34
         Top             =   990
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   990
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2760
         TabIndex        =   30
         Top             =   315
         Width           =   720
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������뿴״̬��"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6840
         TabIndex        =   29
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6120
         TabIndex        =   28
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   315
         Width           =   900
      End
      Begin VB.Label clblSysNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
         Height          =   180
         Index           =   1
         Left            =   6120
         TabIndex        =   25
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ڣ�"
         Height          =   180
         Index           =   2
         Left            =   8760
         TabIndex        =   24
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���ƣ�"
         Height          =   180
         Index           =   5
         Left            =   2760
         TabIndex        =   23
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   1920
         TabIndex        =   22
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.Frame cframSearch 
      Caption         =   "���������Ա��"
      ForeColor       =   &H00800000&
      Height          =   1020
      Left            =   45
      TabIndex        =   20
      Top             =   840
      Width           =   10395
      Begin VB.OptionButton coptChoise 
         Caption         =   "���֤��"
         Height          =   240
         Index           =   2
         Left            =   4320
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox ctxtId 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   7
         Top             =   600
         Width           =   2025
      End
      Begin VB.ComboBox ccmbQueryUnit 
         Height          =   300
         Left            =   4920
         TabIndex        =   2
         Top             =   225
         Width           =   3735
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "FrmRegisterAnnual.frx":0000
         Left            =   3240
         List            =   "FrmRegisterAnnual.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox ctxtHealthNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   2025
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "����֤��"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton ccmdSearch 
         Caption         =   "��ʾ��Ա(F5)"
         Height          =   360
         Left            =   8760
         TabIndex        =   8
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox ctxtName 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   990
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "����"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "��λ(&L)"
         Height          =   375
         Left            =   8760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ��λ"
         Height          =   180
         Index           =   7
         Left            =   4320
         TabIndex        =   32
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2760
         TabIndex        =   31
         Top             =   300
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   7455
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16007
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkClear 
         Caption         =   "��������"
         Height          =   345
         Left            =   5160
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Value           =   1  'Checked
         Width           =   1410
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   1680
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
Attribute VB_Name = "FrmRegisterAnnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��˺�
'����޸ģ��

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

'���ܣ����ص�ǰ�����Ƿ��Ѽ��أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub ccmbTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        cdtpDate.SetFocus
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
        '��¼���û��¼����Ŀ����ֱ�ӱ��档
        If mobj���.����.������Ϣ.Count > 0 Then
            ciptBase.SetFocus
        Else
            ciptBase_LastLostFocus
        End If
    Else
        mstr��λ������ = ""
    End If
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
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
    If ctbMain.Buttons(4).Enabled Then
        ccmbUnit.SetFocus
        SendKeys "{F2}"
        'mobjGUI_BeforeOperate "����", blnCancel
    End If
End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c�ֵ��" Then
        c�ֵ��.Visible = False
    End If

End Sub

'���ܣ����һ����ţ���ʾ����ϵͳ��ŵ������Ա����Ϣ��
Private Sub clstPersonList_DblClick()
    Dim lobj�����Ա  As Object 'clsPersonExamed.
    Dim lobj��� As Object      'clsMedicalExam
    Dim lstrItem As String      'ѡ�������ݡ�����������š�ϵͳ��š�
    On Error GoTo errHandler
    
    With clstPersonList
        '�������������š�
        lstrItem = .List(.ListIndex)
        lstrItem = Left(lstrItem, InStr(lstrItem, " ") - 1)
        
        '��ȡ�������һ������¼��ϵͳ��š�
        If coptChoise(0).Value Then
            lstrItem = func���ݽ���������Ż�ȡϵͳ����(lstrItem)
        End If
        
        '��ʾ�����Ա��Ϣ��
        SubGetPersonInfo lstrItem
        
        '���б����ʧ��
        .Visible = False
        
    End With
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "clstPersonList_Click", 6666, lstrError, False
    
End Sub

Private Sub clstPersonList_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And clstPersonList.ListIndex >= 0 Then
        clstPersonList_DblClick
    End If
End Sub

Private Sub ctxtId_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
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
    gfsubHideComboList ccmbQueryUnit
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
   
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    MousePointer = 11
    csbMain.Panels(1) = "�������ڳ�ʼ�������Ժ�..."
    
    '���治�ɲ�����
    cframSearch.Enabled = False
    cfram������Ϣ.Enabled = False
    ctbMain.Enabled = False
    
    
    Set mobj����� = CreateObject("�����󲿼�.clsMedicalExam")
    Set mobj��� = CreateObject("�����󲿼�.clsMedicalExam")
    Set mobj��켯 = CreateObject("�����󲿼�.clsMedicalExamSet")
    Set mobj����ģ�� = CreateObject("�����󲿼�.ClsMedicalExamTemplate")
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
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
        Set .c������ = ctbMain
        Set .c¼��� = ciptBase
        Set .c�ֵ�� = c�ֵ��
        'Set .c״̬�� = csbMain
        
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcol��������ť, ""
    End With
    
    '���
    subClear
    
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
    csbMain.Panels(1) = "�����ʼ��ʧ�ܣ�"
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
        ccmbQueryUnit.AddItem lcolInfo(i)
        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    '�����еķǸ�������ģ����뵽���������б���С�
    Set lobj����ģ�弯 = CreateObject("�����󲿼�.ClsMedicalExamTemplateSet")
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
    
    '��Ҫ����ʱ��ʼ������ؼ���
    If mblnTakePhoto Then
        '��ʼ���ؼ���
        cctlCatchPhoto.funcInitVideo
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pobjҵ�����.ҵ������("�Ƿ���ٵǼ�") = "��" Then
        mbln����¼�� = True
    Else
        mbln����¼�� = False
    End If
     
    '�ָ�����ɲ�����
    cframSearch.Enabled = True
    cfram������Ϣ.Enabled = True
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmRegisterAnnual", "Timer1_Timer", 6666, lstrError, False
    End If
    ctbMain.Enabled = True
    MousePointer = 0
    csbMain.Panels(1) = ""
    If cframSearch.Enabled Then
        ctxtName.SetFocus
    End If
End Sub


'���ܣ��Զ������б��
Private Sub ccmbQueryUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbQueryUnit
    Exit Sub
errHandler:
End Sub

Private Sub ccmbQueryUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ccmbQueryUnit_LostFocus()
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbQueryUnit, ccmbQueryUnit.Text)
    
    If i = -1 Then
        '�ӵ�ccmbQueryUnit�С�
        ccmbQueryUnit.AddItem ccmbQueryUnit.Text
    End If
    
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbQueryUnit.SetFocus
    End If
End Sub

Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    'ѡ������
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    MousePointer = 11
    csbMain.Panels(1).Text = "���ڻ�ȡ����ģ����Ϣ�����Ժ�..."
    
    '��ȡ���Թܱ�š�
    If mobj���.����.������ <> ccmbTemplate.Text Then
        mobj���.����.������ = ccmbTemplate.Text
        
        '��������ģ���ȡ���������п��õ���ĸ��
        mobj����ģ��.������ = ccmbTemplate.Text
        
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
                ctbMain.Buttons(4).Enabled = False
                '��ʾ�������޿��õ���ĸ��
                Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
            End If
        Else
            '����ĸ������ѡ����ĸ��
            clblLetter.Caption = mobj���.����.�Թܱ����ĸ
            cvscLetter.Enabled = False
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
        DoEvents
    End If
    ctbMain.Buttons(4).Enabled = True
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "ccmbTemplate_Click", 6666, lstrError, False
    ciptBase.pblnTemp = False
    Exit Sub
    Resume
End Sub

'�Զ������б��
Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbUnit
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
        If mstr��λ������ <> pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text) Then
            pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
        End If
    End If
    Exit Sub
errHandler:
End Sub
'���õ�λ��λ
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��

    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ccmbQueryUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    
    '�ѽ���ص���λ¼���
    ccmbQueryUnit.SetFocus
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub

Private Sub ccmdSearch_Click()
    Dim lobj�����Ա As Object  '������ȡ�����Ա���������¼��
    Dim lobj��� As Object      '�����Ա�������졣
    Dim lobjRec As Object       'Recordset����켯���󷵻ص�Ԫ�ؼ���
    Dim lstrϵͳ��� As String  '����֤�Ŷ�Ӧ�����ϵͳ��š�
    Dim i As Integer
    
    On Error GoTo errHandler
    MousePointer = 11
    csbMain.Panels(1) = "���ڲ�������ָ�������������Ա�����Ժ�..."
    
    '����ս��档
    subClear
    
    If coptChoise(0).Value Then
        If Trim(ctxtName.Text) = "" And ccmbSex.Text = "" And Trim(ccmbQueryUnit.Text) = "" Then
            Err.Raise 6666, , "���������������Ա𡢵�λ���ƣ�"
        End If
        '������켯�����������λ�������ԡ�
        With mobj��켯
            .subClear
            .���� = Trim(ctxtName.Text)
            .�Ա� = ccmbSex.Text
            .��λ���� = Trim(ccmbQueryUnit.Text)
        End With
        
        '��ȡ����ָ����λ����������¼��
        Set lobjRec = mobj��켯.Ԫ�ؼ�("distinct �����������,����,��λ����")
        If lobjRec.recordcount = 0 Then
            'û�ҵ���Ӧ�����Ա��
            Err.Raise 6666, , "�������Ա��û���ڱ���������������޷��������Ǽǡ���ѡ�����Ǽǡ�"
        Else
            If lobjRec.recordcount > 1 Then
                '���ҵ�������¼ʱ����list�������뵽�����Ա�б���С�
                clstPersonList.Clear
                Do While Not lobjRec.EOF
                    clstPersonList.AddItem lobjRec("�����������") & " " & lobjRec("����") & " " & IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
                    lobjRec.movenext
                Loop
                '�б�ɼ�
                clstPersonList.Visible = True
                clstPersonList.SetFocus
                
            Else
                'ֻ�ҵ�һ����,��ȡ�������һ������¼��ϵͳ��š�
                lstrϵͳ��� = func���ݽ���������Ż�ȡϵͳ����(lobjRec!�����������)
            End If
        End If
        
        lobjRec.Close
        
    ElseIf coptChoise(1).Value Then
        '������֤��ת����ϵͳ��š�
        If Trim(ctxtHealthNo.Text) = "" Then
            Err.Raise 6666, , "��������뽡��֤�ţ�"
        End If
        lstrϵͳ��� = pobjҵ�����.Func���ݽ���֤����Ż�ȡ���ϵͳ���(Trim(ctxtHealthNo.Text))
        If lstrϵͳ��� = "" Then
            Err.Raise 6666, , "������Ľ���֤��û�ж�Ӧ������¼��"
        End If
    Else
        '�����֤�Ų�ѯ��
        If Trim(ctxtId.Text) = "" Then
            Err.Raise 6666, , "������������֤�ţ�"
        End If
        '������켯�����������λ�������ԡ�
        With mobj��켯
            .subClear
            .���֤�� = Trim(ctxtId.Text)
        End With
        
        '��ȡ����ָ����λ����������¼��
        Set lobjRec = mobj��켯.Ԫ�ؼ�("ϵͳ���,����,��λ����")
        If lobjRec.recordcount = 0 Then
            'û�ҵ���Ӧ�����Ա��
            Err.Raise 6666, , "�������Ա��û���ڱ���������������޷��������Ǽǡ���ѡ�����Ǽǡ�"
        ElseIf lobjRec.recordcount = 1 Then
            lstrϵͳ��� = lobjRec("ϵͳ���")
        Else
            '���ҵ�������¼ʱ����list�������뵽�����Ա�б���С�
            clstPersonList.Clear
            Do While Not lobjRec.EOF
                clstPersonList.AddItem lobjRec("ϵͳ���") & " " & lobjRec("����") & " " & IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
                lobjRec.movenext
            Loop
            '�б�ɼ�
            clstPersonList.Visible = True
            clstPersonList.SetFocus
            
        End If
    End If
    
    '��ʾ�������Ա�Ļ�����Ϣ��
    If lstrϵͳ��� <> "" Then
        SubGetPersonInfo lstrϵͳ���
    End If
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmRegisterAnnual", "ccmdSearch_Click", 6666, lstrError, False
        
        If coptChoise(0).Value Then
            ctxtName.SetFocus
        ElseIf coptChoise(1).Value Then
            ctxtHealthNo.SetFocus
        Else
            ctxtId.SetFocus
        End If
    End If
    Set lobj�����Ա = Nothing
    Set lobj��� = Nothing
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
    Case vbKeyF5
        '��ʾ��Ա��
        ccmdSearch_Click
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
        If lobjRec.recordcount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
            mstr��λ������ = lobjRec!������
            
            If mstr��λ������ <> "" Then
                '�޸ģ�2001-8-23����ʾ��λ���ԣ���
                On Error Resume Next
                sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
            End If
        End If
    End If
    
    '�ѽ���ص���λ¼���
    ccmbUnit.SetFocus
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub


Private Sub clstPersonList_LostFocus()
    On Error Resume Next
    clstPersonList.Visible = False
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
                clblAge.Caption = DateDiff("yyyy", ldatBirth, Date)
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
            If Val(����) >= Val(clblAge) Then
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

Private Sub coptChoise_Click(Index As Integer)
    On Error GoTo errHandler
    ctxtName.Enabled = False
    ccmbSex.Enabled = False
    ccmbQueryUnit.Enabled = False
    ccmdLocateUnit.Enabled = False
    ctxtId.Enabled = False
    ctxtHealthNo.Enabled = False
    
    If coptChoise(0).Value Then
        'ѡ������������
        ctxtName.Enabled = True
        ccmbSex.Enabled = True
        ccmbQueryUnit.Enabled = True
        ccmdLocateUnit.Enabled = True
        ctxtName.SetFocus
    ElseIf coptChoise(1).Value Then
        'ѡ�����뽡��������š�
        ctxtHealthNo.Enabled = True
        ctxtHealthNo.SetFocus
    ElseIf coptChoise(2).Value Then
        'ѡ���������֤�š�
        ctxtId.Enabled = True
        ctxtId.SetFocus
    End If
    
    Exit Sub
errHandler:
    'sfsub������ "�����沿��", "FrmRegisterAnnual", "coptChoise_Click", Err.Number, Err.Description, False
End Sub

Private Sub ctxtHealthNo_GotFocus()
    On Error Resume Next
    With ctxtHealthNo
        .SelStart = 0
        .SelLength = Len(Trim(ctxtHealthNo.Text))
    End With
End Sub

Private Sub ctxtHealthNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ctxtName_GotFocus()
    On Error Resume Next
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
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
    ccmbTemplate.Text = ""
    clblLetter.Caption = ""
    cdtpDate.Value = Date
    clblName.Caption = ""
    clblSex.Caption = ""
    clblAge.Caption = ""
    ccmbUnit.Text = ""
    ciptBase.ClearContent
    cfram������Ϣ.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set mobj��� = Nothing
    Set mobj��켯 = Nothing
    Set mobj����ģ�� = Nothing
    '�ر������
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    
    mblnInUse = False
End Sub


'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim i As Integer
    Dim lstr��ˮ�� As String
    On Error GoTo errHandler
    
    Select Case Operate
    Case "���"
        subClear
        Cancel = True
    Case "�޸�"
        Dim lstr��һ���� As String
        
        '�ر������
        If mblnTakePhoto Then
            cctlCatchPhoto.subDisconnect
        End If
        
        '�������¼���ϵͳ��š�
        lstr��һ���� = mobj���.ϵͳ���
        If Not mobj���.�Ƿ��Ѵ��� And lstr��һ���� <> "" Then
            lstr��һ���� = mobj���.func��ȡϵͳ��ŵ�ǰһ����(lstr��һ����)
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
        '�ж��Ƿ���Ҫ���ࡣ
        If mblnTakePhoto = True Then
            '�ж��Ƿ�����
            If cctlCatchPhoto.Photo Is Nothing Then
                Err.Raise 6666, , "û�����࣬����������󱣴棡"
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
        MousePointer = 11
        csbMain.Panels(1) = "���ڱ������Ǽ���Ϣ�����Ժ�..."
        
        '�����Թܱ�Ų�����
        With mobj���
            If .����.������ <> ccmbTemplate.Text Then
                .����.������ = ccmbTemplate.Text
            End If
            If .����.�Թܱ����ĸ <> clblLetter.Caption Then
                .����.�Թܱ����ĸ = clblLetter.Caption
            End If
            .�����Ա.���� = clblName.Caption
            .�����Ա.�Ա� = clblSex.Caption
            .�����Ա.��λ���� = ccmbUnit.Text
            
            If mblnTakePhoto Then
                .�����Ա.��Ƭ = cctlCatchPhoto.Photo
            End If
            If Val(clblAge.Caption) > 0 Then
                .�����Ա.�������� = DateAdd("yyyy", -Val(clblAge.Caption), Date)
            End If
            
            On Error Resume Next
            .�����Ա.������ݺ��� = ciptBase.Box1("���֤��").Text
            .�����Ա.�������� = ciptBase.Box1("��������").TrueText
            .�����Ա.Ƭ�� = ciptBase.Box1("Ƭ��").TrueText
            .�����Ա.��ҵ��� = ciptBase.Box1("��ҵ���").TrueText
            If .�����Ա.��λ������ <> mstr��λ������ Then
                .�����Ա.��λ������ = mstr��λ������
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
            
            '����Ϊ��졣
            .������ = P_EXAM_ANNUAL
            .������� = Format(cdtpDate.Value, "yyyy-mm-dd")
        End With
        On Error GoTo errHandler
        
        pobjҵ�����.Sub���Ǽ� mobj���
        
        csbMain.Panels(2) = "�ϴα�������ϵͳ��ţ�" & mobj���.ϵͳ��� & "���Թܱ�ţ�" & mobj���.�Թܱ�� & "��"
        
        If cchkClear = 1 Then
            subClear
        End If
        
        clblSysNo.Caption = ""
        
        '�ָ����ࡣ
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "�ָ�" Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
        
        '���水ť�����á�
        ctbMain.Buttons(4).Enabled = False
        
        If coptChoise(0).Value Then
            ctxtName.SetFocus
        Else
            ctxtHealthNo.SetFocus
        End If
        
        '�Թ���ĸ������ѡ��
        cvscLetter.Enabled = False
        
        Cancel = True
        csbMain.Panels(1) = ""
        MousePointer = 0
     Case "�˳�"
        '����������¼û�б��棬�˻�ϵͳ��š�
        If mobj���.ϵͳ��� <> "" And Not mobj���.�Ƿ��Ѵ��� Then
            If Not sffuncMsg("��ȷ��Ҫ�˳������棬���Ҳ����浱ǰ��¼��������Ա�Ǽ���Ϣ��", sfѯ��) Then
                Cancel = True
                Exit Sub
            End If
            
            '�˻�ϵͳ��š�
            mobj���.sub�˻�ϵͳ��� mobj���.ϵͳ���
        End If
        
        'ȡ������ͨ�ö�����˳���ť�Ĵ���
        Set mobjGUI.Form = Nothing
        Unload Me
    End Select
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1) = ""
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
    csbMain.Panels(1) = "������ʾ��ǰ�����Ա����Ϣ�����Ժ�..."
    
    '������ʱ���ɲ�����
    ctbMain.Enabled = False
    cfram������Ϣ.Enabled = False
    
    '���˻ؾ�ϵͳ��š�
    If Not mobj���.�Ƿ��Ѵ��� And mobj���.ϵͳ��� <> "" Then
        mobj���.sub�˻�ϵͳ��� mobj���.ϵͳ���
    End If
    
    '������������
    Set mobj����� = CreateObject("�����󲿼�.clsMedicalExam")
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
        clblName.Caption = .����
        clblSex.Caption = .�Ա�
        clblAge.Caption = .����
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
'        If mstr��λ������ <> "" Then
'            sub��ʾ��λ���� ciptBase, mstr��λ������
'        End If
        On Error GoTo errHandler
    End With
    
    '�����µ�ϵͳ���
    lstrSysNo = mobj���.Func����ϵͳ���
    mobj���.ϵͳ��� = lstrSysNo
    clblSysNo.Caption = lstrSysNo
    
    '�����������䡣
    mobj���.�����Ա.����������� = mobj�����.�����Ա.�����������
    
    
    '�����������������Ӷ���ȡ���Թܱ�š�
    mobj���.����.������ = ccmbTemplate.Text
    
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
            ctbMain.Buttons(4).Enabled = False
            '��ʾ�������޿��õ���ĸ��
            Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
        End If
    Else
        '����ĸ������ѡ����ĸ��
        cvscLetter.Enabled = False
    End If
    
    '¼�������Բ�����
    cfram������Ϣ.Enabled = True
    
    '���水ť���á�
    ctbMain.Buttons(4).Enabled = True
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        sfsub������ "�����沿��", "FrmRegisterAnnual", "SubGetPersonInfo", Err.Number, Err.Description, True
    End If
    
    '�ָ�����ɲ�����
    ctbMain.Enabled = True
    cframSearch.Enabled = True
    MousePointer = 0
    csbMain.Panels(1) = ""
    
    Exit Sub
    Resume
End Sub
    
Private Function func���ݽ���������Ż�ȡϵͳ����(ByVal para����������� As String) As String
    Dim lobj�����Ա  As Object 'clsPersonExamed.
    Dim lobj��� As Object      'clsMedicalExam
    Dim lstrϵͳ��� As String
    
    On Error GoTo errHandler
    
    '��ȡ�������һ������¼��
    '���������Ա����
    Set lobj�����Ա = CreateObject("�����󲿼�.clsPersonExamed")
    lobj�����Ա.����������� = para�����������
    Set lobj��� = lobj�����Ա.Func��ȡ�������һ�����
    If Not lobj��� Is Nothing Then
        lstrϵͳ��� = lobj���.ϵͳ���
    Else
        Err.Raise 6666, , "�������Ա��û���ڱ���������������޷��������Ǽǡ���ѡ�����Ǽǡ�"
    End If
            
    func���ݽ���������Ż�ȡϵͳ���� = lstrϵͳ���
    
    Exit Function
errHandler:
    sfsub������ "�����沿��", "FrmRegisterAnnual", "func���ݽ���������Ż�ȡϵͳ����", Err.Number, Err.Description, True
    Exit Function
    Resume

End Function
