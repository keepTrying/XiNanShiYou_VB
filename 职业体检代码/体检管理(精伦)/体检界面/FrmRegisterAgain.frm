VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Begin VB.Form FrmRegisterAgain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����Ǽ�"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10335
   ClipControls    =   0   'False
   Icon            =   "FrmRegisterAgain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   2400
      Top             =   600
   End
   Begin VB.TextBox ctxtTemplate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   2175
   End
   Begin VB.PictureBox cpicPhoto 
      Height          =   1785
      Left            =   360
      ScaleHeight     =   1725
      ScaleWidth      =   1365
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1425
   End
   Begin MSComctlLib.Toolbar ctblTool 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkPrint 
         Caption         =   "��ӡ��쵥"
         Height          =   345
         Left            =   6720
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CheckBox cchkClear 
         Caption         =   "��������"
         Height          =   345
         Left            =   4920
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1425
      End
   End
   Begin VB.Frame cframBase 
      Caption         =   "�Ǽǻ�����Ϣ"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   10095
      Begin VB.TextBox ctxtTubeNo 
         Height          =   315
         Left            =   5280
         TabIndex        =   26
         Top             =   480
         Width           =   2415
      End
      Begin VB.ListBox clstItem 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   3300
         ItemData        =   "FrmRegisterAgain.frx":0442
         Left            =   120
         List            =   "FrmRegisterAgain.frx":0444
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2190
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
         Height          =   300
         Left            =   2520
         TabIndex        =   0
         Top             =   1080
         Width           =   2475
         _ExtentX        =   4366
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
         Format          =   20119552
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.VScrollBar cvscLetter 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5775
         TabIndex        =   2
         Top             =   450
         Width           =   435
      End
      Begin ¼��ؼ�.ctlInputFrame ciptBase 
         Height          =   3780
         Left            =   2520
         TabIndex        =   14
         Top             =   2280
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   6668
         BackColor       =   15791081
         Enabled         =   -1  'True
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
         ����InputBox0001=   "���֤��"
         ȱʡֵInputBox0001=   ""
         ����ȱʡֵInputBox0001=   ""
         ����InputBox0001=   0
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         EnableInputBox0001=   0   'False
         �����ѡInputBox0001=   0   'False
         ErrColor        =   15791081
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ��"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label clblUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         TabIndex        =   22
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label clblAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3720
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label clblSex 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label clblName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         TabIndex        =   19
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���ƣ�"
         Height          =   180
         Index           =   8
         Left            =   5280
         TabIndex        =   17
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   9
         Left            =   3720
         TabIndex        =   18
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   7
         Left            =   5280
         TabIndex        =   16
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   6
         Left            =   2520
         TabIndex        =   15
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڣ�"
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
         Height          =   180
         Index           =   1
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   900
      End
      Begin VB.Label clblSysNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         TabIndex        =   5
         Top             =   450
         Width           =   495
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������뿴״̬��"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6210
         TabIndex        =   4
         Top             =   450
         Width           =   1545
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7860
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15610
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
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
Attribute VB_Name = "FrmRegisterAgain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��

Private WithEvents mobjGUI As cls����ͨ�ö���    '����ͨ�ö������ڳ�ʼ����������¼���ؼ����ơ�
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj��� As Object                       '������

Public pstr��ϵͳ��� As String                 '��¼ԭ����¼��ϵͳ��š�
Private mstrϵͳ��� As String                   '���������ĸ�������¼��ϵͳ��š�
Private mcolTubeNo As New Collection             '��ǰ���������ѡ����Թ���ĸ��

Private mstr�ϸ������� As String
Private mstrϵͳ��Ź̶����� As String

'ҵ�����á�
Private mbln����¼�� As Boolean

Private mblnInUse As Boolean                     '��Ӧ����pblnInUse��

Private mcol�շ���Ŀ As Collection               'item:���,key����š�

'���ܣ����ص�ǰ�����Ƿ��Ѽ��أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��
Public Property Get pblnInUse() As Boolean
    On Error GoTo errHandler
    pblnInUse = mblnInUse
    Exit Property
errHandler:
    'sfsub������ "�����沿��", "FrmRegisterAgain", "Property Get pblnInUse", Err.Number, Err.Description, True
End Property

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

'���ܣ������ʼ������ʼ�������������µĳ�ʼ�������ڶ�ʱ��Timer1_timer����ɣ���
Private Sub Form_Load()
    On Error GoTo errHandler

    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
        
    '��ʾ�ȴ�״̬��
    MousePointer = 11
    csbMain.Panels(1) = "�������ڳ�ʼ�������Ժ�..."
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    mstr�ϸ������� = ""
    
    '��ʼ������ͨ�ö���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    Dim lcol��������ť As New Collection           '�������ϵİ�ť��ʼ�����ϡ�
    With lcol��������ť
        .Add "ѡ����Ŀ(&I)111"
        .Add "����"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctblTool
        Set .c״̬�� = csbMain
        Set .c¼��� = ciptBase
        
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcol��������ť, ""
    End With
    
    '��ʼʱ�����水ť�����á�
    ctblTool.Buttons(1).Enabled = False
    ctblTool.Buttons(2).Enabled = False
    
    cdtpDate.Value = Format(Date, "yyyy-mm-dd")
    
    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
        ctxtTubeNo.Visible = True
        ctxtTubeNo.TabIndex = 1
        clblLetter.Visible = False
        cvscLetter.Visible = False
    Else
        ctxtTubeNo.Visible = False
        clblLetter.Visible = True
        cvscLetter.Visible = True
    End If
        
    'Ϊ�˼ӿ촰������ٶȣ����³�ʼ���������ڶ�ʱ������ɡ�
    Timer1.Enabled = True
    Exit Sub
    
errHandler:
    '������
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAgain", "Form_Load", 6666, lstrError, False
    
End Sub

'���ܣ����form_load���µĴ����ʼ��������
Private Sub Timer1_Timer()
    Dim lcolInfo As Collection   '���չ������䲾�б���ĵ�λ���Ƽ���
    Dim lobjRec As Object        '��ҵ������ȡ�Ĵ�����������
    Dim i As Integer
       
    On Error GoTo errHandler
    
    '��ʱ�����������á�
    Timer1.Enabled = False
    
    '������ʱ���ɲ�����
    cframBase.Enabled = False
    ciptBase.Enabled = False
    ctblTool.Enabled = False
    
    '����������
    Set mobj��� = CreateObject("������.clsMedicalExam")
    mstrϵͳ��Ź̶����� = mobj���.ϵͳ��Ź̶�����
    
    '�ж�ҵ�������Ƿ��ӡ��쵥��
    If pobjҵ�����.ҵ������("�Ƿ��ӡ��쵥") = "��" Then
        cchkPrint.Visible = True
    End If
    
    '��ʾ������Ա�����Ϣ��
    SubGetPersonInfo
    
        
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmRegisterAgain", "Timer1_Timer", 6666, lstrError, False
        '���治�ɲ�����
        cframBase.Enabled = False
    End If
    
    '�ָ�����ɲ�����
    cframBase.Enabled = True
    ciptBase.Enabled = True
    ctblTool.Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub

'���
Private Sub subClear()
    On Error Resume Next
    clblSysNo.Caption = ""
    clblLetter.Caption = ""
    ctxtTubeNo.Text = ""
    cdtpDate.Value = Date
    ciptBase.ClearContent
    clstItem.Clear
    clblName.Caption = ""
    clblSex.Caption = ""
    clblAge.Caption = ""
    clblUnit.Caption = ""
    cframBase.Caption = "�Ǽǻ�����Ϣ"
    'ѡ����Ŀ�����水ť�����á�
    ctblTool.Buttons(1).Enabled = False
    ctblTool.Buttons(2).Enabled = False

    Set mcol�շ���Ŀ = New Collection
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '����������¼û�б��棬�˻�ϵͳ��š�
    If Not mobj��� Is Nothing Then
        If mobj���.ϵͳ��� <> "" And Not mobj���.�Ƿ��Ѵ��� Then
            '�˻�ϵͳ��š�
            mobj���.sub�˻ظ���ϵͳ��� mobj���.ϵͳ���
        End If
    End If
    
    '�ͷŶ���
    Set mobj��� = Nothing
    
    '���ô���û�������ı�־��
    mblnInUse = False

End Sub


Private Sub cvscLetter_Change()
    On Error Resume Next
    '����������������Ӧ����ĸ��
    If mcolTubeNo.Count > 0 Then
        clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
    End If
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    Dim lstr��ˮ�� As String
    
    On Error GoTo errHandler
    
    Select Case Operate
    Case "ѡ����Ŀ"
        Dim lcol������Ŀ As Collection
        Dim lobj���ģ�� As Object
        
        '��ȡ���������е������Ŀ��
        Set lcol������Ŀ = mobj���.����.�����Ŀ��("")
        
        '����ѡ����Ŀ��������ԡ�
        frmSelectItem.pstr�������� = ctxtTemplate.Text
        Set frmSelectItem.pcol������Ŀ = lcol������Ŀ
        Set frmSelectItem.pcol�շ���Ŀ = mcol�շ���Ŀ
        '����ѡ����Ŀ���档
        frmSelectItem.Show 1
        If frmSelectItem.pblnOk Then
            '��ȡѡ�еĸ�����Ŀ��
            Set lcol������Ŀ = frmSelectItem.pcol������Ŀ
            '�޸�������������ʾ���б��С�
            clstItem.Clear
            mobj���.����.Subɾ�����������Ŀ
            For i = 1 To lcol������Ŀ.Count
                mobj���.����.Sub��������Ŀ lcol������Ŀ(i)("����")
                clstItem.AddItem lcol������Ŀ(i)("����") & " " & lcol������Ŀ(i)("����")
            Next
            
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
    Case "�޸�"
        FrmEditRegister.ϵͳ��� = mstrϵͳ���
        FrmEditRegister.Move Me.Left, Me.Top
        FrmEditRegister.Show 1
        Cancel = True
    Case "����"
        MousePointer = 11
        csbMain.Panels(1) = "���ڱ������Ǽ���Ϣ�����Ժ�..."
        With mobj���
            If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
                If .����.�Թܱ����ĸ <> clblLetter.Caption Then
                    .����.�Թܱ����ĸ = clblLetter.Caption
                End If
            Else
                .�Թܱ�� = ctxtTubeNo.Text
            End If
            .������� = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            '����������Ϊ�������Ǽǡ�
            .������ = P_EXAM_AGAIN
        End With
        
        '���渴��Ǽ���Ϣ��
        pobjҵ�����.Sub���Ǽ� mobj���, pstr��ϵͳ���, IIf(cchkPrint.Value = 1, True, False), mcol�շ���Ŀ
                
        '��¼��ǰ�������ϵͳ��š�
        mstrϵͳ��� = clblSysNo.Caption
        
        '��״̬������ʾ������ϵͳ��š��Թܱ�š�
        csbMain.Panels(2) = "�ϴα�������ϵͳ��ţ�" & mstrϵͳ��� & "���Թܱ�ţ�" & mobj���.�Թܱ�� & "��"
        
        If cchkClear = 1 Then
            subClear
        End If
        
        '�Թ���ĸ������ѡ��
        cvscLetter.Enabled = False
        
        csbMain.Panels(1) = ""
        MousePointer = 0
        Cancel = True
    Case "�˳�"
'        '����������¼û�б��棬�˻�ϵͳ��š�
'        If mobj���.ϵͳ��� <> "" And Not mobj���.�Ƿ��Ѵ��� Then
'            '�˻�ϵͳ��š�
'            mobj���.sub�˻ظ���ϵͳ��� mobj���.ϵͳ���
'        End If
        
    End Select
    Exit Sub
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterAgain", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    csbMain.Panels(1) = ""
    MousePointer = 0
    Cancel = True
    Exit Sub
    Resume
End Sub

'���ܣ���ʾָ��ϵͳ��ŵ������Ա����Ϣ�ڽ����ϡ�
Private Sub SubGetPersonInfo()
    Dim lobj��� As Object     'clsMedicalExam������¼��
    Dim lobj���ģ�� As Object 'clsMedicalExamTemplate
    Dim lcolInfo As New Collection
    Dim i As Integer
    Dim j As Integer
    
    Dim lstrTemp As String
    Dim lstrTubeNo As String
    Dim lstrSysNo As String
    
    On Error GoTo errHandler
    MousePointer = 11
    csbMain.Panels(1) = "������ʾ��ǰ������Ա����Ϣ�����Ժ�..."
    
    '������ʱ���ɲ�����
    ctblTool.Enabled = False
    
    '���˻ؾ�ϵͳ��š�
    If Not mobj���.�Ƿ��Ѵ��� And mobj���.ϵͳ��� <> "" Then
        mobj���.sub�˻ظ���ϵͳ��� mobj���.ϵͳ���
    End If
        
    '������������
    Set lobj��� = CreateObject("������.clsMedicalExam")
    lobj���.ϵͳ��� = pstr��ϵͳ���
    
    ctxtTemplate.Text = lobj���.������������
    
    '��������¼����������ʼ��¼��塣
    If mstr�ϸ������� <> lobj���.����.������ Then
        mstr�ϸ������� = lobj���.����.������
        '���³�ʼ��¼��塣
        On Error Resume Next
        mobjGUI.sub��ʼ��¼��� mstr�ϸ�������
        On Error GoTo errHandler
    End If
    
    '��������ģ�����
    Set lobj���ģ�� = CreateObject("������.clsMedicalExamTemplate")
    lobj���ģ��.������ = ctxtTemplate.Text
    
    '��ȡ��������������Ŀ��
    Set lcolInfo = lobj���ģ��.�����Ŀ��
    
    '��ʾ������Ŀ��
    clstItem.Clear
    For i = 1 To lcolInfo.Count
        clstItem.AddItem lcolInfo(i).���� & " " & lcolInfo(i).����
    Next i
    
    '��ȡ������¼�ĸ�����Ϣ��
    Set lcolInfo = lobj���.����.������Ϣ
    
    '��д������Ϣ�����
    sub��¼���ֵ ciptBase, mobjGUI, lcolInfo
    
    DoEvents
    
    '��ʾ������Ϣ��
    With lobj���.�����Ա
        clblName.Caption = .����
        clblSex.Caption = .�Ա�
        clblAge.Caption = .����
        clblUnit.Caption = .��λ����
        '��Ƭ
        cpicPhoto.Picture = .��Ƭ
    End With
    cframBase.Caption = "�Ǽǻ�����Ϣ��" & clblName.Caption & "��"
    '�����µ�ϵͳ���
    lstrSysNo = mobj���.Func���临��ϵͳ���(pstr��ϵͳ���)
    mobj���.ϵͳ��� = lstrSysNo
    clblSysNo.Caption = lstrSysNo
    
    '�����������䡣
    mobj���.�����Ա.����������� = lobj���.�����Ա.�����������
    
    
    '���ø����������������Ӷ���ȡ���Թܱ�š�
    mobj���.����.������ = ctxtTemplate.Text
    
    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
        '�Թܱ����ĸΪ��ʱcvscLetter����
        If mobj���.����.�Թܱ����ĸ = "" Then
            '����ĸ�����ŷֿ�������mcoltubeNo��
            lstrTubeNo = lobj���ģ��.�Թ���ĸ���
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
                ctblTool.Buttons(1).Enabled = False
                ctblTool.Buttons(2).Enabled = False
                '��ʾ�������޿��õ���ĸ��
                Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
            End If
        Else
            '����ĸ������ѡ����ĸ��
            clblLetter.Caption = mobj���.����.�Թܱ����ĸ
            cvscLetter.Enabled = False
        End If
    Else
        ctxtTubeNo = mobj���.�Թܱ��
    End If
    
    '���渽����Ϣ
    For i = 1 To ciptBase.ItemCount
        If ciptBase.InfoCollection(i).�ֵ����� <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
            mobj���.����.Sub�����Ϣֵ ciptBase.InfoCollection(i).����, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
        Else
            mobj���.����.Sub�����Ϣֵ ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
        End If
        
    Next i
            
    '���水ť���á�
    ctblTool.Buttons(1).Enabled = True
    ctblTool.Buttons(2).Enabled = True
    
    '���и�����Ŀ�����޸ġ�
    For i = 0 To ciptBase.ItemCount - 1
        ciptBase.ItemEnable(i) = False
    Next
    
    '�޸ģ�2002-10-10������ζ����ƣ���ʾ����
    On Error Resume Next
    If sffunc�жϼ��ϼ�ֵ�Ƿ����(mobj���.����.������Ϣ, "�����") Then
        ciptBase.Box1("�����").Text = lobj���ģ��.�շѱ�׼���
        mobj���.����.Sub�����Ϣֵ "�����", lobj���ģ��.�շѱ�׼���
    End If
    
    Set mcol�շ���Ŀ = New Collection
    clstItem.Refresh
'    DoEvents
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        sfsub������ "�����沿��", "FrmRegisterAgain", "SubGetPersonInfo", Err.Number, Err.Description, False
    End If
    Set lobj���ģ�� = Nothing
    Set lobj��� = Nothing
    
    '�ָ�����ɲ�����
    ctblTool.Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
