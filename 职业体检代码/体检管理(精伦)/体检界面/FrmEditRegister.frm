VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmEditRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�޸����Ǽ���Ϣ"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10335
   ClipControls    =   0   'False
   Icon            =   "FrmEditRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin ¼��ؼ�.ctlInputDictGrid c�ֵ�� 
      Height          =   3135
      Left            =   2640
      TabIndex        =   27
      Top             =   4200
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
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
      Interval        =   2
      Left            =   2280
      Top             =   600
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "��������"
      Height          =   270
      Left            =   8040
      TabIndex        =   12
      Top             =   240
      Width           =   1410
   End
   Begin VB.Frame frmPhoto 
      Caption         =   "����"
      ForeColor       =   &H00800000&
      Height          =   4305
      Left            =   5400
      TabIndex        =   25
      Top             =   2760
      Width           =   4695
      Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
         Height          =   3570
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4485
         _ExtentX        =   8017
         _ExtentY        =   6297
         BackColor       =   0
         FontSize        =   9.75
         OriginalSize    =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7470
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18177
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin VB.Frame cframJBXX 
      Caption         =   "�Ǽǻ�����Ϣ"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1725
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   10215
      Begin VB.TextBox ctxtTubeNo 
         Height          =   315
         Left            =   5640
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox ctxt��쵥�� 
         Height          =   315
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox ccmb������� 
         Height          =   300
         ItemData        =   "FrmEditRegister.frx":0442
         Left            =   8640
         List            =   "FrmEditRegister.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox ccmbTemplate 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         ItemData        =   "FrmEditRegister.frx":045C
         Left            =   3360
         List            =   "FrmEditRegister.frx":045E
         TabIndex        =   7
         Top             =   1200
         Width           =   3525
      End
      Begin VB.TextBox ctxtSysNo 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2265
      End
      Begin VB.TextBox ctxtName 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1200
      End
      Begin VB.ComboBox ccmbSex 
         Height          =   300
         ItemData        =   "FrmEditRegister.frx":0460
         Left            =   1560
         List            =   "FrmEditRegister.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   960
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "��λ(&T)"
         Height          =   375
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox ctxtAge 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1200
         Width           =   495
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
         TabIndex        =   30
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   132775936
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥�ţ�"
         Height          =   180
         Index           =   7
         Left            =   6840
         TabIndex        =   29
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ͣ�"
         Height          =   180
         Index           =   8
         Left            =   8640
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   4
         Left            =   1560
         TabIndex        =   23
         Top             =   960
         Width           =   540
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6120
         TabIndex        =   22
         Top             =   480
         Width           =   675
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5880
         TabIndex        =   21
         Top             =   480
         Width           =   375
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
         Caption         =   "������ڣ�"
         Height          =   180
         Index           =   2
         Left            =   8400
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���ƣ�"
         Height          =   180
         Index           =   5
         Left            =   3360
         TabIndex        =   17
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   2640
         TabIndex        =   16
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ¼��ؼ�.ctlInputFrame ciptBase 
      Height          =   4455
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   5025
      _ExtentX        =   8864
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
      Caption         =   "Frame1"
      Rows            =   6
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
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   2760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmEditRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��˺�
'����޸ģ��

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

Private mobj��� As Object                   '������

Private mblnTakePhoto  As Boolean            'ҵ�����á��Ƿ����ࡱ��
Private mbln����¼�� As Boolean

Private mstrϵͳ��� As String               '��ʼҪ�޸ĵ�����¼��ϵͳ��ţ�����������������

Private mstrϵͳ��Ź̶����� As String

Private mstr��λ������ As String
Private mblnInUse As Boolean                 '��Ӧ����pblnInUse��

Public pstrϵͳ������� As String '�޸ģ�2002-10-10�����Ϊ�ζ��������Ӹ����ԡ�

Private mcol�շ���Ŀ As Collection
Private mcol�����Ŀ As Collection

'���ܣ���ȡ�������Ƿ��������ı�־��
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

'���ܣ����ô���������������ǰҪ�޸ĵ�����¼��ϵͳ��š�
Public Property Let ϵͳ���(ByVal vNewValue As String)
    mstrϵͳ��� = vNewValue
    
End Property


Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
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
            ciptBase.SetFocus
        End If
    End If
End Sub

Private Sub ccmbUnit_Click()
    Dim i As Integer
    
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
        '�޸ģ�2001-8-23����ʾ��λ���ԣ���
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

Private Sub ccmb�������_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub ciptBase_LastLostFocus()
    Dim blnCancel As Boolean
    On Error Resume Next
    '�Զ����档
    If ctbMain.Buttons(1).Enabled Then
'        ctxtSysNo.SetFocus
'        SendKeys "{F2}"
        mobjGUI_BeforeOperate "����", blnCancel
    End If
End Sub


Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub
Private Sub ctxtAge_GotFocus()
    On Error Resume Next
    With ctxtAge
        .SelStart = 0
        .SelLength = Len(Trim(ctxtAge.Text))
    End With
End Sub
Private Sub ctxtAge_KeyPress(KeyAscii As Integer)
'    On Error GoTo errhandler
'    '�ж��Ƿ�Ϊ����
'    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = 13 Then
'    Else
'        sffuncMsg "����ֻ��Ϊ���֣�������¼�롣", sf����
'        KeyAscii = 0
'    End If
'    Exit Sub
'errhandler:
End Sub

Private Sub ctxtAge_LostFocus()
'    On Error GoTo errhandler
'    '�ƶ����㡣
'    If ctxtAge <> "" Then
'        If Val(ctxtAge) > 150 Or Val(ctxtAge) <= 0 Then
'            sffuncMsg "�������>0�����Ҳ����ܳ���150���������������䡣", sf����
'            ctxtAge.SetFocus
'        End If
'    End If
'    Exit Sub
'errhandler:
End Sub


Private Sub ctxtSysNo_Change()
    On Error Resume Next
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
End Sub

Private Sub ctxtTubeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt��쵥��.SetFocus
    End If
End Sub

Private Sub ctxt��쵥��_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
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
    Dim lcol��������ť As New Collection           '�������ϵİ�ť��ʼ�����ϡ�
   
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    Set mcol�����Ŀ = New Collection
    Set mcol�շ���Ŀ = New Collection
    '���治�ɲ�����
    cframJBXX.Enabled = False
    ciptBase.Enabled = False
    frmPhoto.Enabled = False
    ctbMain.Enabled = False
    
    '��ʼ������ͨ�ö���
    Set mobjGUI = New cls����ͨ�ö���
    mobjGUI.pbln�Զ������ֵ�߶� = False
    With lcol��������ť
        .Add "����"
        .Add "|"
        .Add "�����Ŀ(&T)102"
        .Add "|"
        .Add "�����Ƭ(&A)111"
        .Add "������Ƭ(&E)103"
        .Add "|"
        .Add "��ӡ"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
        Set .c״̬�� = csbMain
        Set .c¼��� = ciptBase
        Set .c�ֵ�� = c�ֵ��
    End With
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""
                
    '��ʼʱ�����水ť�����ã���������Ϸ���ϵͳ��ź�ſ��ã���
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
    
    If pobjҵ�����.ҵ������("�Ƿ�����") = "��" Then
        mblnTakePhoto = True
    Else
        mblnTakePhoto = False
    End If
    
    '��Ҫ����ʱ��ʼ������ؼ�
    If mblnTakePhoto Then
        '��ʼ���ؼ�
        cctlCatchPhoto.funcInitVideo
        '�жϳ�ʼ���Ƿ�ɹ�
        If cctlCatchPhoto.VideoIsOk = False Then
            sffuncMsg "�����豸��ʼ��ʧ�ܣ�����ԭ���ҵ�����������ò��������࣡", sf����
        End If
    Else
        cctlCatchPhoto.Enabled = False
    End If
    
    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
        ctxtTubeNo.Visible = True
        ctxtTubeNo.TabIndex = 1
        clblTubeNo.Visible = False
        clblLetter.Visible = False
    Else
        ctxtTubeNo.Visible = False
        clblTubeNo.Visible = True
        clblLetter.Visible = True
    End If
    
    '������ʼ���������ڶ�ʱ����ɡ�
    Timer1.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmEditRegister", "Form_Load", 6666, lstrError, False
    
    ctbMain.Enabled = True
End Sub


'���ܣ����Form_Load���µĴ����ʼ��������
Private Sub Timer1_Timer()
    Dim lobj����ģ�弯 As Object
    Dim lcol��λ���Ƽ� As New Collection
    Dim lcol����ģ�弯 As New Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    Timer1.Enabled = False
    
    Set lobj����ģ�弯 = CreateObject("������.ClsMedicalExamTemplateSet")
    lobj����ģ�弯.�������� = 3
    
    '�����е�����ģ����뵽ccmb������
    Set lcol����ģ�弯 = lobj����ģ�弯.Ԫ�ؼ�
    For i = 1 To lcol����ģ�弯.Count
        ccmbTemplate.AddItem lcol����ģ�弯(i)
    Next i
    Set lcol����ģ�弯 = Nothing
    Set lobj����ģ�弯 = Nothing
    
    If ccmbTemplate.ListCount = 0 Then
        sffuncMsg "������������ʱ�޷����У����Ƚ��롰�������á��������棬���ø�����������ݣ�", sf����
    End If
    
    '���뵥λ���������б��
    Set lcol��λ���Ƽ� = pobjҵ�����.���չ������䲾.��λ���Ƽ�
    For i = 1 To lcol��λ���Ƽ�.Count
        ccmbUnit.AddItem lcol��λ���Ƽ�(i)
    Next i
    Set lcol��λ���Ƽ� = Nothing
        
    '��������������Ҫ��ȫ�ֶ���
    Set mobj��� = CreateObject("������.clsMedicalExam")
    '�޸ģ�2002-10-10������ϵͳ������ƣ���
    If pstrϵͳ������� <> "" Then
        mobj���.ϵͳ������� = pstrϵͳ�������
    End If
    mobj���.ϵͳ��� = mstrϵͳ���
    
    mstrϵͳ��Ź̶����� = mobj���.ϵͳ��Ź̶�����
    
    If mstrϵͳ��� <> "" Then
        ctxtSysNo = mstrϵͳ���
        
        '��ʾ��ʼϵͳ��ŵ����ݡ�
        If mobj���.�Ƿ��Ѵ��� Then
            subShowRegisterInfo
            
            '���水ť���á�
            ctbMain.Buttons(1).Enabled = True
            ctbMain.Buttons(5).Enabled = True
            ctbMain.Buttons(6).Enabled = True
            ctbMain.Buttons(8).Enabled = True
            
        End If
    Else
        ctxtSysNo = mstrϵͳ��Ź̶�����
    End If
    
    If pobjҵ�����.ҵ������("�Ƿ���ٵǼ�") = "��" Then
        mbln����¼�� = True
    Else
        mbln����¼�� = False
    End If
    
    Err.Clear
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmEditRegister", "Timer1_Timer", 6666, lstrError, False
    End If
    
    '����ɲ�����
    cframJBXX.Enabled = True
    ciptBase.Enabled = True
    frmPhoto.Enabled = True
    ctbMain.Enabled = True
    If ctbMain.Buttons(1).Enabled Then
        ctxtName.SetFocus
    Else
        ctxtSysNo.SetFocus
        ctxtSysNo.SelStart = Len(ctxtSysNo)
        ctxtSysNo.SelLength = 0
    End If
    
    Exit Sub
    
    Resume
End Sub

Private Sub ccmbTemplate_Click()
    Dim i As Long
    On Error GoTo errHandler
    
    If mobj���.����.������ = ccmbTemplate.Text Then Exit Sub
    
    '�������ø�����Ϣ¼��塣
    On Error Resume Next
    mobjGUI.sub��ʼ��¼��� ccmbTemplate.Text
        
    '�޸ģ�2001-8-23����ʾ��λ���ԣ���
    If mstr��λ������ <> "" Then
        sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
    End If
        
    On Error GoTo errHandler
    mobj���.����.������ = ccmbTemplate.Text
    
    
    '�޸ģ�2002-7-26��������ݡ��Ƿ�����ѡ���������͡�
    On Error Resume Next
    Dim lobj����ģ�� As Object
    Set lobj����ģ�� = CreateObject("������.clsMedicalExamTemplate")
    lobj����ģ��.������ = ccmbTemplate.Text
    If lobj����ģ��.�Ƿ����� Then
        ccmb�������.ListIndex = 1
    Else
        ccmb�������.ListIndex = 0
    End If
    
    '�޸ģ�2002-10-10������ζ����ƣ���ʾ����
    ciptBase.Box1("�����").Text = lobj����ģ��.�շѱ�׼���
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmEditRegister", "ccmbTemplate_Click", 6666, lstrError, False
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
    '�����б��
    gfsubShowComboList ccmbUnit
    
    Exit Sub
errHandler:
End Sub

Private Sub ccmbUnit_LostFocus()
    Dim i As Integer
    
    On Error GoTo errHandler
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
    On Error GoTo errHandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��
    
    On Error GoTo errHandler
    
    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
            mstr��λ������ = lobjRec!������
        
            '��ʾ�������Ա���ڵ�λ���������ԡ�
            '�޸ģ�2001-8-23�������������
            On Error Resume Next
            sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
        End If
    End If
    
    '�ѽ���ص���λ¼���
    ccmbUnit.SetFocus
    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmEditRegister", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub



Private Sub ctxtName_GotFocus()
    On Error Resume Next
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
End Sub

Private Sub ctxtSysNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '�ƶ����㡣
        ctxtName.SetFocus
    End If
    
End Sub
'���ܣ�¼��ϵͳ��ź󣬻�ȡ����ʾ������¼�ĵǼ���Ϣ��
Private Sub ctxtSysNo_LostFocus()
    On Error GoTo errHandler
    
    '�����δ�仯����������
    If mobj��� Is Nothing Then Exit Sub
    
    If mobj���.ϵͳ��� = ctxtSysNo.Text Then Exit Sub
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
    
    '��ʾ�����Ϣ��
    mobj���.ϵͳ��� = ctxtSysNo.Text
    subShowRegisterInfo
    
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    ctbMain.Buttons(6).Enabled = True
    ctbMain.Buttons(8).Enabled = True
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmEditRegister", "ctxtSysNo_LostFocus", 6666, lstrError, False
    
    ctxtSysNo.SetFocus
    Exit Sub
    Resume
End Sub

Private Sub ctxtSysNo_GotFocus()
    On Error Resume Next
    With ctxtSysNo
        '��ʾϵͳ��ŵĹ̶����֡�
        If ctxtSysNo.Text = "" Then
            ctxtSysNo.Text = mstrϵͳ��Ź̶�����
        End If
        .SelStart = Len(Trim(ctxtSysNo.Text))
        .SelLength = 0
    End With

End Sub
Private Sub subClear()
    On Error Resume Next
    ctxtName.Text = ""
    ctxtAge.Text = ""
    ccmbUnit.Text = ""
    
    '�޸ģ�������ա�
    Dim ldbl����� As Double
    ldbl����� = ciptBase.Box1("�����").Text
    ciptBase.ClearContent
    ciptBase.Box1("�����").Text = ldbl�����

End Sub

Private Sub ciptBase_LostFocus()
    On Error Resume Next
    If ActiveControl.Name <> "c�ֵ��" Then
        c�ֵ��.Visible = False
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '�ͷŶ���
    Set mobj��� = Nothing
    Set mobjGUI = Nothing
    
    '���ô���û�м��ر�־��
    mblnInUse = False
    
    pstrϵͳ������� = ""
End Sub

'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lstrϵͳ��� As String
    Dim i As Integer
    Dim lcolԭ�����Ŀ As Collection
    Dim lstrFile As String
    
    On Error GoTo errHandler
    Select Case Operate
    Case "����"
         '�ж��Ƿ���Ҫ���ࡣ
        If mblnTakePhoto = True Then
            '�ж��Ƿ�����
            If cctlCatchPhoto.Photo Is Nothing Then
                sffuncMsg "û�����࣬����������󱣴棡", sf����
                Cancel = True
                Exit Sub
            End If
        End If
        
        '�����ǿ���¼�룬���¼���Ƿ��д���
        If mobj���.����.������Ϣ.Count > 0 Then
            '�޸ģ�2001-9-12��������¼��Ƿ��ֵ����ݣ�������ʱϵͳ������ʾ���󣩡�
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
        
        MousePointer = 11
        csbMain.Panels(1) = "���ڱ������Ǽ���Ϣ�����Ժ�..."
        
        '�������������ԡ�
        With mobj���
            .�����Ա.���� = Trim(ctxtName.Text)
            .�����Ա.�Ա� = Trim(ccmbSex.Text)
            .�����Ա.��λ���� = Trim(ccmbUnit.Text)
            
            If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
                .�Թܱ�� = ctxtTubeNo.Text
            End If
            .��쵥�� = ctxt��쵥��
            
            If mblnTakePhoto Then
                .�����Ա.��Ƭ = cctlCatchPhoto.Photo
'                .�����Ա.��Ƭѹ�� = cctlCatchPhoto.Photo
            End If
            If Val(ctxtAge) > 0 Then
                .�����Ա.�������� = DateAdd("yyyy", -Val(ctxtAge), Date)
            End If
            .�����Ա.���� = ctxtAge
            
            On Error Resume Next
            .�����Ա.������ݺ��� = ciptBase.Box1("���֤��").Text
            .�����Ա.�������� = ciptBase.Box1("��������").TrueText
            .�����Ա.Ƭ�� = ciptBase.Box1("Ƭ��").TrueText
            .�����Ա.��ҵ��� = ciptBase.Box1("��ҵ���").TrueText
            
            If .�����Ա.��λ������ <> mstr��λ������ Then
                .�����Ա.��λ������ = mstr��λ������
            End If
            .������� = Format(cdtpDate.Value, "yyyy-mm-dd")
            
            On Error GoTo errHandler
            
            '���渽����Ϣ
            For i = 1 To ciptBase.ItemCount
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
            
        End With
        
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
        
        'ִ�б��档
        If mcol�շ���Ŀ.Count = 0 Then
            pobjҵ�����.Sub���Ǽ� mobj���
        Else
            pobjҵ�����.Sub���Ǽ� mobj���, , , mcol�շ���Ŀ
        End If
        If cchkClear = 1 Then
            subClear
            ctbMain.Buttons(1).Enabled = False
            ctbMain.Buttons(5).Enabled = False
            ctbMain.Buttons(6).Enabled = False
            ctbMain.Buttons(8).Enabled = False
        End If
        
        '�ָ����ࡣ
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "�ָ�" Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
        
        '����ص�ϵͳ��š�
        ctxtSysNo.SetFocus
        ctxtSysNo.SelStart = Len(ctxtSysNo.Text)
        ctxtSysNo.SelLength = 0
        MousePointer = 0
        csbMain.Panels(1) = ""
        
        '���Խ���ͨ�ö���Ա������Ժ�Ĵ���
        Cancel = True
    
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
                mblnTakePhoto = True
            End If
        End If
    Case "��ӡ"
        Dim lcol��� As Collection
        Set lcol��� = New Collection
        '��ӡ����
        lcol���.Add ctxtSysNo.Text
        pobjҵ�����.Sub��ӡ���� "����", lcol���, True
    
    End Select
    
    Exit Sub
    
errHandler:
    '������
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterFirst", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    '�ָ����档
    MousePointer = 0
    csbMain.Panels(1) = ""
    ctxtSysNo.SetFocus
    Cancel = True
    Exit Sub
    Resume
End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal ���� As String, ByVal ���� As String, ByVal �������� As String, ByVal IsError As Boolean)
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    
    On Error GoTo errHandler
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
errHandler:
    '������
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmRegisterFirst", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemSetFocus Index
    Exit Sub
    Resume


End Sub

'���ܣ���ʾ������mobj��족�������ڽ����ϡ�
Private Sub subShowRegisterInfo()
    Dim lcolInfo As New Collection '��츽����Ŀ���ϡ�
    Dim lstr�Թܱ�� As String
    Dim lstrItem As String
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errHandler
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctbMain.Buttons(6).Enabled = False
    ctbMain.Buttons(8).Enabled = False
    
    If Not mobj���.�Ƿ��Ѵ��� Then
        '����¼������ʱ������ʾ
        ctxtSysNo.SelStart = Len(ctxtSysNo)
        ctxtSysNo.SelLength = 0
        subClear
        Err.Raise 6666, , "������Ų����ڣ���������������š�"
    End If
    '�ж��Ƿ����������ۡ�
    If mobj���.���״̬ = P_ENDED_STATUS Then
        Err.Raise 6666, , "���������ۣ��������޸����Ǽ���Ϣ��"
    End If
    '�ж��Ƿ񸴲�����¼��
    If mobj���.������ = P_EXAM_AGAIN Then
        Err.Raise 6666, , "�������޸ĸ����¼��"
    End If
    '���ҵ���������д�Ǽ���Ϣ
    With mobj���
    
        ctxtName.Text = .�����Ա.����
        If .�����Ա.�Ա� = "" Then
            ccmbSex.ListIndex = 0
        Else
            ccmbSex.ListIndex = gffuncItemIsInComboBox(ccmbSex, .�����Ա.�Ա�)
        End If
        If IsDate(.�����Ա.��������) Then
            ctxtAge.Text = DateDiff("yyyy", .�����Ա.��������, Date)
        Else
            ctxtAge.Text = ""
        End If
        ccmbUnit.Text = .�����Ա.��λ����
        If ccmbTemplate.Text <> .����.������ Then
            ccmbTemplate.Text = .����.������
            '���³�ʼ�����ӱ��
            On Error Resume Next
            mobjGUI.sub��ʼ��¼��� ccmbTemplate.Text
            On Error GoTo errHandler
        End If
        lstr�Թܱ�� = .�Թܱ��
        If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
            clblLetter.Caption = Left(lstr�Թܱ��, InStr(1, lstr�Թܱ��, ":") - 1)
            clblTubeNo.Caption = Right(lstr�Թܱ��, Len(lstr�Թܱ��) - InStr(1, lstr�Թܱ��, ":"))
        Else
            ctxtTubeNo = lstr�Թܱ��
        End If
        If IsDate(.�������) Then
            cdtpDate.Value = Format(.�������, "yyyy-mm-dd")
        Else
            cdtpDate.Value = Date
        End If
        '�޸ģ�2001-12-29����ʾ������ͣ���
        If .������ = P_EXAM_ANNUAL Then
            ccmb�������.ListIndex = 1
        Else
            ccmb�������.ListIndex = 0
        End If
        
        '�޸ģ�2001-8-23��
        mstr��λ������ = .�����Ա.��λ������
        
        ctxt��쵥�� = .��쵥��
    End With
    
            
    '��ȡ���и�����Ŀ��������
    Set lcolInfo = mobj���.����.������Ϣ
    
    '��д������Ϣ�����
    sub��¼���ֵ ciptBase, mobjGUI, lcolInfo
    
    '��ò���ʾ��Ƭ��
    If Not mobj���.�����Ա.��Ƭ Is Nothing Then
        Set cctlCatchPhoto.Photo = mobj���.�����Ա.��Ƭ
    Else
        cctlCatchPhoto.subClear
    End If
    
    '����δ��ʼ���ʱ���޸�����
    If mobj���.���״̬ = P_LOGIN_STATUS Then
        ccmbTemplate.Enabled = True
    Else
        ccmbTemplate.Enabled = False
    End If
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    ctbMain.Buttons(6).Enabled = True
    ctbMain.Buttons(8).Enabled = True
    Exit Sub
errHandler:
    sfsub������ "�����沿��", "FrmEditRegister", "subShowRegisterInfo", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
