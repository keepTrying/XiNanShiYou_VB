VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGiveConclusion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11445
   ClipControls    =   0   'False
   Icon            =   "frmGiveConclusion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox cchkȫѡ 
      Caption         =   "ȫѡ"
      Height          =   255
      Left            =   2520
      TabIndex        =   32
      Top             =   720
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.OptionButton coptType 
      Caption         =   "δ�½���"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   31
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "���½���"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   30
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   60
      TabIndex        =   21
      Top             =   7200
      Width           =   10335
      Begin VB.TextBox ctxtDoctor 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   345
         Left            =   3960
         TabIndex        =   27
         Top             =   240
         Width           =   1140
      End
      Begin VB.Frame cframPrintPaper 
         Appearance      =   0  'Flat
         Caption         =   "������飺"
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   5160
         TabIndex        =   22
         Top             =   120
         Width           =   5055
         Begin VB.OptionButton coptPaper 
            BackColor       =   &H00F0F3E9&
            Caption         =   "��ӡ�����֪ͨ��"
            Height          =   240
            Index           =   0
            Left            =   1920
            TabIndex        =   25
            Top             =   240
            Width           =   1980
         End
         Begin VB.OptionButton coptPaper 
            BackColor       =   &H00F0F3E9&
            Caption         =   "��ӡ�������"
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1650
         End
         Begin VB.OptionButton coptPaper 
            BackColor       =   &H00F0F3E9&
            Caption         =   "����ӡ"
            Height          =   240
            Index           =   2
            Left            =   4080
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   930
         End
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
         Height          =   345
         Left            =   1080
         TabIndex        =   26
         Top             =   240
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
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
         Format          =   69074944
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�½���ʱ�䣺"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�½���ҽʦ��"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "�г������Ա"
      ForeColor       =   &H80000008&
      Height          =   3435
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   2355
      Begin VB.OptionButton coptBatch 
         Caption         =   "ϵͳ���(�����)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1485
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton coptBatch 
         Caption         =   "����δ�½���"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton coptBatch 
         Caption         =   "�������"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox ctxtNo 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   1800
         Width           =   2145
      End
      Begin VB.ComboBox ccmbSheet 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2220
      End
      Begin VB.CommandButton ccmdAddPerson 
         Caption         =   "���(&Q)"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker cdtpStart 
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
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   69074944
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ѡ������"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "�����Ŀ�����"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2925
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   7305
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdResult 
         Height          =   2505
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4419
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
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
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   15791081
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "^������Ŀ       |^���           |^������Ŀ          |^���           "
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   1111
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkˢ���� 
         Caption         =   "ˢ����"
         Height          =   375
         Left            =   9600
         TabIndex        =   33
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame cfram���� 
      Appearance      =   0  'Flat
      Caption         =   "����(���밴���޸ġ���ť�ſ����޸ģ�"
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   7320
      TabIndex        =   10
      Top             =   4320
      Width           =   3975
      Begin VB.CommandButton ccmdUpdateConclusion 
         Appearance      =   0  'Flat
         Caption         =   "�޸�(&M)"
         Enabled         =   0   'False
         Height          =   465
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin RichTextLib.RichTextBox ctxtDiagnosis 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1085
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmGiveConclusion.frx":0442
      End
      Begin RichTextLib.RichTextBox ctxtConclusion 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   873
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmGiveConclusion.frx":04DF
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
      Begin RichTextLib.RichTextBox ctxtTemplate 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmGiveConclusion.frx":057C
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϻʹ��������"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ۣ�"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   4200
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdPerson 
      Height          =   3225
      Left            =   2520
      TabIndex        =   19
      Top             =   960
      Width           =   8745
      _cx             =   87571521
      _cy             =   87561785
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
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
      MousePointer    =   0
      BackColor       =   8454016
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   8454016
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmGiveConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mcolFieldIndex As Collection 'cgrdPerson�����и���������Ӧ���кš�

Private mstrϵͳ��Ź̶����� As String
Private mbln����ȡ������ As Boolean  '��ʾ��ǰ�û��Ƿ����ȡ�����۵�Ȩ�ޡ�
Private mblnInUse As Boolean

Private mobj���� As cls�û��������� '�޸ģ�2001-12-29�����Ӹö��󣩡�

Private mblnSys As Boolean

'���ܣ�������ǰ�����Ƿ��Ѽ��أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��
Public Property Get pblnInUse() As Boolean
    On Error GoTo errHandler
    pblnInUse = mblnInUse
    Exit Property
errHandler:
    sfsub������ "�����沿��", "frmGiveConclusion", "Property Get pblnInUse", Err.Number, Err.Description, True
End Property



Private Sub cchkȫѡ_Click()
    Dim i As Long
    For i = 1 To cgrdPerson.Rows - 1
        cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("�����֤"), i, mcolFieldIndex("�����֤")) = IIf(cchkȫѡ.Value = 1, flexChecked, flexUnchecked)
    Next
End Sub

Private Sub cchkˢ����_Click()
    On Error Resume Next
    If coptBatch(2).Value Then
        ctxtNo.SetFocus
    End If
End Sub

'�޸ģ�2002-1-16��������ʾ�����Ϣ����
Private Sub cgrdPerson_AfterSort(ByVal Col As Long, Order As Integer)
    On Error Resume Next
    If cgrdPerson.Row > 0 Then
        cgrdPerson_Click
    End If
End Sub

Private Sub cgrdPerson_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If coptType(0) Then
        If Col <> mcolFieldIndex("�����֤") Then Cancel = True
    Else
        Cancel = True
    End If
End Sub

Private Sub coptBatch_Click(Index As Integer)
    On Error Resume Next
    If coptBatch(Index).Value Then
        If Index = 0 Then
            cdtpStart.SetFocus
        ElseIf Index = 1 Then
'            ctxt��λ����.SetFocus
        Else
            ctxtNo.SetFocus
        End If
    End If
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error Resume Next
    subReset
    If coptType(0).Value Then
        '�½���
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
        ctbMain.Buttons(4).Enabled = True
        ctbMain.Buttons(5).Enabled = False
        Frame1(0).ForeColor = &HFF0000    '��ɫ��
        Frame1(0).Enabled = True
        cdtpDate.Enabled = True
        coptBatch(1).Visible = True
    Else
        'ȡ������
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
        ctbMain.Buttons(4).Enabled = False
        ctbMain.Buttons(5).Enabled = True
        Frame1(0).ForeColor = &HFF00FF    '��ɫ��
        Frame1(0).Enabled = True
        cdtpDate.Enabled = False
        coptBatch(1).Visible = False
        If coptBatch(1).Value Then
            coptBatch(0).Value = True
        End If
    End If
    If coptBatch(0) Then
        cdtpStart.SetFocus
    ElseIf coptBatch(2) Then
        ctxtNo.SetFocus
    End If
End Sub

Private Sub ctxtNo_GotFocus()
    On Error Resume Next
    If cchkˢ����.Value = 1 Then
        ctxtNo = ""
    ElseIf ctxtNo.Text = "" Then
        ctxtNo.Text = mstrϵͳ��Ź̶�����
        ctxtNo.SelLength = 0
        ctxtNo.SelStart = Len(mstrϵͳ��Ź̶�����)
    End If
End Sub

Private Sub ctxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And ctxtNo <> "" Then
        ccmdAddPerson_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    '�޸ģ�2002-7-1�������ȡ�����۵Ĳ�������Ϊ������ѡ��
    With lcol��������ť
        .Add "ȥ����Ա(&D)129"
        .Add "�����Ա(&R)106"
        .Add "|"
        .Add "�������(&O)109"
        .Add "ȡ������(&K)104"
        .Add "|"
        .Add "����(&O)111"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
'        Set .c״̬�� = csbMain
    End With
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""
    
    '��ʾ"�½���ҽʦ"Ϊ��ǰ�û�����
    ctxtDoctor = um�û���
    cdtpDate.Value = Date
    cdtpStart.Value = Date
    
    '��ȡϵͳ��Ź̶����֡�
    Dim lobj��� As Object '�����󣬻�ȡϵͳ��ŵĹ̶����֡�
    Set lobj��� = CreateObject("������.clsMedicalExam")
    mstrϵͳ��Ź̶����� = lobj���.ϵͳ��Ź̶�����
    Set lobj��� = Nothing
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    '�ж��Ƿ���ȡ�����۵�Ȩ�ޡ�
    mbln����ȡ������ = umfuncУ���û�Ȩ��("ȡ��������")
    If Not mbln����ȡ������ Then
        mbln����ȡ������ = umfuncУ���û�Ȩ��("������_ȡ��������")
    End If
    
    Dim lobj����ģ�弯 As Object
    Dim lcolInfo As Collection
    Dim i As Long
    
    Set lobj����ģ�弯 = CreateObject("������.ClsMedicalExamTemplateSet")
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    
    ccmbSheet.Clear
    '�޸ģ�2002-8-14��������ӡ�<����>��ѡ���
    If lcolInfo.Count > 0 Then
        ccmbSheet.AddItem "<����>"
    End If
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next i
    If ccmbSheet.ListCount > 1 Then
        ccmbSheet.ListIndex = 1
    End If

    '�޸ģ�2001-12-29����ȡ��������ֵ����
    On Error Resume Next
    Set mobj���� = New cls�û���������
    mobj����.�û���� = um�û����
    mobj����.ҵ���� = "������"
    
    If mobj����.������ֵ("�½���ʱˢ����") = "��" Then
        cchkˢ����.Value = 1
    Else
        cchkˢ����.Value = 0
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmGiveConclusion", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If coptBatch(0) Then
        cdtpStart.SetFocus
    ElseIf coptBatch(1) Then
'        ctxt��λ����.SetFocus
    Else
        ctxtNo.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub

Private Sub cdtpStart_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdAddPerson.SetFocus
    End If
End Sub



Private Sub ccmdAddPerson_Click()
    Dim lobjRec As Object    '��ȡ�Ŀ����½��۵�����¼��
    Dim llngMaxRow As Long
    Dim i As Long, j As Long
    
    On Error GoTo errHandler
    
'    If coptBatch(1).Value And ctxt��λ���� = "" Then
'        MsgBox "�����뵥λ���ƣ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
'        ctxt��λ����.SetFocus
'        Exit Sub
'    End If
    If coptBatch(2).Value And ctxtNo = "" Then
        MsgBox "������ϵͳ��ţ���ˢ�������ϵ����룡", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        ctxtNo.SetFocus
        Exit Sub
    End If
    
    '��ȡ�����ϵͳ��ţ���Χ�������ڵĿ����������۵�����¼��
    If ctbMain.Buttons(5).Enabled Then
        'Ҫȡ�����ۣ���ȡ���½��۵�����¼��
        Set lobjRec = pobjҵ�����.Func��ȡ��������ȷ��������¼(IIf(coptBatch(0).Value, Format(cdtpStart.Value, "yyyy-mm-dd"), ""), "", IIf(ccmbSheet.ListIndex > 0, ccmbSheet.Text, ""), IIf(coptBatch(2).Value, ctxtNo.Text, ""))
    Else
        'Ҫ������ۣ���ȡ��δ�½��۵�����¼��
        If Not coptBatch(1).Value Then
            Set lobjRec = pobjҵ�����.Func��ȡ���½��۵�δȷ��������¼(IIf(coptBatch(0).Value, Format(cdtpStart.Value, "yyyy-mm-dd"), ""), "", IIf(ccmbSheet.ListIndex > 0, ccmbSheet.Text, ""), IIf(coptBatch(2).Value, ctxtNo.Text, ""))
        Else
            Set lobjRec = pobjҵ�����.Func��ȡ�������½��۵�δȷ��������¼()
        End If
    End If
    
    '�޸ģ�2001-11-15�����Ϊ�������ʾЧ�ʣ�ʹ��DataSource��
    cgrdPerson.Redraw = False
    
    If lobjRec.recordcount = 0 Then
        If ctbMain.Buttons(5).Enabled Then
            'ȡ�����ۡ�
            sffuncMsg "û����ָ����Χ�������ϣ��������������۵�����¼��" & Chr(13) & Chr(10) & "��ע�⣺��ֻ��ȡ�����Լ����µ������ۡ�", sf����
        Else
            '�½��ۡ�
            sffuncMsg "û����ָ����Χ�������ϣ�����δ�������۵�����¼��" & Chr(13) & Chr(10) & "����볣�桢������¼����棬��������Ҫ�½��۵������Ա������¼�����Ƿ��ѵǼ������������Ŀ���������", sf����
        End If
        If cgrdPerson.Rows = 1 Then
            Set cgrdPerson.DataSource = lobjRec
            cgrdPerson.Rows = lobjRec.recordcount + 1
        End If
    Else
        '��ʾ��ȡ��¼��cgrdPerson�С�
        If cgrdPerson.Rows = 1 Then
            Set cgrdPerson.DataSource = lobjRec
            cgrdPerson.Rows = lobjRec.recordcount + 1
        Else
            '�ų��ظ��ļ�¼��
            gfsubAppendGridFromRecWithUnique cgrdPerson, lobjRec, mcolFieldIndex("ϵͳ���")
        End If
    End If
    lobjRec.Close
    
    '��ȡcgrdPerson�и��е��кš�
    Set mcolFieldIndex = New Collection
    For i = 0 To cgrdPerson.Cols - 1
        mcolFieldIndex.Add i, cgrdPerson.TextMatrix(0, i)
    Next
    
    '���ò�������Ϊ��ɫ��
    For i = 1 To cgrdPerson.Rows - 1
        Select Case cgrdPerson.TextMatrix(i, mcolFieldIndex("������"))
        Case "����", "����"
        Case Else
            cgrdPerson.Cell(flexcpBackColor, i, 0, i, cgrdPerson.Cols - 1) = &H8A5AFA
        End Select
        If coptType(0).Value Then
            'Ĭ�������������֤��
            Select Case cgrdPerson.TextMatrix(i, mcolFieldIndex("������"))
            Case "����", "����"
                cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("�����֤"), i, mcolFieldIndex("�����֤")) = flexChecked
            Case Else
                cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("�����֤"), i, mcolFieldIndex("�����֤")) = flexUnchecked
            End Select
        Else
            If cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("�����֤"), i, mcolFieldIndex("�����֤")) <> flexChecked And cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("�����֤"), i, mcolFieldIndex("�����֤")) <> flexUnchecked Then
                If cgrdPerson.TextMatrix(i, mcolFieldIndex("�����֤")) = "1" Then
                    cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("�����֤"), i, mcolFieldIndex("�����֤")) = flexChecked
                Else
                    cgrdPerson.Cell(flexcpChecked, i, mcolFieldIndex("�����֤"), i, mcolFieldIndex("�����֤")) = flexUnchecked
                End If
            End If
        End If
        cgrdPerson.TextMatrix(i, mcolFieldIndex("�����֤")) = ""
    
    Next
    cdtpStart.SetFocus
    cgrdPerson.AutoSize 0, cgrdPerson.Cols - 1
    If cgrdPerson.Rows = 1 Then
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
    Else
        ctbMain.Buttons(2).Enabled = True
    End If
'    cgrdPerson.ColHidden(mcolFieldIndex("ϵͳ���")) = True
    cgrdPerson.ColHidden(0) = True
    cgrdPerson.Redraw = True
    cgrdPerson.Editable = True
    If coptBatch(0).Value Then
        cdtpStart.SetFocus
    ElseIf coptBatch(1).Value Then
'        ctxt��λ����.SetFocus
    Else
        ctxtNo.SetFocus
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmGiveConclusion", "ccmdAddPerson_Click", 6666, lstrError, False
    cgrdPerson.Redraw = True
    Exit Sub
    Resume
End Sub


Private Sub ccmdUpdateConclusion_Click()
    Dim llngRow As Long '�����Ա����ĵ�ǰ�кš�
    
    On Error GoTo errHandler
    llngRow = cgrdPerson.Row
    
    '����frmUpdateConclusion�����ԡ�
    With frmUpdateConclusion
        .ϵͳ��� = cgrdPerson.TextMatrix(llngRow, mcolFieldIndex("ϵͳ���"))
        .������ = ctxtConclusion.Text
        .��ϴ������ = ctxtDiagnosis.Text
        .���������� = ctxtTemplate.Text
    End With
    
    '�����޸������۴��塣
    frmUpdateConclusion.Show 1, Me
    
    '��ȡfrmUpdateConclusion��Ӧ���ԣ����޸Ľ����ϵ�ctxtConclusion��ctxtDiagnosis��
    With frmUpdateConclusion
        ctxtConclusion.Text = .������
        ctxtDiagnosis.Text = .��ϴ������
        ctxtTemplate.Text = .����������
    End With
    
    '�޸�cgrdPerson�ĵ�ǰ�С�
    With cgrdPerson
        .TextMatrix(llngRow, mcolFieldIndex("������")) = ctxtConclusion.Text
        .TextMatrix(llngRow, mcolFieldIndex("��Ϻʹ������")) = ctxtDiagnosis.Text
        .TextMatrix(llngRow, mcolFieldIndex("����������")) = ctxtTemplate.Text
        
        Select Case .TextMatrix(llngRow, mcolFieldIndex("������"))
        Case "����", "����"
            '��ɫ��
            .Cell(flexcpBackColor, llngRow, 0, llngRow, .Cols - 1) = &HC0FFC0
            ctxtConclusion.BackColor = &HC0FFC0
            .Cell(flexcpChecked, llngRow, mcolFieldIndex("�����֤"), llngRow, mcolFieldIndex("�����֤")) = flexChecked
            
        Case Else
            '��ɫ��
            .Cell(flexcpBackColor, llngRow, 0, llngRow, .Cols - 1) = &H8A5AFA
            ctxtConclusion.BackColor = &H8A5AFA
            .Cell(flexcpChecked, llngRow, mcolFieldIndex("�����֤"), llngRow, mcolFieldIndex("�����֤")) = flexUnchecked
        End Select
        
    End With
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmGiveConclusion", "ccmdUpdateConclusion_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub cgrdPerson_Click()
    Dim lobjRec As Object
    Dim lobj���� As Object   'clsMedicalExamSheet��������Ի�ȡ��ǰ�����Ա���������
    Dim lcolInfo As Collection '��ǰ�����Ա�������Ŀ������,item:clsFactTestItem��
    Dim lobjItem As Variant    'clsFactTestItem��lcolInfo�е�Ԫ�ء�
    Dim i As Long
    
    On Error GoTo errHandler
    If cgrdPerson.Row <= 0 Then Exit Sub
    
    MousePointer = 11
'    csbMain.Panels(1) = "���ڻ�ȡ��ǰ�����Ա������������Ժ�..."
    
    '��ȡ����ʾ��ǰ�����Ա�������۵��޸�����
    ctxtConclusion.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("������"))
    ctxtDiagnosis.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("��Ϻʹ������"))
    ctxtTemplate.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("����������"))
    
    '�����������
    Set lobj���� = CreateObject("������.clsMedicalExamSheet")
    lobj����.ϵͳ��� = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("ϵͳ���"))
    
    cgrdResult.Redraw = False
    
    
    '�Ż��㷨��2001-4-22��
    Set lobjRec = lobj����.�Ż������Ŀ��("����")
    If lobjRec.recordcount > 0 Then
        cgrdResult.Rows = lobjRec.recordcount + 1
    Else
        cgrdResult.Rows = 1
    End If
    i = 1
    Do While Not lobjRec.EOF
        cgrdResult.TextMatrix(i, 0) = lobjRec!�����Ŀ����
        cgrdResult.TextMatrix(i, 1) = IIf(IsNull(lobjRec!�����), "", lobjRec!�����)
        If cgrdResult.TextMatrix(i, 1) <> "" Then
            cgrdResult.TextMatrix(i, 1) = cgrdResult.TextMatrix(i, 1) & IIf(IsNull(lobjRec!��λ), "", lobjRec!��λ)
        End If
        If IIf(IsNull(lobjRec!�������), "", lobjRec!�������) = "���ϸ�" Then
            cgrdResult.Cell(flexcpBackColor, i, 1, i, 1) = &H8A5AFA
        Else
            cgrdResult.Cell(flexcpBackColor, i, 1, i, 1) = vbWhite
        End If
        i = i + 1
        lobjRec.movenext
    Loop
    
    '�Ż��㷨��2001-4-22��
    Set lobjRec = lobj����.�Ż������Ŀ��("����")
    i = 1
    Do While Not lobjRec.EOF
        If i = cgrdResult.Rows Then
            cgrdResult.Rows = cgrdResult.Rows + 1
        End If
        cgrdResult.TextMatrix(i, 2) = lobjRec!�����Ŀ����
        cgrdResult.TextMatrix(i, 3) = IIf(IsNull(lobjRec!�����), "", lobjRec!�����)
        If cgrdResult.TextMatrix(i, 3) <> "" Then
            cgrdResult.TextMatrix(i, 3) = cgrdResult.TextMatrix(i, 3) & IIf(IsNull(lobjRec!��λ), "", lobjRec!��λ)
        End If
        If IIf(IsNull(lobjRec!�������), "", lobjRec!�������) = "���ϸ�" Then
            cgrdResult.Cell(flexcpBackColor, i, 3, i, 3) = &H8A5AFA
        Else
            cgrdResult.Cell(flexcpBackColor, i, 3, i, 3) = vbWhite
        End If
        i = i + 1
        lobjRec.movenext
    Loop
    Do While i < cgrdResult.Rows
        cgrdResult.TextMatrix(i, 2) = ""
        cgrdResult.TextMatrix(i, 3) = ""
        i = i + 1
    Loop
    
    'ˢ�����������
    cgrdResult.Redraw = True
    
    '����ȡ�����ۣ���ʾ�½������ڣ����ҽʦ��
    If ctbMain.Buttons(5).Enabled Then
        If IsDate(cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("�½�������"))) Then
            cdtpDate.Value = Format(cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("�½�������")), "yyyy-mm-dd")
        Else
            cdtpDate.Value = Format(Date, "yyyy-mm-dd")
        End If
        ctxtDoctor.Text = cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("�½���ҽʦ����"))
    End If
    
    '����"ȥ����Ա"���������Ա����ť���á�
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(2).Enabled = True
    If ctbMain.Buttons(5).Enabled Then
        'ȡ�����ۣ����޸ġ���ť�����á�
        ccmdUpdateConclusion.Enabled = False
    Else
        '�½��ۣ����á��޸ġ���ť���á�
        ccmdUpdateConclusion.Enabled = True
    End If
    
    Select Case cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("������"))
    Case "����", "����"
        '��ɫ��
        ctxtConclusion.BackColor = &HC0FFC0
    Case Else
        '��ɫ��
        ctxtConclusion.BackColor = &H8A5AFA
    End Select
    
    
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmGiveConclusion", "cgrdPerson_Click", 6666, lstrError, False
    
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub





Private Sub Form_Resize()
    On Error Resume Next
    cgrdPerson.Width = Me.ScaleWidth - cgrdPerson.Left - 60
    Frame2.Width = Me.ScaleWidth - Frame2.Left - 60
    Frame2.Top = Me.ScaleHeight - Frame2.Height - 60
    Frame1(1).Top = Frame2.Top - Frame1(1).Height - 60
    cfram����.Top = Frame1(1).Top
    cfram����.Left = Me.ScaleWidth - cfram����.Width - 60
    Frame1(1).Width = cfram����.Left - Frame1(1).Left - 60
    
    cgrdResult.Width = Frame1(1).Width - cgrdResult.Left - 60
    
    cgrdPerson.Height = Frame1(1).Top - cgrdPerson.Top - 30
    Frame1(0).Height = Frame1(1).Top - Frame1(0).Top - 30
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '�޸ģ�2002-9-26����������������ֵ��
    mobj����.sub���Ǽ���ֵ "�½���ʱˢ����", IIf(cchkˢ����.Value = 1, "��", "��")
    
    '�ͷ�ģ�鼶����
    Set mobjGUI = Nothing
    
    Unload frmUpdateConclusion
    
    '���ñ�־pblnInUse��
    mblnInUse = False
End Sub

'���ܣ��ָ������ʼ̬��
Private Sub subReset()
    On Error GoTo errHandler
    On Error Resume Next
    
    '���cgrdPers��¼������ֻ��"�½���"��"ȡ������"��ť����(��"mbln����ȡ������"=false����"ȡ������"��ť������)��
    cgrdPerson.Rows = 1
    cgrdResult.Rows = 1
    ctxtConclusion.Text = ""
    ctxtDiagnosis.Text = ""
    ctxtTemplate.Text = ""
    cdtpDate.Value = Format(Date, "yyyy-mm-dd")
    ctxtDoctor.Text = um�û���
    
    ccmdUpdateConclusion.Enabled = False
    Frame1(0).Enabled = False
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmGiveConclusion", "subReset", 6666, lstrError, True
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    
    On Error GoTo errHandler
    Select Case Operate
    Case "�������"
        Dim lstr�������� As String
        Dim lbln�Ƿ�Ԥ������ As Boolean
        
        If cgrdPerson.Rows = 1 Then
            sffuncMsg "���������Ҫ�½��۵���Ա�������У�", sf����
            Exit Sub
        End If
        If cgrdPerson.Rows > 2 Then
            If Not sffuncMsg("��ȷ��Ҫ�������������������Ա����������" & Chr(13) & Chr(10) & "����ѡ���ǡ�������Щ�˵������ۻᱣ�����������Ҳ������޸������������Ҫ�޸��������ֻ����ȡ�����ۡ�", sfѯ��) Then
                Exit Sub
            End If
        End If
        MousePointer = 11
        
        '��ȡ��Ҫ��ӡ�����顣
        If coptPaper(0).Value Then
            lstr�������� = "�����֪ͨ��"
        ElseIf coptPaper(1).Value Then
            lstr�������� = "�������"
        Else
            lstr�������� = ""
        End If
        
        '���������Ա�����������˵������ۡ�
        Dim llngRow As Long
        Dim lbln�����֤ As Boolean
        i = 1
        llngRow = cgrdPerson.Rows - 1
        For i = 1 To llngRow
'            csbMain.Panels(1) = "����ȷ����" & i & "���ˣ���" & (llngRow) & "���ˣ��������ۡ� "
            '���������ۡ�
            With cgrdPerson
                If .Cell(flexcpChecked, 1, mcolFieldIndex("�����֤"), 1, mcolFieldIndex("�����֤")) = flexChecked Then
                    lbln�����֤ = True
                Else
                    lbln�����֤ = False
                End If
                pobjҵ�����.Subȷ�������� .TextMatrix(1, mcolFieldIndex("ϵͳ���")), .TextMatrix(1, mcolFieldIndex("������")), Format(cdtpDate.Value, "yyyy-mm-dd"), .TextMatrix(1, mcolFieldIndex("��Ϻʹ������")), .TextMatrix(1, mcolFieldIndex("����������")), lstr��������, lbln�Ƿ�Ԥ������, lbln�����֤
            End With
            cgrdPerson.RemoveItem 1
        Next
        
        '�������
        cgrdPerson.Rows = 1
        cgrdResult.Rows = 1
        ccmdUpdateConclusion.Enabled = False
        ctxtConclusion.Text = ""
        ctxtDiagnosis.Text = ""
        ctxtTemplate.Text = ""
'        csbMain.Panels(1) = "���������ϡ�"
    
        MousePointer = 0
        Cancel = True
    Case "ȡ������"
        If cgrdPerson.Row > 0 Then
            If Not mbln����ȡ������ Then
                MsgBox "�Բ�����û��ȡ�����۵�Ȩ�ޣ�", vbOKOnly + vbInformation, "ϵͳ��ʾ"
                Exit Sub
            End If
            'ѯ�ʡ�
            If sffuncMsg("��ȷ��Ҫȡ����ǰ�����Ա��" & cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("����")) & " ������������", sfѯ��) Then
                'ͨ��ҵ�����ȡ����ǰ�����Ա�������ۡ�
                pobjҵ�����.Subȡ�������� cgrdPerson.TextMatrix(cgrdPerson.Row, mcolFieldIndex("ϵͳ���"))
            
                '�ѵ�ǰ�����Ա�������Ա����ɾ����
                cgrdPerson.RemoveItem cgrdPerson.Row
                cgrdResult.Rows = 1
            End If
        Else
            If cgrdPerson.Rows = 1 Then
                sffuncMsg "���������Ҫȡ�����۵���Ա�������У�", sf����
                Exit Sub
            End If
        End If
        Cancel = True
    Case "ȥ����Ա"
        '��������ɾ����ǰ�С�
        If cgrdPerson.Row > 0 Then
            cgrdPerson.RemoveItem cgrdPerson.Row
        End If
        If cgrdPerson.Rows = 1 Then
            ctbMain.Buttons(1).Enabled = False
            ctbMain.Buttons(2).Enabled = False
        End If
        Cancel = True
    Case "�����Ա"
        cgrdPerson.Rows = 1
        cgrdResult.Rows = 1
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Buttons(2).Enabled = False
        ccmdUpdateConclusion.Enabled = False
        Cancel = True
    Case "����"
        Dim lstrFile As String
        ccmdFile.Filter = "Excel�ļ� (*.xls)|*.xls|�ı��ļ� (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            cgrdPerson.SaveGrid lstrFile, flexFileTabText, True
        End If
    
    End Select
    
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmGiveConclusion", "ctbMain_ButtonClick", 6666, lstrError, False
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
