VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmֱ���շ� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ֱ���շ�"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   12030
   ClipControls    =   0   'False
   Icon            =   "frmֱ���շ�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Ʊ�ݺ�"
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   120
      TabIndex        =   52
      Top             =   8040
      Width           =   3615
      Begin VB.TextBox ctxtƱ�ݺ� 
         Height          =   375
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   53
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label clblCurNoArea 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   58
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ�ŶΣ�"
         Height          =   180
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��ǰƱ�ţ�"
         Height          =   180
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   120
      TabIndex        =   35
      Top             =   720
      Width           =   11685
      Begin VB.ComboBox ccmb������ 
         Height          =   300
         Left            =   5520
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox ccmb��Ӧҵ�� 
         Height          =   300
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton ccmd��λ 
         Caption         =   "..."
         Height          =   375
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox ccmbƬ�� 
         Height          =   300
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ccmb�������� 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox ccmb���ܿ��� 
         Height          =   300
         Left            =   9480
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   3
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   2
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   1320
      End
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H00F0F0F0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   4800
         TabIndex        =   54
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   9000
         TabIndex        =   51
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ƭ��"
         Height          =   180
         Left            =   2760
         TabIndex        =   41
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ܿ���"
         Height          =   180
         Index           =   4
         Left            =   8640
         TabIndex        =   39
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ѵ�λ"
         Height          =   180
         Index           =   3
         Left            =   4680
         TabIndex        =   38
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   2
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շѱ��"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   36
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   9840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      Caption         =   "���ü���"
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   3840
      TabIndex        =   18
      Top             =   8040
      Width           =   8010
      Begin VB.ComboBox cmb���ѷ�ʽ 
         Height          =   300
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   810
         Width           =   1365
      End
      Begin VB.TextBox ctxtInput 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-M-d"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   24
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   23
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   810
         Width           =   1455
      End
      Begin VB.TextBox ctxtInput 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   6360
         MaxLength       =   12
         TabIndex        =   11
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox ctxtInput 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   21
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   300
         Width           =   1365
      End
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   20
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   25
         Left            =   150
         TabIndex        =   30
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ѷ�ʽ"
         Height          =   180
         Index           =   24
         Left            =   2565
         TabIndex        =   29
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ҳ����"
         Height          =   180
         Index           =   23
         Left            =   5505
         TabIndex        =   26
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʵ�ս��"
         Height          =   180
         Index           =   22
         Left            =   5520
         TabIndex        =   25
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ�ս���д"
         Height          =   180
         Index           =   21
         Left            =   2565
         TabIndex        =   23
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ�ս��"
         Height          =   180
         Index           =   20
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�������"
      Height          =   930
      Left            =   1440
      TabIndex        =   17
      Top             =   8280
      Visible         =   0   'False
      Width           =   1920
      Begin VB.TextBox ctxtInput 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   300
         Index           =   19
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "1.00"
         Top             =   225
         Width           =   480
      End
      Begin VB.CheckBox cchk��ӡ���۱��� 
         Caption         =   "��ӡ���۱���"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.UpDown cupd�޸Ĵ��۱��� 
         Height          =   360
         Left            =   1500
         TabIndex        =   20
         Top             =   195
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���۱���"
         Height          =   180
         Index           =   19
         Left            =   165
         TabIndex        =   31
         Top             =   285
         Width           =   720
      End
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   979
      ButtonWidth     =   1455
      ButtonHeight    =   926
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkԤ�� 
         Caption         =   "��ӡǰԤ��"
         Height          =   255
         Left            =   9120
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "�����޸��嵥 "
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   120
      TabIndex        =   33
      Top             =   1920
      Width           =   11700
      Begin VB.Frame Frame1 
         Caption         =   "˫��ѡ���շ���Ŀ��"
         Height          =   5595
         Left            =   5880
         TabIndex        =   42
         Top             =   120
         Width           =   5775
         Begin VB.ListBox clst�շѱ�׼ 
            Height          =   4200
            Left            =   120
            TabIndex        =   55
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   4800
            TabIndex        =   10
            Top             =   5160
            Width           =   735
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   3360
            TabIndex        =   9
            Top             =   5160
            Width           =   735
         End
         Begin VB.TextBox ctxt�շ���Ŀ 
            Height          =   270
            Left            =   1200
            TabIndex        =   8
            Top             =   5160
            Width           =   1335
         End
         Begin VB.ComboBox Ccbo�շ���Ŀ���� 
            Height          =   300
            Left            =   2760
            TabIndex        =   44
            Top             =   600
            Width           =   2775
         End
         Begin VB.ListBox clst�շ���Ŀ 
            Height          =   3840
            Left            =   2760
            TabIndex        =   43
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label clblName 
            Height          =   180
            Left            =   2520
            TabIndex        =   50
            Top             =   4800
            Width           =   1410
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   2
            Left            =   4200
            TabIndex        =   49
            Top             =   5160
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "���ۣ�"
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   48
            Top             =   5160
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "�շ���Ŀ��"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   47
            Top             =   5160
            Width           =   900
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�շѱ�׼"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Clab�շ���Ŀ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�շ���Ŀ����"
            Height          =   180
            Left            =   2760
            TabIndex        =   45
            Top             =   360
            Width           =   1200
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   5460
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5685
         _cx             =   60368684
         _cy             =   60368287
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   -1  'True
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Del��������ɾ����ǰѡ�е���Ŀ"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   3960
         Width           =   2970
      End
   End
End
Attribute VB_Name = "frmֱ���շ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

'����������
Public pstr�շѱ�� As String

Dim WithEvents mobj����ͨ�ö��� As cls����ͨ�ö���
Attribute mobj����ͨ�ö���.VB_VarHelpID = -1


Private Const �շ�_�շѱ�� = 1
Private Const �շ�_������ = 2
Private Const �շ�_���ѵ�λ = 3

Private Const ���۱��� = 19
Private Const Ӧ�ս�� = 20
Private Const Ӧ�ս���д = 21
Private Const ʵ�ս�� = 22
Private Const �Ҳ���� = 23
Private Const �������� = 24

Private Const �����嵥_�շ���Ŀ��� = 0
Private Const �����嵥_�շ���Ŀ���� = 1
Private Const �����嵥_���� = 2
Private Const �����嵥_���� = 3
Private Const �����嵥_��� = 4


Dim mstrUndoCount As String          '���ڱ�������ԭ�����ַ���,�Ա������벻�Ϸ�ʱ�ܹ���ԭ
Dim mstrUndoMoney As String          '���ڱ�������ԭ�����ַ���,�Ա������벻�Ϸ�ʱ�ܹ���ԭ
Dim mstrUndoItemName As String

Dim mcur��С���� As Currency
Dim mcur��󵥼� As Currency

Dim mstr���ѵ�λ��� As String  '�ӵ�λ��λ�ӿڵõ��Ľ��ѵ�λ�ı��
Dim mint���ѷ�ʽ��� As Integer '���ѷ�ʽ�ı��

Dim mcur�ܽ�� As Currency

Dim mint���ۿ��� As Integer
Dim mint��Ŀ���� As Integer
Dim msng���۱��� As Single
'�μ�����
Dim mblnʹ�� As Boolean         '�Ƿ���ʹ��ϵͳ
Dim mint�Ƿ��Ҽ� As Integer     '�ڽ��ѵ�λ�ı������Ƿ�ʹ�����Ҽ�
Dim mstr�շѱ�� As String      '���������¼�շѱ��
Dim mblntemp As Boolean         '�ж�������Ŀ�����Ƿ�ִ�й�
Dim mbln���ƺŶ� As Boolean     '�Ƿ�ʹ���շ�Ա�Ŷο��ƹ���

'�޸ģ�2002-10-17����������ӡǰԤ����
Private mobj����  As cls�û���������


Private Sub Ccbo�շ���Ŀ����_Click()
    On Error GoTo errHandler
   
    Dim lobjRec As Object            '���������¼���ݼ�
    
    '�����շ���Ŀ��������,��ȡ�շѱ��ǰ׺
    Set lobjRec = dafuncGetData("select �շ���Ŀ��� from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ����= '" & Ccbo�շ���Ŀ����.Text & "'")
    
    '��ȡ�¼��շ���Ŀ
    Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where left(�շ���Ŀ���,3)='" & Left$(lobjRec("�շ���Ŀ���"), 3) & "' and len(�շ���Ŀ���)>3")
    clst�շ���Ŀ.Clear
    Do While Not lobjRec.EOF
        clst�շ���Ŀ.AddItem lobjRec("�շ���Ŀ����") & " " & lobjRec("�շ���Ŀ���")
        lobjRec.MoveNext
    Loop
    Exit Sub
errHandler:
    MsgBox "��ȡ����ʾָ��������շ���Ŀʧ�ܣ�" & Error, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
End Sub

Private Sub ccmb��Ӧҵ��_Click()
'    If ccmb��Ӧҵ��.ListIndex = 0 Then
'        ccmb������.Visible = True
'    Else
'        ccmb������.Visible = False
'    End If
    
End Sub

Private Sub ccmb�շѱ�׼_Click()
End Sub

Private Sub ccmd��λ_Click()
    Dim lrds������Ϣ As Object               '��λ�Ĵ�����Ϣ
    Dim lrdsTemp As Object
    
    On Error GoTo errHandler
    
    '���õ�λ�����Ķ�λ�ӿڻ�ȡ��λ��Ϣ
    Set lrdsTemp = pobj��λ��λ.func��λ�򵥶�λ(100, 100)
    If Not (lrdsTemp Is Nothing) Then
        If lrdsTemp.RecordCount > 0 Then
            '��ʾ��λ����`
            ctxtInput(�շ�_���ѵ�λ).Text = lrdsTemp("��λ����")
            '��ʾ�������ࡢƬ��
            ccmb��������.Text = lrdsTemp("��������")
            ccmbƬ��.Text = IIf(IsNull(lrdsTemp("Ƭ��")), "", lrdsTemp("Ƭ��"))
            
            '���浥λ��������
            mstr���ѵ�λ��� = lrdsTemp("������")
            ctxtInput(�շ�_���ѵ�λ).SetFocus
        End If
    End If
    
    '��ѯ������Ϣ
    ctxtInput(���۱���).Text = "1.00"
    Set lrds������Ϣ = dafuncGetData("select * from �շѹ���_������Ϣ�� where ��λ���='" & mstr���ѵ�λ��� & "'")
    If Not (lrds������Ϣ.EOF) Then
        If mint���ۿ��� > 0 Then
            ctxtInput(���۱���).Text = IIf(IsNull(lrds������Ϣ("���۱���")), "1.00", lrds������Ϣ("���۱���"))
        End If
    End If

    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ccmd��λ_Click", Err.Number, Err.Description, False
End Sub

Private Sub clst�շѱ�׼_DblClick()
    Dim lrds�շѱ�׼ As Object
    Dim i As Integer
    Dim lcurMoney As Currency
    
    On Error GoTo errHandler
    
    Set lrds�շѱ�׼ = dafuncGetData("select a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,���=a.����*a.���� from �շѹ���_�շѱ�׼��Ϣ�� a,�շѹ���_�շ���Ŀ�ֵ�� b where b.�շ���Ŀ���=a.�շ���Ŀ��� and �շѱ�׼����='" & clst�շѱ�׼.Text & "'")
    
    If lrds�շѱ�׼.EOF Then
        sffuncMsg "�շѱ�׼�����շ���Ŀ��", sf����
        Exit Sub
    Else
        lrds�շѱ�׼.MoveFirst
        Dim llngItemCount As Long
        For i = 0 To lrds�շѱ�׼.RecordCount - 1
            If Not func�����Ŀ�Ƿ���ѡ(lrds�շѱ�׼("�շ���Ŀ���")) Then
                ctxt�շ���Ŀ = lrds�շѱ�׼("�շ���Ŀ���")
                ctxt�շ���Ŀ_LostFocus
                sub�����Ŀ
                llngItemCount = llngItemCount + 1
            End If
            lrds�շѱ�׼.MoveNext
        Next
        For i = 1 To cgrdDetail.Rows - 1
            lcurMoney = Format(lcurMoney + cgrdDetail.ValueMatrix(i, �����嵥_���), "0.00")
        Next
        mcur�ܽ�� = lcurMoney
        ctxtInput(Ӧ�ս��) = lcurMoney * Val(ctxtInput(���۱���).Text)
        ctxtInput(Ӧ�ս���д) = FuncConvertToCapsStr(Val(ctxtInput(Ӧ�ս��)))
        
'        If llngItemCount = lrds�շѱ�׼.RecordCount Then
'            MsgBox "�շѱ�׼�е������շ���Ŀ(" & llngItemCount & "��)����ӵ������嵥�У�" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
'        ElseIf llngItemCount = 0 Then
'            MsgBox "�շѱ�׼�е������շ���Ŀ�ڷ����嵥������ӣ�" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
'        Else
'            MsgBox "�շѱ�׼�в����շ���Ŀ�ڷ����嵥�������,����� " & llngItemCount & " ������ӵ������嵥��" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
'        End If
    End If
                
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ccmb�շѱ�׼_Click", Err.Number, Err.Description, False
End Sub

Private Sub clst�շ���Ŀ_Click()
    Dim lobjRec As Object
    On Error GoTo errHandler
   ctxt�շ���Ŀ = Right(clst�շ���Ŀ.List(clst�շ���Ŀ.ListIndex), Len(clst�շ���Ŀ.List(clst�շ���Ŀ.ListIndex)) - InStr(clst�շ���Ŀ.List(clst�շ���Ŀ.ListIndex), " "))
    
    Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���='" & ctxt�շ���Ŀ & "'")
    If lobjRec.RecordCount > 0 Then
        ctxt���� = lobjRec("����")
        ctxt���� = 1
        mcur��С���� = IIf(IsNull(lobjRec("��С����").Value), 0, lobjRec("��С����").Value)
        mcur��󵥼� = IIf(IsNull(lobjRec("��󵥼�").Value), 99999999, lobjRec("��󵥼�").Value)
        clblName.Caption = lobjRec!�շ���Ŀ����
        ctxt����.SelStart = 0
        ctxt����.SelLength = Len(ctxt����)
        ctxt����.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "clst�շ���Ŀ_Click", Err.Number, Err.Description, False
          
End Sub

Private Sub ctxtInput_Change(Index As Integer)
    On Error GoTo errhandle
    Static lcurMoney As Currency
    Static lintAge As Integer
    Static lsngR As Single

    Select Case Index
        Case Ӧ�ս��
            ctxtInput(Ӧ�ս���д).Text = FuncConvertToCapsStr(Val(ctxtInput(Ӧ�ս��).Text))
            ctxtInput(�Ҳ����).Text = Format(Val(ctxtInput(ʵ�ս��).Text) - Val(ctxtInput(Ӧ�ս��).Text), "0.00")
            
        Case ���۱���
            If ctxtInput(���۱���).Text = vbNullString Then ctxtInput(���۱���).Text = "1.00"
            If Val(ctxtInput(���۱���).Text) > 1 Then ctxtInput(���۱���).Text = "1.00"
            If Val(ctxtInput(���۱���).Text) < 0 Then ctxtInput(���۱���).Text = "0.00"
            If Not IsNumeric(ctxtInput(���۱���).Text) Then ctxtInput(���۱���).Text = "1.00"
            
            
            ctxtInput(Ӧ�ս��).Text = mcur�ܽ�� * Val(ctxtInput(���۱���).Text)
            
            ctxtInput(Ӧ�ս���д).Text = FuncConvertToCapsStr(Val(ctxtInput(Ӧ�ս��).Text))
            ctxtInput(�Ҳ����).Text = Format(Val(ctxtInput(ʵ�ս��).Text) - Val(ctxtInput(Ӧ�ս��).Text), "0.00")
            
        Case ʵ�ս��
            If ctxtInput(ʵ�ս��).Text = vbNullString Then ctxtInput(ʵ�ս��).Text = 0
            If Not IsNumeric(ctxtInput(ʵ�ս��).Text) Then
                ctxtInput(ʵ�ս��).Text = CStr(lcurMoney)
            Else
                lcurMoney = Val(ctxtInput(ʵ�ս��).Text)
            End If
            ctxtInput(�Ҳ����).Text = Format(Val(ctxtInput(ʵ�ս��).Text) - Val(ctxtInput(Ӧ�ս��).Text), "0.00")
            
        Case �Ҳ����
            If Val(ctxtInput(�Ҳ����).Text) < 0 Then
                ctxtInput(�Ҳ����).ForeColor = &HFF
            Else
                ctxtInput(�Ҳ����).ForeColor = &HFF0000
            End If
    End Select
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ctxtInput_Change", Err.Number, Err.Description, False
End Sub

Private Sub ctxtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errhandle
    Select Case KeyAscii
        Case vbKeyReturn
            If Index = ʵ�ս�� And ctlb������.Buttons(1).Enabled Then
                Call mobj����ͨ�ö���_BeforeOperate("�շ�", False)
            ElseIf Index = �շ�_���ѵ�λ Then
                Ccbo�շ���Ŀ����.SetFocus
            End If
        Case Else
            If Index = �շ�_���ѵ�λ Then
                ctxtInput(���۱���).Text = "1.00"
            End If
        End Select
Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ctxtInput_KeyPress", Err.Number, Err.Description, False
End Sub




Private Sub cgrdDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim lcurMoney As Currency
    
    On Error GoTo errhandle
    'ctlb������.Buttons("�շ�(&G)").Enabled = True
    Select Case cgrdDetail.TextMatrix(0, Col)
        Case "����"
            '�ж�������Ƿ���ֵ
            If Len(cgrdDetail.TextMatrix(Row, Col)) > 4 Then
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
            Else
                If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) And Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    '����ֵ
                    '������
                    cgrdDetail.TextMatrix(Row, �����嵥_���) = cgrdDetail.TextMatrix(Row, �����嵥_����) * cgrdDetail.TextMatrix(Row, �����嵥_����)
                Else
                    '������ֵ
                    'Undo
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
                End If
            End If
        Case "����"
            Dim lcur���� As Currency
            If mcur��С���� = mcur��󵥼� Then
                sffuncMsg "���շ���Ŀ�����Ѷ�,�����޸ģ�", sf����
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                Exit Sub
            End If
            
            If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) Then
                If Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    If Val(cgrdDetail.TextMatrix(Row, Col)) <= mcur��󵥼� And Val(cgrdDetail.TextMatrix(Row, Col)) >= mcur��С���� Then
                        cgrdDetail.TextMatrix(Row, �����嵥_���) = cgrdDetail.TextMatrix(Row, �����嵥_����) * cgrdDetail.TextMatrix(Row, �����嵥_����)
                    Else
                        sffuncMsg "����ĵ��۳�����Χ��", sf����
                        cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                    End If
                Else
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                End If
            Else
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
            End If
        Case "�շ���Ŀ����"
            If cgrdDetail.TextMatrix(Row, Col) = "" Then
                sffuncMsg "���������շ���Ŀ���ƣ�", sf����
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoItemName
                Exit Sub
            End If
            '�ж���Ŀ�����Ƿ��ظ���
            For i = 1 To cgrdDetail.Rows - 1
                If i <> Row And cgrdDetail.TextMatrix(Row, Col) = cgrdDetail.TextMatrix(i, �����嵥_�շ���Ŀ����) Then
                    sffuncMsg "�շ���Ŀ���Ʋ������ظ���", sf����
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoItemName
                    Exit Sub
                End If
            Next
        Case Else
    End Select
    
    For i = 1 To cgrdDetail.Rows - 1
        lcurMoney = lcurMoney + cgrdDetail.ValueMatrix(i, �����嵥_���)
    Next
    mcur�ܽ�� = lcurMoney
    
    ctxtInput(Ӧ�ս��) = lcurMoney * Val(ctxtInput(���۱���).Text)
    ctxtInput(Ӧ�ս���д) = FuncConvertToCapsStr(Val(ctxtInput(Ӧ�ս��)))
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "cing�����嵥_AfterEdit", Err.Number, Err.Description, False
    
End Sub

Private Sub cgrdDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
     
    Select Case Col
        Case �����嵥_����
            ctlb������.Buttons("�շ�(&G)").Enabled = False
            mstrUndoCount = cgrdDetail.TextMatrix(Row, Col)
            
        Case �����嵥_����
            ctlb������.Buttons("�շ�(&G)").Enabled = False
                        
            '��ȡ��С����,��󵥼�.
            Dim lobjRec As Object
            Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���='" & cgrdDetail.TextMatrix(Row, 0) & "'")
            If lobjRec.RecordCount > 0 Then
                mcur��С���� = IIf(IsNull(lobjRec("��С����").Value), 0, lobjRec("��С����").Value)
                mcur��󵥼� = IIf(IsNull(lobjRec("��󵥼�").Value), 99999999, lobjRec("��󵥼�").Value)
            Else
                sffuncMsg "δ�ҵ����շ���Ŀ��������Ϣ����������Ϣ�����ѱ��޸Ļ�ɾ�������˳��շѽ��棬���½��룡"
            End If
            mstrUndoMoney = cgrdDetail.TextMatrix(Row, Col)
        Case �����嵥_�շ���Ŀ����
            ctlb������.Buttons("�շ�(&G)").Enabled = False
            mstrUndoItemName = cgrdDetail.TextMatrix(Row, Col)
        Case Else
            ctlb������.Buttons("�շ�(&G)").Enabled = True
            Cancel = True
    End Select
End Sub



Private Sub cgrdDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete
            mobj����ͨ�ö���_BeforeOperate "ɾ��", False
    End Select

End Sub

Private Sub cgrdDetail_LostFocus()
    On Error Resume Next
    ctlb������.Buttons("�շ�(&G)").Enabled = True
End Sub


Private Sub clst�շ���Ŀ_DblClick()
    Dim lobjRec As Object
    On Error GoTo errHandler
   ctxt�շ���Ŀ = Right(clst�շ���Ŀ.List(clst�շ���Ŀ.ListIndex), Len(clst�շ���Ŀ.List(clst�շ���Ŀ.ListIndex)) - InStr(clst�շ���Ŀ.List(clst�շ���Ŀ.ListIndex), " "))
    
    Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���='" & ctxt�շ���Ŀ & "'")
    If lobjRec.RecordCount > 0 Then
        ctxt���� = lobjRec("����")
        ctxt���� = 1
        clblName.Caption = lobjRec!�շ���Ŀ����
        
        '����շ���Ŀ
        If Not func�����Ŀ�Ƿ���ѡ(ctxt�շ���Ŀ) Then
            sub�����Ŀ
        End If
        
        ctxt�շ���Ŀ = ""
        clblName = ""
        ctxt���� = ""
    End If
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "clst�շ���Ŀ_DblClick", Err.Number, Err.Description, False
          
End Sub

Private Sub sub�����Ŀ()
    Dim lcurMoney As Double
    Dim i As Long
    On Error GoTo errHandler
    
    cgrdDetail.AddItem ctxt�շ���Ŀ & vbTab & clblName & vbTab & _
                    ctxt���� & vbTab & ctxt���� & vbTab & Format(Val(ctxt����) * Val(ctxt����), "0.0")
    For i = 1 To cgrdDetail.Rows - 1
        lcurMoney = Format(lcurMoney + cgrdDetail.ValueMatrix(i, �����嵥_���), "0.00")
    Next
    mcur�ܽ�� = lcurMoney
    ctxtInput(Ӧ�ս��) = lcurMoney * Val(ctxtInput(���۱���).Text)
    ctxtInput(Ӧ�ս���д) = FuncConvertToCapsStr(Val(ctxtInput(Ӧ�ս��)))

    
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "sub�����Ŀ", Err.Number, Err.Description, True
End Sub

Private Sub ctxt����_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 And clblName <> "" Then
        Dim lcur���� As Currency
        If mcur��С���� = mcur��󵥼� And ctxt���� <> mcur��С���� Then
            ctxt���� = mstrUndoMoney
            ctxt�շ���Ŀ.SetFocus
            sffuncMsg "���շ���Ŀ�����Ѷ��������޸ģ�", sf����
            Exit Sub
        Else
            If IsNumeric(ctxt����) Then
                If Val(ctxt����) > 0 Then
                    If Val(ctxt����) <= mcur��󵥼� And Val(ctxt����) >= mcur��С���� Then
                        
                    Else
                        ctxt���� = mstrUndoMoney
                        ctxt�շ���Ŀ.SetFocus
                        sffuncMsg "����ĵ��۳�����Χ��", sf����
                        Exit Sub
                    End If
                Else
                    ctxt���� = mstrUndoMoney
                End If
            Else
                ctxt���� = mstrUndoMoney
            End If
            ctxt����.SelStart = 0
            ctxt����.SelLength = Len(ctxt����)
            ctxt����.SetFocus
        End If
        
    End If
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ctxt����_KeyUp", Err.Number, Err.Description, False
    
End Sub

Private Sub ctxtƱ�ݺ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        ctxt�շ���Ŀ.SetFocus
    End If
End Sub

Private Sub ctxtƱ�ݺ�_LostFocus()
    Dim lstrƱ�ݺ� As String
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    
    If Not IsNumeric(ctxtƱ�ݺ�) Then
        MsgBox "Ʊ�ݺű��������֣�", vbInformation, "ϵͳ��ʾ"
        ctxtƱ�ݺ�.SetFocus
        Exit Sub
    End If
    If mbln���ƺŶ� Then
        '����Ʊ�ݺ��Ƿ����ѷ���ĺŶ���
        Set lobjRec = dafuncGetData("select ���,ֹ�� from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where '" & ctxtƱ�ݺ� & "' between ��� and ֹ�� and �û����='" & um�û���� & "' and �Ƿ�����='��'")
        If lobjRec.RecordCount = 0 Then
            MsgBox "�����õĵ�ǰƱ�ݺŲ�����δ�����Ʊ�ݺŶη�Χ�ڣ����ܽ����շѣ������½������ã�", vbInformation, "ϵͳ��ʾ"
            ctlb������.Buttons(1).Enabled = False
            Exit Sub
        End If
        clblCurNoArea = lobjRec(0) & "��" & lobjRec(1)
    End If
    lstrƱ�ݺ� = Format(Val(ctxtƱ�ݺ�) - 1, String(Len(ctxtƱ�ݺ�), "0"))
    dafuncGetData "update ϵͳ����_ϵͳ������ɼ�¼�� set ��ǰֵ=" & lstrƱ�ݺ� & " where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'"
    ctlb������.Buttons(1).Enabled = True
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ctxtƱ�ݺ�_LostFocus", Err.Number, Err.Description, False
End Sub

Private Sub ctxt�շ���Ŀ_GotFocus()
    On Error Resume Next
    If ctxt�շ���Ŀ = "�޸���Ŀ��" Then
        ctxt�շ���Ŀ = ""
    End If
End Sub

Private Sub ctxt�շ���Ŀ_LostFocus()
    Dim lobjRec As Object
    Dim pint��Ŀ���� As Integer
    On Error GoTo errHandler
    
    '�����շ���Ŀ��Ż�ȡ�շ���Ŀ���ơ�
    clblName.Caption = ""
    pint��Ŀ���� = Val(pobj�շѹ���.ҵ������("��Ŀ����"))
    If pint��Ŀ���� = 0 Then pint��Ŀ���� = 2
    
    If ctxt�շ���Ŀ <> "" And Len(ctxt�շ���Ŀ) = 3 * pint��Ŀ���� Then
    
        Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���='" & ctxt�շ���Ŀ & "'")
        If lobjRec.RecordCount > 0 Then
            clblName.Caption = lobjRec!�շ���Ŀ����
            ctxt���� = lobjRec("����")
            mstrUndoMoney = lobjRec("����")
            mcur��С���� = IIf(IsNull(lobjRec("��С����").Value), 0, lobjRec("��С����").Value)
            mcur��󵥼� = IIf(IsNull(lobjRec("��󵥼�").Value), 99999999, lobjRec("��󵥼�").Value)

            If ctxt���� = "" Then ctxt���� = 1
        Else
            ctxt�շ���Ŀ = "�޸���Ŀ��"
        End If
    End If
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ctxt�շ���Ŀ_LostFocus", Err.Number, Err.Description, False
    
End Sub

Private Sub ctxt����_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        If clblName <> "" And ctxt�շ���Ŀ <> "" Then
            If Not func�����Ŀ�Ƿ���ѡ(ctxt�շ���Ŀ.Text) Then
                sub�����Ŀ
            End If
        End If
        ctxt�շ���Ŀ = ""
        clblName = ""
        ctxt���� = ""
        ctxt�շ���Ŀ.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "ctxt����_KeyDown", Err.Number, Err.Description, False

End Sub

Private Sub cupd�޸Ĵ��۱���_DownClick()
    On Error Resume Next
    If Val(ctxtInput(���۱���).Text) > 0 Then
        ctxtInput(���۱���).Text = Format(CStr(Val(ctxtInput(���۱���).Text) - 0.01), "0.00")
    Else
        ctxtInput(���۱���).Text = "0.00"
    End If
End Sub

Private Sub cupd�޸Ĵ��۱���_UpClick()
On Error GoTo errhandle
    If Val(ctxtInput(���۱���).Text) < 1 Then
        ctxtInput(���۱���).Text = Format(CStr(Val(ctxtInput(���۱���).Text) + 0.01), "0.00")
    Else
        ctxtInput(���۱���).Text = "1.00"
    End If
Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "Form_UpClick", Err.Number, Err.Description, False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And ActiveControl.Name <> "ctxt����" And ActiveControl.Name <> "ctxt����" Then
        SendKeys Chr(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim lcol������ As Collection
    Dim i As Long, lobjRec As Recordset
    On Error GoTo errhandle
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
    mblnʹ�� = True
        
    Set mobj����ͨ�ö��� = New cls����ͨ�ö���
    Set mobj����ͨ�ö���.Form = Me
    Set mobj����ͨ�ö���.c������ = ctlb������
    
    Set lcol������ = New Collection
    
    lcol������.Add "�շ�(&G)101"
    lcol������.Add "|"
    lcol������.Add "ɾ��"
    lcol������.Add "���"
    lcol������.Add "|"
    lcol������.Add "�˳�"
    
    mobj����ͨ�ö���.subInitialize lcol������, ""
    
    mint���ۿ��� = Val(pobj�շѹ���.ҵ������("���ۿ���"))
    mint��Ŀ���� = Val(pobj�շѹ���.ҵ������("��Ŀ����"))
    
    sub��ʼ������
    
    If pstr�շѱ�� <> "" Then
        '�ڲ��շ�,��ʾ������Ϣ��
        sub��ʾ������Ϣ
        
        'û��Ȩ���޸ģ������޸ķ�����Ϣ��
        If Not umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣ�޸�") Then
            Frame4.Enabled = False
            Frame1.Enabled = False
            cgrdDetail.Editable = False
'            Label4.Caption = "��û��Ȩ���޸��ڲ��շ���Ϣ��"
        End If
    End If
    ctxtƱ�ݺ� = func��ȡƱ�ݺ�()
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
Private Function func��ȡƱ�ݺ�() As String
    Dim lstrƱ�ݺ� As String
    Dim lLen As Integer
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select ��ǰֵ from ϵͳ����_ϵͳ������ɼ�¼�� where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'")
    If mbln���ƺŶ� Then
        If lobjRec.RecordCount = 0 Then
            '�ҳ���һ��δʹ�õ���С�Ŷ�
            Set lobjRec = dafuncGetData("select ���,ֹ�� from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where �û����='" & um�û���� & "' and �Ƿ�����='��' order by ���")
            If lobjRec.RecordCount = 0 Then
                MsgBox "����ǰû����δ�����Ʊ�ݺŶ���Ϣ�����ܽ����շѣ�", vbInformation, "ϵͳ��ʾ"
                ctlb������.Buttons(1).Enabled = False
                func��ȡƱ�ݺ� = ""
                Exit Function
            Else
                clblCurNoArea = lobjRec(0) & "��" & lobjRec(1)
                lstrƱ�ݺ� = lobjRec(0)
                lLen = Len(lstrƱ�ݺ�)
                dafuncGetData "insert into ϵͳ����_ϵͳ������ɼ�¼��(ҵ������,�������,��������,��ǰֵ,����,�Ƿ����ر�,��ǰ���) values('�շѹ���" & um�û���� & "','�վݺ�','C'," & Format(Val(lstrƱ�ݺ�) - 1, String(lLen, "0")) & ",9,'��',2008)"
    '            dafuncGetData "update ϵͳ����_ϵͳ������ɼ�¼�� set ��ǰֵ='" & Format(Val(lstrƱ�ݺ�) - 1, String(lLen, "0")) & "' where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'"
            End If
        Else
            lstrƱ�ݺ� = IIf(IsNull(lobjRec(0)), "0", lobjRec(0))
            lLen = Len(lstrƱ�ݺ�)
            '����Ʊ�ݺ��Ƿ���ȷ
            Set lobjRec = dafuncGetData("select ���,ֹ�� from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where '" & lstrƱ�ݺ� & "' between ��� and ֹ�� and �û����='" & um�û���� & "' and �Ƿ�����='��'")
            If lobjRec.RecordCount = 0 Then
                '���ú��Ƿ����µĺŶε���ʼ��
                Set lobjRec = dafuncGetData("select ���,ֹ�� from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where '" & Format(Val(lstrƱ�ݺ�) + 1, String(lLen, "0")) & "' between ��� and ֹ�� and �û����='" & um�û���� & "' and �Ƿ�����='��'")
                If lobjRec.RecordCount = 0 Then
                    MsgBox "�����õĵ�ǰƱ�ݺŲ�����δ�����Ʊ�ݺŶη�Χ�ڣ����ܽ����շѣ������½������ã�", vbInformation, "ϵͳ��ʾ"
                    ctlb������.Buttons(1).Enabled = False
                    func��ȡƱ�ݺ� = ""
                    Exit Function
                Else
                    clblCurNoArea = lobjRec(0) & "��" & lobjRec(1)
                End If
            Else
                clblCurNoArea = lobjRec(0) & "��" & lobjRec(1)
            End If
            '���úŶ��Ƿ��Ѿ����꣬��Ҫ����Ʊ��
            Set lobjRec = dafuncGetData("select ID from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where �û����='" & um�û���� & "' and �Ƿ�����='��' and ֹ��='" & lstrƱ�ݺ� & "'")
            If lobjRec.RecordCount Then
                '���úŶ���Ϊ������
                dafuncGetData "update �շѹ���_�շ�Ա�Ŷ���Ϣ�� set �Ƿ�����='��' where ID=" & lobjRec(0)
                '�ҳ���һ��δʹ�õ���С�Ŷ�
                Set lobjRec = dafuncGetData("select ���,ֹ�� from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where �û����='" & um�û���� & "' and �Ƿ�����='��' order by ���")
                If lobjRec.RecordCount = 0 Then
                    MsgBox "����ǰû����δ�����Ʊ�ݺŶ���Ϣ�����ܽ����շѣ�", vbInformation, "ϵͳ��ʾ"
                    ctlb������.Buttons(1).Enabled = False
                    '����䵱ǰƱ�ݺ����ã��������¿�ʼ
                    dafuncGetData "delete ϵͳ����_ϵͳ������ɼ�¼�� where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'"
                    func��ȡƱ�ݺ� = ""
                    Exit Function
                Else
                    clblCurNoArea = lobjRec(0) & "��" & lobjRec(1)
                    lstrƱ�ݺ� = lobjRec(0)
                    lLen = Len(lstrƱ�ݺ�)
                    dafuncGetData "update ϵͳ����_ϵͳ������ɼ�¼�� set ��ǰֵ=" & Format(Val(lstrƱ�ݺ�) - 1, String(lLen, "0")) & " where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'"
                    MsgBox "��ǰƱ�ݺŶ��Ѿ����꣬���ڴ�ӡ���ϰ�װ��ȷ����Ʊ�ݣ�", vbInformation, "ϵͳ��ʾ"
                End If
            Else
                lstrƱ�ݺ� = Format(Val(lstrƱ�ݺ�) + 1, String(lLen, "0"))
                '����վݺ��Ƿ��ظ�
                Set lobjRec = dafuncGetData("select * from �շѹ���_������Ϣ�� where �վݺ�='" & lstrƱ�ݺ� & "'")
                If lobjRec.RecordCount Then
                    MsgBox "ϵͳ���Ѿ����ڸ�Ʊ�ݺ��ˣ���ע���飡", vbInformation, "ϵͳ��ʾ"
                End If
            End If
        End If
    Else
        If lobjRec.RecordCount = 0 Then
            lstrƱ�ݺ� = "1"
        ElseIf ctxtƱ�ݺ� = "" Then
            lstrƱ�ݺ� = Format(Val(lobjRec(0)) + 1, String(Len(lobjRec(0)), "0"))
        Else
            lstrƱ�ݺ� = Format(Val(ctxtƱ�ݺ�) + 1, String(Len(ctxtƱ�ݺ�), "0"))
        End If
        '����վݺ��Ƿ��ظ�
        Set lobjRec = dafuncGetData("select * from �շѹ���_������Ϣ�� where �վݺ�='" & lstrƱ�ݺ� & "'")
        If lobjRec.RecordCount Then
            MsgBox "ϵͳ���Ѿ����ڸ�Ʊ�ݺ��ˣ���ע���飡", vbInformation, "ϵͳ��ʾ"
        End If
    End If
    func��ȡƱ�ݺ� = lstrƱ�ݺ�
End Function

Private Sub sub��ʾ������Ϣ()
    Dim lobjRec As Object
    Dim i As Long
    
    On Error GoTo errHandler
    
    If pstr�շѱ�� <> "" Then
        '�޸��շѼ�¼��
        Set lobjRec = dafuncGetData("select a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,�շ���Ŀ����=(select �շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���=a.�շ���Ŀ���),a.����,������λ=(select ������λ from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���=a.�շ���Ŀ���),a.����,a.���,a.�շ�״̬,a.���ѷ�ʽ,a.������,a.���ѵ�λ���,���ѵ�λ���� ,a.��������,a.�˷�����,�շ��˱��=a.�շ���,�շ���=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.�շ���),�˷��˱��=a.�˷���,�˷���=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.�˷���) ,���ܿ��Ҿ����˱��=a.���ܿ��Ҿ�����,���ܿ��Ҿ�����=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.���ܿ��Ҿ�����),���ܿ��ұ��,���ܿ���=(select ���� from ϵͳ����_�����ֵ�� where ���=a.���ܿ��ұ��),���۱���,��ע1,��ע2  from �շѹ���_������Ϣ�� a where �շѱ��='" & pstr�շѱ�� & "'  and �շ�״̬=0")
        
        cgrdDetail.Rows = 1
    
        Do While Not lobjRec.EOF
            cgrdDetail.AddItem lobjRec("�շ���Ŀ���") & vbTab & _
                lobjRec("�շ���Ŀ����") & vbTab & _
                lobjRec("����") & vbTab & _
                lobjRec("����") & vbTab & _
                lobjRec("���")
            lobjRec.MoveNext
        Loop
        If lobjRec.RecordCount > 0 Then
            lobjRec.MoveFirst
            mstr���ѵ�λ��� = IIf(IsNull(lobjRec("���ѵ�λ���").Value), "", lobjRec("���ѵ�λ���").Value)
        
            ctxtInput(�շ�_�շѱ��).Text = lobjRec("�շѱ��")
            
            If IIf(IsNull(lobjRec("���ܿ���")), "", lobjRec("���ܿ���")) <> "" Then
                For i = 0 To ccmb���ܿ���.ListCount - 1
                    If ccmb���ܿ���.List(i) = IIf(IsNull(lobjRec("���ܿ���")), "", lobjRec("���ܿ���")) Then
                        ccmb���ܿ���.ListIndex = i
                        Exit For
                    End If
                Next
            Else
                ccmb���ܿ���.ListIndex = -1
            End If
            
            ccmb��������.Text = IIf(IsNull(lobjRec("��ע1").Value), "", lobjRec("��ע1").Value)
            ccmbƬ��.Text = IIf(IsNull(lobjRec("��ע2").Value), "", lobjRec("��ע2").Value)
        
        
            ctxtInput(�շ�_������).Text = lobjRec("������")
            ctxtInput(�շ�_���ѵ�λ).Text = IIf(IsNull(lobjRec("���ѵ�λ����").Value), "", lobjRec("���ѵ�λ����").Value)
        
            Set lobjRec = dafuncGetData("select ���۱��� from �շѹ���_������Ϣ�� where ��λ���='" & mstr���ѵ�λ��� & "'")
            
            If lobjRec.EOF Then
                ctxtInput(���۱���).Text = "1.00"
            Else
                ctxtInput(���۱���).Text = Format(lobjRec("���۱���").Value, "0.00")
            End If
            
        End If
        
        Dim lcurMoney As Currency
        lcurMoney = 0
        For i = 1 To cgrdDetail.Rows - 1
            lcurMoney = lcurMoney + cgrdDetail.ValueMatrix(i, �����嵥_���)
        Next
        mcur�ܽ�� = lcurMoney
        
        ctxtInput(Ӧ�ս��) = lcurMoney * Val(ctxtInput(���۱���).Text)
        ctxtInput(Ӧ�ս���д) = FuncConvertToCapsStr(Val(ctxtInput(Ӧ�ս��)))
        
    End If

    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "sub��ʾ������Ϣ", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mobj����ͨ�ö��� = Nothing
    
End Sub


Private Sub sub�������()
    Dim i As Integer
    
    On Error GoTo errhandle
    mcur�ܽ�� = 0
    
    '�ڱ������Ҫ����շѱ��;�켽��;2002/9/30
    mstr���ѵ�λ��� = ""
  
    Dim lobjCtrl As Control
    For Each lobjCtrl In ctxtInput
        lobjCtrl.Text = ""
    Next
    cgrdDetail.Rows = 1
    
    ctxtInput(���۱���).Text = "1.00"
    
    For i = Ӧ�ս�� To ��������
        ctxtInput(i).Text = ""
    Next
    
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select getdate()")
    ctxtInput(��������).Text = Format(lobjRec(0), "yyyy-mm-dd")
    
'    ctxtInput(��������).Text = Date
    ctxtInput(���۱���).Text = "1.00"
    ccmb���ܿ���.Text = um�û���������
    
    If ctxtInput(�շ�_�շѱ��).Enabled Then
        ctxtInput(�շ�_�շѱ��).SetFocus
    Else
        ctxtInput(�շ�_������).SetFocus
    End If
    
    
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "sub����շѽ���", Err.Number, Err.Description, True
End Sub


Private Sub sub��ʼ������()
    
    On Error GoTo errhandle
    
    Dim lobj�շѱ�׼ As Object
    Dim lobj���� As Object
    Dim lobj���ѷ�ʽ As Object

    mstrUndoCount = ""
    mstrUndoMoney = ""
    mstr���ѵ�λ��� = ""
    mint���ѷ�ʽ��� = 0
    mcur�ܽ�� = 0
    
    Dim i As Long
    Dim j As Long
    
    Set lobj�շѱ�׼ = dafuncGetData("select �շѱ�׼����,���Ƿ� from �շѹ���_�շѱ�׼��Ϣ�� group by ���Ƿ�,�շѱ�׼����")
    Set lobj���� = dafuncGetData("select * from ϵͳ����_�����ֵ��")
    Set lobj���ѷ�ʽ = dafuncGetData("select * from �շѹ���_���ѷ�ʽ�ֵ��")
    
    
    '��ʼ�� "cing�����嵥"
    With cgrdDetail
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, �����嵥_�շ���Ŀ���) = "�շ���Ŀ���"
        .ColWidth(�����嵥_�շ���Ŀ���) = 1310
        .ColAlignment(�����嵥_�շ���Ŀ���) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_�շ���Ŀ����) = "�շ���Ŀ����"
        .ColWidth(�����嵥_�շ���Ŀ����) = 1320
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 480
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 500
        
        .TextMatrix(0, �����嵥_���) = "���"
        .ColWidth(�����嵥_���) = 570
    End With
    
    '��ʼ�� "�շѱ�׼"
    clst�շѱ�׼.Clear
    Do While Not lobj�շѱ�׼.EOF
        clst�շѱ�׼.AddItem lobj�շѱ�׼("�շѱ�׼����").Value
        lobj�շѱ�׼.MoveNext
    Loop
    
    '��ʼ�� "���ܿ���"�б�
    ccmb���ܿ���.Clear
    If Not (lobj���� Is Nothing) Then
        Do While Not lobj����.EOF
            ccmb���ܿ���.AddItem lobj����("����").Value
            ccmb���ܿ���.ItemData(ccmb���ܿ���.ListCount - 1) = "1" & lobj����!���
            lobj����.MoveNext
        Loop
    End If
    
    ccmb���ܿ���.ListIndex = -1
    
    
    If umfuncУ���û�Ȩ��("�շѹ���_����") Then
        Frame6.Enabled = True
        Frame6.Caption = "�������"
        lblCaption(19).Enabled = True
        ctxtInput(19).Enabled = True
        cupd�޸Ĵ��۱���.Enabled = True
        cchk��ӡ���۱���.Enabled = True
        Select Case mint���ۿ���
            Case 0
                cupd�޸Ĵ��۱���.Enabled = False
                ctxtInput(19).Enabled = False
            Case 1
                cupd�޸Ĵ��۱���.Enabled = True
                ctxtInput(19).Enabled = True
            Case 2
                cupd�޸Ĵ��۱���.Enabled = False
                ctxtInput(19).Enabled = False
            Case Else
        End Select
        
    Else
        Frame6.Caption = "�������(��Ȩ��)"
        Frame6.Enabled = False
        lblCaption(19).Enabled = False
        ctxtInput(���۱���).Enabled = False
        cupd�޸Ĵ��۱���.Enabled = False
        cchk��ӡ���۱���.Enabled = False
    End If
        
    '��ʼ�����ѷ�ʽ�б�
    If Not (lobj���ѷ�ʽ Is Nothing) Then
        Do While Not lobj���ѷ�ʽ.EOF
            cmb���ѷ�ʽ.AddItem lobj���ѷ�ʽ("����").Value
            cmb���ѷ�ʽ.ItemData(cmb���ѷ�ʽ.ListCount - 1) = "1" & lobj���ѷ�ʽ("���")
            lobj���ѷ�ʽ.MoveNext
        Loop
        cmb���ѷ�ʽ.ListIndex = 0
    End If
    
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select getdate()")
    ctxtInput(��������).Text = Format(lobjRec(0), "yyyy-mm-dd")
    
    '��ȡ�շ���Ŀ���ࡣ
    Set lobjRec = dafuncGetData("select �շ���Ŀ���,�շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where len(�շ���Ŀ���)=3  order by �շ���Ŀ��� ")
    Do While Not lobjRec.EOF
        Ccbo�շ���Ŀ����.AddItem lobjRec("�շ���Ŀ����")
        lobjRec.MoveNext
    Loop
    
    Ccbo�շ���Ŀ����.ListIndex = 0
    
    '��ȡ��������
    Set lobjRec = dafuncGetData("select * from ϵͳ����_���������ֵ���ͼ order by ���")
    ccmb��������.Clear
    ccmb��������.AddItem ""
    Do While Not lobjRec.EOF
        ccmb��������.AddItem lobjRec("����").Value
        lobjRec.MoveNext
    Loop
    
    '��ȡƬ��
    Set lobjRec = dafuncGetData("select * from ϵͳ����_Ƭ���ֵ���ͼ order by ���")
    ccmbƬ��.Clear
    ccmbƬ��.AddItem ""
    Do While Not lobjRec.EOF
        ccmbƬ��.AddItem lobjRec("����").Value
        lobjRec.MoveNext
    Loop
    
    '��ȡ�������ʺ�.
    Set lobjRec = dafuncGetData("select ������+' '+�ʺ� from �շѹ���_���п��������ñ�")
    ccmb������.Clear
    Do While Not lobjRec.EOF
        ccmb������.AddItem lobjRec(0)
        
        lobjRec.MoveNext
    Loop
    If ccmb������.ListCount > 0 Then
        ccmb������.ListIndex = 0
    End If
    
    ccmb��Ӧҵ��.AddItem "һ��", 0
    ccmb��Ӧҵ��.AddItem "����", 1
    ccmb��Ӧҵ��.ListIndex = 0
    
    '��ȡ�Ƿ�ʹ���շ�ԱƱ�ݵĺŶο��ƹ���.
    Set lobjRec = dafuncGetData("select ����ֵ from �շѹ���_ҵ�����ñ� where ������Ŀ='���ƺŶ�'")
    If lobjRec.RecordCount = 0 Then
        mbln���ƺŶ� = False
    ElseIf lobjRec(0) = "0" Then
        mbln���ƺŶ� = False
    Else
        mbln���ƺŶ� = True
    End If
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "sub��ʼ������", Err.Number, Err.Description, True
End Sub


Private Sub mobj����ͨ�ö���_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long, j As Long
    Dim lobjRec As Recordset
    
    On Error GoTo errhandle
    Select Case Operate
        Case "�շ�"
            'У�����ݺϷ��ԡ�
            If Not ValidateData Then Exit Sub
            
            If ctxtƱ�ݺ� = "" Then
                MsgBox "Ʊ�ݺ����ò���ȷ�������շѣ�", vbInformation, "ϵͳ��ʾ"
                Exit Sub
            End If
            '����վݺ��Ƿ��ظ�
            Set lobjRec = dafuncGetData("select * from �շѹ���_������Ϣ�� where �վݺ�='" & ctxtƱ�ݺ� & "'")
            If lobjRec.RecordCount Then
                MsgBox "ϵͳ���Ѿ����ڸ�Ʊ�ݺ��ˣ�������¼������Ʊ�ݺţ�", vbInformation, "ϵͳ��ʾ"
                'ctxtƱ�ݺ�.SetFocus
                Exit Sub
            End If
            
            mint���ѷ�ʽ��� = Right(cmb���ѷ�ʽ.ItemData(cmb���ѷ�ʽ.ListIndex), Len(Trim(Str(cmb���ѷ�ʽ.ItemData(cmb���ѷ�ʽ.ListIndex)))) - 1)
            
            '�ռ�Ҫ����ķ�����Ϣ��
            Dim lstr���ܿ��ұ�� As String
            Dim lcol��¼ As Collection
            Dim lcol���� As Collection
            Dim lstr�շѱ�� As String
            
            If ccmb���ܿ���.ListIndex >= 0 Then
                lstr���ܿ��ұ�� = ccmb���ܿ���.ItemData(ccmb���ܿ���.ListIndex)
                lstr���ܿ��ұ�� = Right(lstr���ܿ��ұ��, Len(lstr���ܿ��ұ��) - 1)
            Else
                lstr���ܿ��ұ�� = um�û��������ұ��
            End If
            Set lcol���� = New Collection
            For i = 1 To cgrdDetail.Rows - 1
                Set lcol��¼ = New Collection
                For j = 0 To cgrdDetail.Cols - 1
                    lcol��¼.Add cgrdDetail.TextMatrix(i, j), cgrdDetail.TextMatrix(0, j)
                Next
                '����շ������ֶ�
                lcol��¼.Add ctxtInput(�շ�_������).Text, "������"
                lcol��¼.Add mstr���ѵ�λ���, "���ѵ�λ���"
                lcol��¼.Add ctxtInput(�շ�_���ѵ�λ).Text, "���ѵ�λ����"
                lcol��¼.Add lstr���ܿ��ұ��, "���ܿ��ұ��"
                lcol��¼.Add um�û����, "���ܿ��Ҿ�����"
                lcol��¼.Add ccmb��������.Text, "��ע1"
                lcol��¼.Add ccmbƬ��.Text, "��ע2"
                lcol����.Add lcol��¼
            Next
            
            '���滮����Ϣ��
            lstr�շѱ�� = pobj�շѹ���.func���۱���(lcol����, pstr�շѱ��)
            
            '�����շ�ȷ����Ϣ��
            Dim lcol�շѱ�ż� As Collection
            
            Set lcol�շѱ�ż� = New Collection
            lcol�շѱ�ż�.Add lstr�շѱ��
            
            Dim lcol�շ�ȷ����Ϣ As Collection
            
            Set lcol�շ�ȷ����Ϣ = New Collection
            With lcol�շ�ȷ����Ϣ
                .Add Val(ctxtInput(���۱���).Text), "���۱���"
                .Add mint���ѷ�ʽ���, "�շѷ�ʽ"
                .Add CDate(ctxtInput(��������).Text), "��������"
                .Add um�û����, "�շ���"
                
                '2006-5-15
                If ccmb������.Visible Then
                    .Add ccmb������.Text, "��������"
                Else
                    .Add "", "��������"
                End If
                .Add ccmb��Ӧҵ��.Text, "��Ӧҵ��"
                
            End With
            
            Call pobj�շѹ���.sub�շ�ȷ��(lcol�շѱ�ż�, lcol�շ�ȷ����Ϣ)
            
            mcur�ܽ�� = 0
            ctxtInput(Ӧ�ս��) = "0"
            
            sub�������
                        
            '��ӡƱ�ݡ�
            'Call func¼��Ʊ�ݺ�
            
            pobj�շѹ���.sub��ӡƱ�� lstr�շѱ��, IIf(cchkԤ��.Value = 1, True, False), True
            
            ctxtƱ�ݺ� = func��ȡƱ�ݺ�()
            If ctxtƱ�ݺ� <> "" And mbln���ƺŶ� Then
                If CLng(ctxtƱ�ݺ�) < 100 Then MsgBox "Ʊ�ݺű���λ�������Ƿ���ȷ��", vbInformation, "ϵͳ��ʾ"
            End If
            ctxtInput(�շ�_���ѵ�λ).SetFocus
            
        Case "ɾ��"
            Dim lcurMoney As Currency
            
            If cgrdDetail.Row > 0 Then
                cgrdDetail.RemoveItem cgrdDetail.Row
                For i = 1 To cgrdDetail.Rows - 1
                    lcurMoney = lcurMoney + cgrdDetail.ValueMatrix(i, �����嵥_���)
                Next
                mcur�ܽ�� = lcurMoney
                
                ctxtInput(Ӧ�ս��) = lcurMoney * Val(ctxtInput(���۱���).Text)
                ctxtInput(Ӧ�ս���д) = FuncConvertToCapsStr(Val(ctxtInput(Ӧ�ս��)))
            End If
            
        Case "���"
            sub�������
            
        Case Else
    End Select
    Exit Sub
    
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", "mobj����ͨ�ö���_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Function func�����Ŀ�Ƿ���ѡ(ByVal para�շ���Ŀ��� As String) As Boolean
On Error GoTo errhandle
    Dim i As Long
    func�����Ŀ�Ƿ���ѡ = False
    If cgrdDetail.Rows = 1 Then
        func�����Ŀ�Ƿ���ѡ = False
        Exit Function
    End If
    
    For i = 1 To cgrdDetail.Rows - 1
        If para�շ���Ŀ��� = cgrdDetail.TextMatrix(i, �����嵥_�շ���Ŀ���) Then
            func�����Ŀ�Ƿ���ѡ = True
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", " func�����Ŀ�Ƿ���ѡ()", Err.Number, Err.Description
End Function

Private Function ValidateData() As Boolean
    On Error GoTo errhandle
    ValidateData = False
    If ctxtInput(�շ�_������).Text = vbNullString And ctxtInput(�շ�_���ѵ�λ) = vbNullString Then
        sffuncMsg """������"" �� ""���ѵ�λ"" ������������֮һ��", sf����
        Exit Function
    End If
    If cgrdDetail.Rows = 1 Then
        sffuncMsg "�޷�����Ϣ���Ա��棡", sf����
        Exit Function
    End If
    
    If IsNumeric(ctxtInput(19).Text) Then
        If CDbl(ctxtInput(19).Text) = 0 Then
            sffuncMsg "���۱��ʲ���Ϊ0��", sf����
            ctxtInput(19).Text = "1.00"
            Exit Function
        End If
    Else
        sffuncMsg "���۱���¼�벻��ȷ��", sf����
        Exit Function
    End If
    ValidateData = True
    Exit Function
errhandle:
    sfsub������ "�շѽ��沿��", "frmֱ���շ�", " ValidateData()", Err.Number, Err.Description, True
End Function


