VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "¼��ؼ�.ocx"
Begin VB.Form frm�˷� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Cchk�˷Ѵ�ӡ��ʶ 
      Caption         =   "�˷�ʱ��ӡƱ��"
      Height          =   255
      Left            =   9000
      TabIndex        =   31
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox ctxt��ʾ 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   60
      TabIndex        =   29
      Text            =   "���Ժ�..."
      Top             =   7515
      Visible         =   0   'False
      Width           =   900
   End
   Begin MSComctlLib.ProgressBar cprg���� 
      Height          =   120
      Left            =   960
      TabIndex        =   28
      Top             =   7575
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Max             =   50
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   4875
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      TabIndex        =   13
      Top             =   6915
      Width           =   10995
      Begin VB.Timer ctmr��ʱ 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   165
         Top             =   240
      End
      Begin VB.TextBox ctxt�˷��� 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6420
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox ctxt�˷����� 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   270
         Width           =   1905
      End
      Begin VB.TextBox ctxt�ܽ�� 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         Height          =   330
         Left            =   4125
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   285
         Width           =   1335
      End
      Begin VB.TextBox ctxt�˷����� 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   1365
      End
      Begin ¼��ؼ�.ctlInputBox cinb�˷����� 
         Height          =   360
         Index           =   2
         Left            =   5805
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   580
         Text            =   ""
         Label           =   "�˷���"
         Enabled         =   0   'False
         ����            =   ""
         ����            =   0
         ����������ֵ  =   0   'False
         ���������Сֵ  =   0   'False
         �����ѡ        =   0   'False
      End
      Begin MSComCtl2.DTPicker cdtp���� 
         Height          =   345
         Index           =   2
         Left            =   3300
         TabIndex        =   9
         Top             =   735
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   36951
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "�˷���"
         Height          =   210
         Left            =   5640
         TabIndex        =   25
         Top             =   375
         Width           =   600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "�շ�����"
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "�ܽ��"
         Height          =   180
         Left            =   3330
         TabIndex        =   21
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˷�����"
         Height          =   240
         Left            =   8145
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar cstu״̬�� 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   7380
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "�˷�"
            TextSave        =   "�˷�"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   11055
      Begin VSFlex6Ctl.vsFlexGrid cgrd������Ϣ 
         Height          =   4245
         Left            =   75
         TabIndex        =   10
         Top             =   540
         Width           =   10815
         _cx             =   23743108
         _cy             =   23731520
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
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
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648447
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   "�շ�����        |�շѱ��       |�շ���Ŀ     |����   |���    |������     |���ѵ�λ               |�վݺ� |���۱���   "
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         DataMember      =   ""
         Begin VB.TextBox ctxtͬһ���� 
            Appearance      =   0  'Flat
            Height          =   1650
            Left            =   0
            TabIndex        =   30
            Top             =   255
            Visible         =   0   'False
            Width           =   1710
         End
      End
      Begin VB.Label clab�˷Ѽ�¼�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   9900
         TabIndex        =   38
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˷Ѽ�¼����"
         Height          =   180
         Left            =   8760
         TabIndex        =   37
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label clab��¼�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   8200
         TabIndex        =   36
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ü�¼����"
         Height          =   180
         Left            =   7080
         TabIndex        =   35
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˷���Ϣ"
         Height          =   180
         Left            =   5640
         TabIndex        =   34
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Height          =   135
         Left            =   5280
         TabIndex        =   33
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "��ѯҪ�˷ѵķ���"
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   15
      TabIndex        =   6
      Top             =   720
      Width           =   10995
      Begin VB.CheckBox cchk�Ƿ�ͨ����λ�ӿڲ�ѯ 
         Height          =   225
         Left            =   10380
         TabIndex        =   27
         Top             =   855
         Width           =   210
      End
      Begin VB.CheckBox cchk��ʱ���ѯ 
         Height          =   285
         Left            =   6600
         TabIndex        =   0
         Top             =   810
         Value           =   1  'Checked
         Width           =   225
      End
      Begin ¼��ؼ�.ctlInputBox cinb�˷����� 
         Height          =   360
         Index           =   1
         Left            =   5520
         TabIndex        =   2
         Top             =   360
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   850
         Text            =   ""
         Label           =   "������(&F)"
         ����            =   ""
         ����            =   0
         ����������ֵ  =   0   'False
         ���������Сֵ  =   0   'False
         �����ѡ        =   0   'False
      End
      Begin ¼��ؼ�.ctlInputBox cinb�˷����� 
         Height          =   360
         Index           =   0
         Left            =   8040
         TabIndex        =   3
         Top             =   360
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   1030
         Text            =   ""
         Label           =   "���ѵ�λ(&J)"
         ����            =   ""
         ����            =   0
         ����������ֵ  =   0   'False
         ���������Сֵ  =   0   'False
         �����ѡ        =   0   'False
      End
      Begin ¼��ؼ�.ctlInputBox cinb�˷����� 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   1030
         Text            =   ""
         Label           =   "�շ�����(&N)"
         ����            =   ""
         ����            =   0
         ����������ֵ  =   0   'False
         ���������Сֵ  =   0   'False
         �����ѡ        =   0   'False
      End
      Begin MSComCtl2.DTPicker cdtp���� 
         Height          =   300
         Index           =   0
         Left            =   3675
         TabIndex        =   5
         Top             =   810
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   36951
      End
      Begin MSComCtl2.DTPicker cdtp���� 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   810
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   36951
      End
      Begin ¼��ؼ�.ctlInputBox cinb�˷����� 
         Height          =   360
         Index           =   4
         Left            =   2880
         TabIndex        =   32
         Top             =   360
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LeftOfTextbox   =   850
         Text            =   ""
         Label           =   "�վݺ�(&S)"
         Length          =   8
         ����            =   ""
         ����            =   0
         ����������ֵ  =   0   'False
         ���������Сֵ  =   0   'False
         �����ѡ        =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "ͨ����λ�����ӿڲ�ѯ��λ"
         Height          =   195
         Left            =   8055
         TabIndex        =   26
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "��ʱ���ѯ"
         Height          =   240
         Left            =   5520
         TabIndex        =   18
         Top             =   855
         Width           =   990
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "���ڷ�Χ(B)"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   885
         Width           =   990
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "��(E)"
         Height          =   180
         Left            =   2880
         TabIndex        =   11
         Top             =   870
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker cdtp���� 
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   529
      _Version        =   393216
      Format          =   20185089
      CurrentDate     =   36951
   End
End
Attribute VB_Name = "frm�˷�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************frm�˷�***********************************************************************************
'����ʱ�䣺                 2001-3-29
'�����ˣ�                   ����
'�޸�ʱ�䣺
'�޸��ˣ�
'***************************BEGIN*****************************************************************************************
Option Explicit
Public pblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mstr��λ��� As String

'Private pobj�շѹ��� As Object
'Private pobjҵ������ As Object
'Private pobj��λ��λ As Object  '��λ�����ӿ�

Private rs���Ҽ�¼ As ADODB.Recordset

Private mstrSQL As String  '�����ַ���
'���峣��
Private Const �շ����� = 3
Private Const ������ = 1
Private Const ���ѵ�λ = 0
Private Const �˷��� = 2
Private Const �վݺ� = 4

Private Const ��ʼ���� = 1
Private Const �������� = 0
Private Const �˷����� = 2
  

'���ܣ�ѡ���Ƿ�ͨ����λ��λ�ӿڽ��в�ѯ��λ����
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cchk�Ƿ�ͨ����λ�ӿڲ�ѯ_Click()
    cinb�˷�����(���ѵ�λ).Text = ""
End Sub



Private Sub cdtp����_Change(Index As Integer)
    On Error Resume Next
    cdtp����(�˷�����).Value = CDate(Now)
End Sub

'���ܣ����������Ϣ���е�һ�У�ˢ����ʾ��ѡ��ļ�¼
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cgrd������Ϣ_Click()
    
    Call subѡ��ͬ������
    Call ��̬����TextBox
End Sub

 



'���ܣ��ڷ�����Ϣ�������β��ְ���
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cgrd������Ϣ_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
    End If
End Sub

'���ܣ����ѵ�λ��������ݱ仯�����Ӧ�Ľ��ѵ�λ���Ҳ��Ӧ�ı�
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cinb�˷�����_Change(Index As Integer)
    On Error Resume Next
    If Index = ���ѵ�λ Then
        mstr��λ��� = ""
    End If
End Sub

'���ܣ�˫�� ���ѵ�λ����򣬵������ѵ�λ��λ�ӿڽ���
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cinb�˷�����_DblClick(Index As Integer)
    On Error Resume Next
    Dim lobj��λ��Ϣ As Recordset
    
    If Index = ���ѵ�λ Then
         If cchk�Ƿ�ͨ����λ�ӿڲ�ѯ Then
            Set lobj��λ��Ϣ = pobj��λ��λ.func��λ�򵥶�λ(Screen.Width / 2, Screen.Height / 2)
            If Not (lobj��λ��Ϣ Is Nothing) Then
                cinb�˷�����(���ѵ�λ).Text = lobj��λ��Ϣ.Fields("��λ����").Value
                mstr��λ��� = lobj��λ��Ϣ.Fields("������").Value
                Set lobj��λ��Ϣ = Nothing
            End If
         End If
    End If
 
End Sub

'���ܣ�������ý���,������λ�����ӿڽ��棬�Ի�ȡ��λ��Ϣ
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-4-12
Private Sub cinb�˷�����_GotFocus(Index As Integer)
    On Error Resume Next
    Dim lobj��λ��Ϣ As Recordset
    
    If Index = ���ѵ�λ Then
         If cchk�Ƿ�ͨ����λ�ӿڲ�ѯ Then
            Set lobj��λ��Ϣ = pobj��λ��λ.func��λ�򵥶�λ(Screen.Width / 2, Screen.Height / 2)
            If Not (lobj��λ��Ϣ Is Nothing) Then
                cinb�˷�����(���ѵ�λ).Text = lobj��λ��Ϣ.Fields("��λ����").Value
                mstr��λ��� = lobj��λ��Ϣ.Fields("������").Value
                Set lobj��λ��Ϣ = Nothing
            End If
         End If
    End If
    
End Sub

'���ܣ�������λ�����ӿڽ��棬�Ի�ȡ��λ��Ϣ
'���룺�û������κμ�
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cinb�˷�����_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    
    Dim lobj��λ��Ϣ As Recordset

    If Index = ���ѵ�λ Then
         If KeyAscii = 8 Then
             If cchk�Ƿ�ͨ����λ�ӿڲ�ѯ Then
                cinb�˷�����(���ѵ�λ).Text = ""
                Exit Sub
             End If
         End If
         If cchk�Ƿ�ͨ����λ�ӿڲ�ѯ = 0 Then Exit Sub
         Set lobj��λ��Ϣ = pobj��λ��λ.func��λ�򵥶�λ(Screen.Width / 2, Screen.Height / 2)
         If Not (lobj��λ��Ϣ Is Nothing) Then
             cinb�˷�����(���ѵ�λ).Text = lobj��λ��Ϣ.Fields("��λ����").Value
             mstr��λ��� = lobj��λ��Ϣ.Fields("������").Value
             KeyAscii = 0
             Set lobj��λ��Ϣ = Nothing
             
         End If
    Else
         If KeyAscii = 13 Then
                cgrd������Ϣ.Clear
                cgrd������Ϣ.Rows = 1
                'cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |���۱���"
                cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |�վݺ� |���۱���"
                Call func���Ҽ�¼(func��ѯ����)
                Call sub�����
         ElseIf KeyAscii = 39 Then
              KeyAscii = 0
         End If
    End If
    
End Sub


'���ܣ���״̬������ʾ������ʾ
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub ctlb������_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lsngW, lsngH, lsngSepW As Single
    
    On Error Resume Next
    lsngW = ctlb������.ButtonWidth
    lsngH = ctlb������.ButtonHeight
    lsngSepW = ctlb������.Buttons(2).Width
    
    With cstu״̬��
    If X <= lsngW And Y <= lsngH Then
       .Panels(1).Text = " ���Ҽ�¼"
    Else
       .Panels(1).Text = ""
    End If
    
    If X <= 2 * lsngW + lsngSepW And X > lsngW + lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "�������ϵ�����"
    End If
    
    If X <= 3 * lsngW + 2 * lsngSepW And X > 2 * lsngW + 2 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "���������ѡ�ķ��ý����˷�"
    End If
    
    If X <= 4 * lsngW + 3 * lsngSepW And X > 3 * lsngW + 3 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "��ӡƱ��"
    End If
    
    If X <= 5 * lsngW + 4 * lsngSepW And X > 4 * lsngW + 4 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "�ر��˷Ѵ���"
    End If
    
    End With
End Sub


'���ܣ���ʱ������ˢ����ʾ������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29

Private Sub ctmr��ʱ_Timer()
    On Error Resume Next
    If cprg����.Value < cprg����.Max Then
       cprg����.Value = cprg����.Value + 5
    Else
       ctmr��ʱ.Enabled = False
    End If
   Me.Refresh
End Sub


'���ܣ�����һЩ��ݼ�
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyB
            cdtp����(��ʼ����).SetFocus
        Case vbKeyE
            cdtp����(��������).SetFocus
                    
        End Select
    End If
End Sub


'���ܣ�װ�ش��壬��ʼ�����棬�����ֿؼ����Ը�ֵ
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub Form_Load()
    If pblnInUse Then Exit Sub
    Dim lcol��������ť As Collection
    
    On Error GoTo errhandler
    pblnInUse = True                              'ָʾ����������
    
'    Set pobj�շѹ��� = CreateObject("�շ�ҵ�����.cls�շѹ���")
'    Set pobjҵ������ = CreateObject("�շ�ҵ�����.clsҵ������")
'    Set pobj��λ��λ = CreateObject("��λ����ҵ��.ClsUnitInterface")
    
    ��̬����TextBox
    
    '��ʼ��������
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = ctlb������
    Set lcol��������ť = New Collection
    lcol��������ť.Add "��ѯ(&Q)105"
    lcol��������ť.Add "|"
    lcol��������ť.Add "���"
    lcol��������ť.Add "�˷�(&T)122"
    lcol��������ť.Add "��ӡ"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    mobjGUI.subInitialize lcol��������ť, ""
    Set lcol��������ť = Nothing
    
    '���ܣ���ȡ�Ƿ����˷ѵ�Ȩ�ޡ�ʱ�䣺2002/02/20 ���ߣ��켽��
    If umfuncУ���û�Ȩ��("�շѹ���_�˷�") Then
        ctlb������.Buttons(4).Enabled = True
        ctxt�˷���.Text = um�û���                    '��ȡ��ǰ�û���
        ctxt�˷�����.Text = Date
    Else
        ctlb������.Buttons(4).Enabled = False
    End If
    '���ܣ���ȡ�Ƿ��д�ӡƱ�ݵ�Ȩ�ޡ�ʱ�䣺2002/02/20 ���ߣ��켽��
    If umfuncУ���û�Ȩ��("�շѹ���_Ʊ�ݴ�ӡ") Then
        ctlb������.Buttons(5).Enabled = True
    Else
        ctlb������.Buttons(5).Enabled = False
    End If
    
    cdtp����(��ʼ����).Value = Date               '��ʼ����ʼ���������Ϊ��������
    cdtp����(��������).Value = Date               '��ʼ���������������Ϊ��������
    'cgrd������Ϣ.Cols = 7
    cgrd������Ϣ.Rows = 1
    ' ���˷�ʱ��ʶ
    Cchk�˷Ѵ�ӡ��ʶ.Value = 0
    Exit Sub
errhandler:
    Call sfsub������("�շѽ������", "frm�˷�", "Form_Load", Err.Number, Err.Description, False)
End Sub

'���ܣ���ʾ��ѯ����ڱ����
'���룺��
'�������
'���أ������ѯ���������ʾ�������ڱ���У�����������Ϊ�հ�
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Function func���Ҽ�¼(ByVal strSQL As String) As ADODB.Recordset
    On Error GoTo errhandle
    'Set func���Ҽ�¼ = pobj�շѹ���.func��ѯ������Ϣ(strSQL)
    Set func���Ҽ�¼ = dafuncGetData("select * from �շѹ���_��ӡ������Ϣ where " & strSQL)
    Exit Function
errhandle:
    Call sfsub������("�շѽ������", "frm�˷�", "func���Ҽ�¼", Err.Number, Err.Description, True)
End Function
'���ܣ��õ��û������ѯ����
'���룺��
'�������
'���أ��������Ĳ�ѯ������Ϊ�գ��򷵻غ���Where���Ĳ�ѯ�����ַ�������֮����""
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Function func��ѯ����() As String
On Error Resume Next
Dim strSQL As String

    strSQL = ""
    '�շ���������
    If Trim(cinb�˷�����(�շ�����).Text) <> "" Then
       strSQL = "'" & Trim(cinb�˷�����(�շ�����).Text) & "',"
    Else
        strSQL = "'',"
    End If
    
    '�޸ģ�2002-6-23����������ݺŲ�ѯ��
    If Trim(cinb�˷�����(�վݺ�).Text) <> "" Then
        strSQL = strSQL & "'" & Trim(cinb�˷�����(�վݺ�).Text) & "',"
    Else
        strSQL = strSQL & "'',"
    End If
    
    '����������
    If Trim(cinb�˷�����(������).Text) <> "" Then
        strSQL = strSQL & "'" & Trim(cinb�˷�����(������).Text) & "',"
    Else
        strSQL = strSQL & "'',"
    End If
    
    '���ѵ�λ
    If Trim(cinb�˷�����(���ѵ�λ).Text) <> "" Then
        strSQL = strSQL & "'" & Trim(cinb�˷�����(���ѵ�λ).Text) & "',"
    Else
        strSQL = strSQL & "'',"
    End If
    
     '�������ڷ�Χ
    If cchk��ʱ���ѯ.Value = 1 Then
    
        If Trim(cdtp����(��ʼ����)) <> "" Then
            strSQL = strSQL & "'" & Trim(cdtp����(��ʼ����)) & "',"
        Else
            strSQL = strSQL & "'',"
        End If
        
        If Trim(cdtp����(��������)) <> "" Then
            strSQL = strSQL & "'" & Trim(cdtp����(��������)) & "'"
        Else
            strSQL = strSQL & "''"
        End If
    Else
        strSQL = strSQL & "'',''"
    End If
    
    func��ѯ���� = strSQL

End Function
'���ܣ������������У��ṩ���շ�Ʊ�ݵĴ�ӡ����.
'ʱ��: 2002/02/20
'���ߣ��켽��
Private Sub sub��ӡƱ��()
On Error GoTo errhandle
    Dim lcol������Ϣ As Collection         '��������Ϣ�����ֶ���Ϣд�뼯����
    Dim lcol���ô�ӡ��Ϣ�� As Collection   '��ŷ�����Ϣ�ļ���
    Dim lrec���Ҽ�¼ As Object             '��Ų�ѯ���ķ�����Ϣ
    Dim lstr��ʽ�ļ��� As String           '��¼��ӡ��ʽ���ļ���
    Dim lrec��ʽ�ļ������� As Object           '��¼��ӡ��ʽ���ļ�������
    Dim lrec������Ϣ As Object             '��¼��ķ�����Ϣ
    Dim lrec����Ʊ����Ϣ As Object         ' ��¼��Ʊ���йص���Ϣ
    Dim i As Long                         'ѭ������
    Dim j As Long                         'ѭ������
    Dim k As Long                         'ѭ������
    Dim lstr������ As String               '��¼����������
    Dim lstr���ѵ�λ As String             ' ��¼���ѵ�λ����
    Dim lsge���۱��� As Single            '��¼���۱���
    Dim lobj���ܼ�¼ As Object
    
    '�ж��Ƿ�ѡ�м�¼
    With cgrd������Ϣ
    If .Row = 0 Then
        MsgBox "��ѡ��Ҫ��ӡ�ķ�����Ϣ��", vbInformation, "��ӡƱ��"
        Exit Sub
    End If
    '�����˳���ť������
    ctlb������.Buttons(7).Enabled = False
    
    
    '���²�ѯ���ݽӿ�
    'ʱ�䣺2002/08/05
    '���ߣ��켽��
    Dim lstr�洢���� As String
    mstrSQL = func��ѯ����
    lstr�洢���� = "exec �շѹ���_�����շ���Ϣ " + mstrSQL
    Set lrec���Ҽ�¼ = dafuncGetData(lstr�洢����)
    
    'Set lrec���Ҽ�¼ = func���Ҽ�¼(mstrSQL)
    
    '�ж��Ƿ�ѡ��ͬһ��������
    Dim lstrtemp As String
    lstrtemp = .TextMatrix(.Row, 0)
    For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                 lrec���Ҽ�¼.MoveFirst
                 lrec���Ҽ�¼.Move i - 1
                If lstrtemp <> lrec���Ҽ�¼("�շ�����") Then
                    MsgBox "����ѡ���˲�ͬ���շ����ţ�������ѡ��", vbOKOnly, "��ӡƱ��"
                    Call subѡ��ͬ������
                    Exit Sub
                End If
             End If
    Next i
    End With
    '��ȡ�������Ϣ��ص�Ʊ����Ϣ
    Set lrec����Ʊ����Ϣ = pobj�շѹ���.funcExecute("select b.Ʊ�����ͱ�� from �շѹ���_�շ���Ŀ�ֵ�� b, �շѹ���_������Ϣ�� c " & _
                                               "Where b.�շ���Ŀ��� = c.�շ���Ŀ��� and c.�շ����� ='" & _
                                               lrec���Ҽ�¼("�շ�����") & "' group by b.Ʊ�����ͱ��", "cls������Ϣ")
    'У���������Ϣ��ص�Ʊ����Ϣ
    If (lrec����Ʊ����Ϣ Is Nothing) Or (lrec����Ʊ����Ϣ.BOF And lrec����Ʊ����Ϣ.EOF) Then
        sffuncMsg "δ�������շ���Ŀ��Ʊ��������Ϣ,�޷����д�ӡ!", sf����
        Exit Sub
    Else
        lrec����Ʊ����Ϣ.MoveFirst
    End If
                                                                                                                             
    '��Ʊ������ȡ��������Ϣ
    For i = 0 To lrec����Ʊ����Ϣ.RecordCount - 1
        '��ȡ��ӡ������Ϣ
        Set lrec������Ϣ = pobj�շѹ���.funcExecute("select * from �շѹ���_��ӡ������Ϣ where Ʊ�����ͱ��=" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & " and �շ�����='" & lrec���Ҽ�¼("�շ�����") & "'", "cls������Ϣ")
        'У�������Ϣ
        If (lrec������Ϣ Is Nothing) Or (lrec������Ϣ.BOF And lrec������Ϣ.EOF) Then
            sffuncMsg "�޿ɴ�ӡ��Ϣ!", sf����
            Exit Sub
        End If
        '���������Ϣ�н����˺ͽ��ѵ�λΪ��ֵ�����
        If IIf(IsNull(lrec������Ϣ("���ѵ�λ����").Value), "", lrec������Ϣ("���ѵ�λ����")) <> "" Then
            lstr���ѵ�λ = lrec������Ϣ("���ѵ�λ����").Value
        Else
            lstr���ѵ�λ = ""
        End If
        If IIf(IsNull(lrec������Ϣ("������").Value), "", lrec������Ϣ("������")) <> "" Then
            lstr������ = lrec������Ϣ("������").Value
        Else
            lstr������ = ""
        End If
        '��ʼ�����۱���ֵ
        lsge���۱��� = 1
        Set lcol���ô�ӡ��Ϣ�� = New Collection
        
        '�޸ģ�2002-9-29������ϲ���ӡ��
        Set lobj���ܼ�¼ = pobj�շѹ���.funcExecute("select �շ���Ŀ���,����=avg(����),����=sum(����),���=sum(���) from �շѹ���_��ӡ������Ϣ " _
                            & "where Ʊ�����ͱ��=" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & " and �շ�����='" & lrec���Ҽ�¼("�շ�����") _
                            & "' group by �շ�����,�շ���Ŀ���", "cls������Ϣ")
        
        '��������Ϣ���뵽�϶�����
        For j = 0 To lobj���ܼ�¼.RecordCount - 1
            '�޸ģ�2002-9-29�������ȡ��ǰ��Ŀ����ϸ��Ϣ��
            Set lrec������Ϣ = pobj�շѹ���.funcExecute("select * from �շѹ���_��ӡ������Ϣ where Ʊ�����ͱ��=" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & " and �շ�����='" & lrec���Ҽ�¼("�շ�����") & "' and �շ���Ŀ���='" & lobj���ܼ�¼("�շ���Ŀ���") & "'", "cls������Ϣ")
            
            Set lcol������Ϣ = New Collection
            For k = 0 To lrec������Ϣ.Fields.Count - 1
                If lrec������Ϣ.Fields(k).Name = "���ѵ�λ����" Or lrec������Ϣ.Fields(k).Name = "������" Or lrec������Ϣ.Fields(k).Name = "���۱���" Then
                    If lrec������Ϣ.Fields(k).Name = "���ѵ�λ����" Then lcol������Ϣ.Add lstr���ѵ�λ, "���ѵ�λ����"
                    If lrec������Ϣ.Fields(k).Name = "������" Then lcol������Ϣ.Add lstr������, "������"
                    If lrec������Ϣ.Fields(k).Name = "���۱���" Then
                        lsge���۱��� = lrec������Ϣ(k).Value
                        lcol������Ϣ.Add lsge���۱���, "���۱���"
                    End If
                ElseIf lrec������Ϣ.Fields(k).Name <> "����" And lrec������Ϣ.Fields(k).Name <> "����" And lrec������Ϣ.Fields(k).Name <> "���" Then
                    '�޸ģ�2002-9-29��������ۡ������������ʾ�������ݡ�
                    lcol������Ϣ.Add lrec������Ϣ(k).Value, lrec������Ϣ.Fields(k).Name
                End If
            Next k
            '�޸ģ�2002-9-29��������ۡ������������ʾ�������ݡ�
            lcol������Ϣ.Add Format(lobj���ܼ�¼("����").Value, "0.00"), "����"
            lcol������Ϣ.Add lobj���ܼ�¼("����").Value, "����"
            lcol������Ϣ.Add Format(lobj���ܼ�¼("���").Value, "0.00"), "���"
            
            lcol������Ϣ.Add "����ֵ", "����"
            lcol������Ϣ.Add "�Ա�ֵ", "�Ա�"
            lcol������Ϣ.Add "סԺ��ֵ", "סԺ��"
            lcol������Ϣ.Add "����ֵ", "����"
            lcol������Ϣ.Add "2002", "��Ժ����"
            lcol������Ϣ.Add "2002", "��Ժ����"
            lcol������Ϣ.Add "��Ժ����Աֵ", "��Ժ����Ա"
            lcol������Ϣ.Add "����ҽ��ֵ", "����ҽ��"
            
            lcol���ô�ӡ��Ϣ��.Add lcol������Ϣ
            
            'If Not lrec������Ϣ.EOF Then lrec������Ϣ.MoveNext
            If Not lobj���ܼ�¼.EOF Then lobj���ܼ�¼.MoveNext
        Next j
        '��ȡ��ʽ�ļ���
        Set lrec��ʽ�ļ������� = pobj�շѹ���.funcExecute("select * from �շѹ���_Ʊ��������Ϣ�� where Ʊ�����ͱ��='" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & "' and ��Ӧҵ��='һ��'", "cls������Ϣ")
        If lrec��ʽ�ļ������� Is Nothing Then
            sffuncMsg "δ���ҵ�Ʊ�ݸ�ʽ�ļ�!", sf����
        End If
        If lrec��ʽ�ļ�������.BOF And lrec��ʽ�ļ�������.EOF Then
            sffuncMsg "δ���ҵ�Ʊ�ݸ�ʽ�ļ�!", sf����
        Else
            lstr��ʽ�ļ��� = lrec��ʽ�ļ�������("Ʊ�ݸ�ʽ�ļ�����")
            Call pobj�շѹ���.sub��ӡƱ��(lcol���ô�ӡ��Ϣ��, App.Path & "\" & lstr��ʽ�ļ���, , lsge���۱���, lrec��ʽ�ļ�������("�������").Value)
        End If
        '�жϼ�¼��
        If Not lrec����Ʊ����Ϣ.EOF Then lrec����Ʊ����Ϣ.MoveNext
    Next i
    '�����˳���ť����
    ctlb������.Buttons(7).Enabled = True
Exit Sub
errhandle:
    sfsub������ "�շѽ������", "frm�˷�", "sub��ӡƱ��", Err.Number, Err.Description, True
End Sub
'���ܣ��˷ѣ��ı��շ�״̬Ϊ��2:���˷ѡ�
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub sub�˷�()
On Error GoTo errhandle
    Dim mcol������Ϣ As Collection  '��������Ϣ�����ֶ���Ϣд�뼯���У�
    Dim lcol���ô�ӡ��Ϣ�� As Collection    '��ŷ�����Ϣ�ļ���
    Dim lcol������Ϣ As Collection         '��������Ϣ�����ֶ���Ϣд�뼯����
    Dim lstr��ʽ�ļ��� As String           '��¼��ӡ��ʽ���ļ���
    Dim lrec��ʽ�ļ������� As Object           '��¼��ӡ��ʽ���ļ�������
    Dim lrec������Ϣ As Object             '��¼��ķ�����Ϣ
    Dim lrec����Ʊ����Ϣ As Object         ' ��¼��Ʊ���йص���Ϣ
    Dim i As Long                         'ѭ������
    Dim j As Long                         'ѭ������
    Dim k As Long                         'ѭ������
    Dim lstr������ As String               '��¼����������
    Dim lstr���ѵ�λ As String             ' ��¼���ѵ�λ����
    Dim lsge���۱��� As Single            '��¼���۱���
    Dim lsge��� As Single                '��¼����
    Dim lbln�˷ѳɹ���� As Boolean        ' ��¼�˷ѳɹ�״̬
    
    Dim lobj���ܼ�¼ As Object
    
    lbln�˷ѳɹ���� = False
    
    '****************�˷���Ϣ����*****************
    With cgrd������Ϣ
    If .Row = 0 Then
        MsgBox "��ѡ��Ҫ�˷ѵķ�����Ϣ��", vbInformation, "�˷�"
        Exit Sub
    End If
    
    '���²�ѯ���ݽӿ�
    'ʱ�䣺2002/08/05
    '���ߣ��켽��
    Dim lstr�洢���� As String      '���������¼ִ�д洢�������
    mstrSQL = func��ѯ����
    lstr�洢���� = "exec �շѹ���_�����շ���Ϣ " + mstrSQL
    Set rs���Ҽ�¼ = dafuncGetData(lstr�洢����)
    
    'Set rs���Ҽ�¼ = func���Ҽ�¼(mstrSQL)
    
    '*****����Ϊ�����޸�-0902*****************
    Dim lstrtemp As String
    lstrtemp = .TextMatrix(.Row, 0)
    For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                 rs���Ҽ�¼.MoveFirst
                 rs���Ҽ�¼.Move i - 1
                If lstrtemp <> rs���Ҽ�¼("�շ�����") Then
                    MsgBox "����ѡ���˲�ͬ���շ����ţ�������ѡ��", vbOKOnly, "�˷�"
                    Call subѡ��ͬ������
                    Exit Sub
                End If
             End If
    Next i
    '*****����Ϊ�����޸�-0902*****************
    
    '���ܣ����ӶԷ�����Ϣ����֤
    'ʱ�䣺2002/08/05
    '���ߣ��켽��
    
    If rs���Ҽ�¼.RecordCount > 0 Then
        For i = 1 To .Rows - 1
            If funУ�����Ϣ(.TextMatrix(.RowSel, 0)) = False Then
                sffuncMsg "�˷�����Ϣ���˷ѣ�"
                Exit Sub
            End If
            
        Next
    End If
    
    If MsgBox("���ѵ�λ     ��" & .TextMatrix(.Row, 6) & Chr(13) & Chr(10) & "������       ��" & .TextMatrix(.Row, 5) & Chr(13) & Chr(10) & "�շ�����     ��" & .TextMatrix(.Row, 0) & Chr(13) & Chr(10) & "�ܽ��       ��" & ctxt�ܽ�� & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "   ����Ҫ�˷���", vbYesNo, "�˷�") = vbNo Then Exit Sub
    dasubBeginTran      '��ʼ����
    For i = 1 To .Rows - 1
        If .IsSelected(i) = True Then
             rs���Ҽ�¼.MoveFirst
             rs���Ҽ�¼.Move i - 1
             Set mcol������Ϣ = New Collection
             '��������ֶ�ֵԭ�ⲻ�������������ݿ⣺�շѹ���_������Ϣ���ֶ����Ӧ
             mcol������Ϣ.Add rs���Ҽ�¼("�շ�����"), "�շ�����"
             mcol������Ϣ.Add rs���Ҽ�¼("�շѱ��"), "�շѱ��"
             mcol������Ϣ.Add rs���Ҽ�¼("�շ���Ŀ���"), "�շ���Ŀ���"
             mcol������Ϣ.Add rs���Ҽ�¼("����"), "����"
             mcol������Ϣ.Add rs���Ҽ�¼("����"), "����"
             mcol������Ϣ.Add rs���Ҽ�¼("���"), "���"
             mcol������Ϣ.Add rs���Ҽ�¼("������"), "������"
             mcol������Ϣ.Add IIf(IsNull(rs���Ҽ�¼("���ѵ�λ���")), "", rs���Ҽ�¼("���ѵ�λ���")), "���ѵ�λ���"
             mcol������Ϣ.Add rs���Ҽ�¼("��������"), "��������"
             mcol������Ϣ.Add rs���Ҽ�¼("�շ���"), "�շ���"
             mcol������Ϣ.Add rs���Ҽ�¼("���ܿ��Ҿ�����"), "���ܿ��Ҿ�����"
             mcol������Ϣ.Add rs���Ҽ�¼("���ܿ��ұ��"), "���ܿ��ұ��"
             mcol������Ϣ.Add rs���Ҽ�¼("���۱���"), "���۱���"
             mcol������Ϣ.Add rs���Ҽ�¼("���ѷ�ʽ"), "���ѷ�ʽ"
             mcol������Ϣ.Add IIf(IsNull(rs���Ҽ�¼("���ѵ�λ����")), "", rs���Ҽ�¼("���ѵ�λ����")), "���ѵ�λ����"
             '���������ֶ���Ҫ�޸ĵ�����
             mcol������Ϣ.Add "2", "�շ�״̬"
             mcol������Ϣ.Add um�û����, "�˷���"
             mcol������Ϣ.Add func��ȡ����������, "�˷�����"
             mcol������Ϣ.Add IIf(IsNull(rs���Ҽ�¼("�վݺ�")), "", rs���Ҽ�¼("�վݺ�")), "�վݺ�"
             '֪ͨҵ����޸ķ�����Ϣ
             pobj�շѹ���.func�޸ķ�����Ϣ mcol������Ϣ
         End If
    Next i
    '���˷ѳɹ����������������˷ѵ�����
     dasubCommitTran
     i = 1
     Do While i <= .Rows - 1
        DoEvents
        If .IsSelected(i) Then
RemoveLine:            .RemoveItem i
            If i <= .Rows - 1 Then
                If .IsSelected(i) Then GoTo RemoveLine
            End If
        End If
        i = i + 1
     Loop
    MsgBox "�˷��ѳɹ���", vbInformation, "�˷�"
    lbln�˷ѳɹ���� = True
    Call subѡ��ͬ������
    Call ��̬����TextBox
    Set mcol������Ϣ = Nothing
    End With
   
    '���ܣ����Ӷ��˷���Ϣ�Ĵ�ӡ���ܡ�
    'ʱ�䣺2002/02/20
    '���ߣ��켽��
    If lbln�˷ѳɹ���� = True Then
         If Cchk�˷Ѵ�ӡ��ʶ.Value = 0 Then Exit Sub
         
         '�޸ģ�2002-9-29������˷�Ʊ�ݴ�ӡ�����������Ա���Բ���
         sub��ӡ�˷�Ʊ�� rs���Ҽ�¼("�շ�����")
    End If
    
    Exit Sub
errhandle:
    dasubRollBack
    Call sfsub������("�շѽ������", "frm�˷�", "sub�˷�", Err.Number, Err.Description, True)
    Exit Sub
    Resume
End Sub

'���ܣ����Ӷ��˷���Ϣ�Ĵ�ӡ���ܡ�
'ʱ�䣺2002/02/20
'���ߣ��켽��
'�޸ģ�2002-9-29��������˷�Ʊ�ݴ�ӡ����������
Private Sub sub��ӡ�˷�Ʊ��(ByVal para�շ����� As String)
    Dim lcol���ô�ӡ��Ϣ�� As Collection    '��ŷ�����Ϣ�ļ���
    Dim lcol������Ϣ As Collection         '��������Ϣ�����ֶ���Ϣд�뼯����
    Dim lstr��ʽ�ļ��� As String           '��¼��ӡ��ʽ���ļ���
    Dim lrec��ʽ�ļ������� As Object           '��¼��ӡ��ʽ���ļ�������
    Dim lrec������Ϣ As Object             '��¼��ķ�����Ϣ
    Dim lrec����Ʊ����Ϣ As Object         ' ��¼��Ʊ���йص���Ϣ
    Dim i As Long                         'ѭ������
    Dim j As Long                         'ѭ������
    Dim k As Long                         'ѭ������
    Dim lstr������ As String               '��¼����������
    Dim lstr���ѵ�λ As String             ' ��¼���ѵ�λ����
    Dim lsge���۱��� As Single            '��¼���۱���
    Dim lsge��� As Single                '��¼����
    
    Dim lobj���ܼ�¼ As Object
    
    On Error GoTo errHanler
    
    '���ô�ӡ��ť������
     ctlb������.Buttons(7).Enabled = False
    '��ȡ�������Ϣ��ص�Ʊ����Ϣ
    Set lrec����Ʊ����Ϣ = pobj�շѹ���.funcExecute("select b.Ʊ�����ͱ�� from �շѹ���_�շ���Ŀ�ֵ�� b, �շѹ���_������Ϣ�� c " & _
                                               "Where b.�շ���Ŀ��� = c.�շ���Ŀ��� and c.�շ����� ='" & _
                                               para�շ����� & "' group by b.Ʊ�����ͱ��", "cls������Ϣ")
    'У���������Ϣ��ص�Ʊ����Ϣ
    If (lrec����Ʊ����Ϣ Is Nothing) Or (lrec����Ʊ����Ϣ.BOF And lrec����Ʊ����Ϣ.EOF) Then
        sffuncMsg "δ�������շ���Ŀ��Ʊ��������Ϣ,�޷����д�ӡ��", sf����
        Exit Sub
    Else
        lrec����Ʊ����Ϣ.MoveFirst
    End If
    
    '��Ʊ������ȡ��������Ϣ
    For i = 0 To lrec����Ʊ����Ϣ.RecordCount - 1
        '��ȡ��ӡ������Ϣ
        Set lrec������Ϣ = pobj�շѹ���.funcExecute("select * from �շѹ���_��ӡ������Ϣ where Ʊ�����ͱ��=" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & " and �շ�����='" & para�շ����� & "'", "cls������Ϣ")
        'У�������Ϣ
        If (lrec������Ϣ Is Nothing) Or (lrec������Ϣ.BOF And lrec������Ϣ.EOF) Then
            sffuncMsg "�޿ɴ�ӡ��Ϣ��", sf����
            Exit Sub
        End If
        '���������Ϣ�н����˺ͽ��ѵ�λΪ��ֵ�����
        If IIf(IsNull(lrec������Ϣ("���ѵ�λ����").Value), "", lrec������Ϣ("���ѵ�λ����")) <> "" Then
            lstr���ѵ�λ = lrec������Ϣ("���ѵ�λ����").Value
        Else
            lstr���ѵ�λ = ""
        End If
        If IIf(IsNull(lrec������Ϣ("������").Value), "", lrec������Ϣ("������")) <> "" Then
            lstr������ = lrec������Ϣ("������").Value
        Else
            lstr������ = ""
        End If
        '��ʼ�����۱���ֵ
        lsge���۱��� = 1
        Set lcol���ô�ӡ��Ϣ�� = New Collection
        
        '�޸ģ�2002-9-29������ϲ���ӡ��
        Set lobj���ܼ�¼ = pobj�շѹ���.funcExecute("select �շ���Ŀ���,����=avg(����),����=sum(����),���=sum(���) from �շѹ���_��ӡ������Ϣ " _
                        & "where Ʊ�����ͱ��=" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & " and �շ�����='" & para�շ����� _
                        & "' group by �շ�����,�շ���Ŀ���", "cls������Ϣ")
        
        '��������Ϣ���뵽�϶�����
        For j = 0 To lobj���ܼ�¼.RecordCount - 1
            '�޸ģ�2002-9-29�������ȡ��ǰ��Ŀ����ϸ��Ϣ��
            Set lrec������Ϣ = pobj�շѹ���.funcExecute("select * from �շѹ���_��ӡ������Ϣ where Ʊ�����ͱ��=" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & " and �շ�����='" & para�շ����� & "' and �շ���Ŀ���='" & lobj���ܼ�¼("�շ���Ŀ���") & "'", "cls������Ϣ")
            
            Set lcol������Ϣ = New Collection
            For k = 0 To lrec������Ϣ.Fields.Count - 1
                If lrec������Ϣ.Fields(k).Name = "���ѵ�λ����" Or lrec������Ϣ.Fields(k).Name = "������" Or lrec������Ϣ.Fields(k).Name = "���۱���" Or lrec������Ϣ.Fields(k).Name = "���" Then
                    If lrec������Ϣ.Fields(k).Name = "���ѵ�λ����" Then lcol������Ϣ.Add lstr���ѵ�λ, "���ѵ�λ����"
                    If lrec������Ϣ.Fields(k).Name = "������" Then lcol������Ϣ.Add lstr������, "������"
                    If lrec������Ϣ.Fields(k).Name = "���۱���" Then
                        lsge���۱��� = lrec������Ϣ(k).Value
                        lcol������Ϣ.Add lsge���۱���, "���۱���"
                    End If
'                        If lrec������Ϣ.Fields(k).Name = "���" Then
'                            lsge��� = 0 - lrec������Ϣ(k).Value
'                            lcol������Ϣ.Add lsge���, "���"
'                        End If
                ElseIf lrec������Ϣ.Fields(k).Name <> "����" And lrec������Ϣ.Fields(k).Name <> "����" And lrec������Ϣ.Fields(k).Name <> "���" Then
                    '�޸ģ�2002-9-29��������ۡ������������ʾ�������ݡ�
                    lcol������Ϣ.Add lrec������Ϣ(k).Value, lrec������Ϣ.Fields(k).Name
                End If
            Next k
            
            '�޸ģ�2002-9-29��������ۡ������������ʾ�������ݡ�
            lcol������Ϣ.Add Format(lobj���ܼ�¼("����").Value, "0.00"), "����"
            lcol������Ϣ.Add lobj���ܼ�¼("����").Value, "����"
            lcol������Ϣ.Add Format(0 - lobj���ܼ�¼("���").Value, "0.00"), "���"
            
            lcol������Ϣ.Add "����ֵ", "����"
            lcol������Ϣ.Add "�Ա�ֵ", "�Ա�"
            lcol������Ϣ.Add "סԺ��ֵ", "סԺ��"
            lcol������Ϣ.Add "����ֵ", "����"
            lcol������Ϣ.Add "2002", "��Ժ����"
            lcol������Ϣ.Add "2002", "��Ժ����"
            lcol������Ϣ.Add "��Ժ����Աֵ", "��Ժ����Ա"
            lcol������Ϣ.Add "����ҽ��ֵ", "����ҽ��"
            
            lcol���ô�ӡ��Ϣ��.Add lcol������Ϣ
            'If Not lrec������Ϣ.EOF Then lrec������Ϣ.MoveNext
            If Not lobj���ܼ�¼.EOF Then lobj���ܼ�¼.MoveNext
        Next j
        '��ȡ��ʽ�ļ���
        Set lrec��ʽ�ļ������� = pobj�շѹ���.funcExecute("select * from �շѹ���_Ʊ��������Ϣ�� where Ʊ�����ͱ��='" & lrec����Ʊ����Ϣ("Ʊ�����ͱ��") & "' and ��Ӧҵ��='һ��'", "cls������Ϣ")
        If lrec��ʽ�ļ������� Is Nothing Then
            sffuncMsg "δ���ҵ�Ʊ�ݸ�ʽ�ļ���", sf����
        End If
        If lrec��ʽ�ļ�������.BOF And lrec��ʽ�ļ�������.EOF Then
            sffuncMsg "δ���ҵ�Ʊ�ݸ�ʽ�ļ���", sf����
        Else
            lstr��ʽ�ļ��� = lrec��ʽ�ļ�������("Ʊ�ݸ�ʽ�ļ�����")
            
            '�޸ģ�2002-6-25��������Ӳ���para�˷ѡ�
            Call pobj�շѹ���.sub��ӡƱ��(lcol���ô�ӡ��Ϣ��, App.Path & "\" & lstr��ʽ�ļ���, , lsge���۱���, lrec��ʽ�ļ�������("�������").Value, True)
        End If
        '�жϼ�¼��
        If Not lrec����Ʊ����Ϣ.EOF Then lrec����Ʊ����Ϣ.MoveNext
    Next i

    '���ô�ӡ��ť����
    ctlb������.Buttons(7).Enabled = True

    Exit Sub
errHanler:
    Call sfsub������("�շѽ������", "frm�˷�", "sub��ӡ�˷�Ʊ��", Err.Number, Err.Description, True)
    ctlb������.Buttons(7).Enabled = True
    Exit Sub
    Resume
End Sub
'���ܣ���ȡ������ϵͳ��ǰʱ��ֵ
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cstu״̬��.Panels(1).Text = "�˷�"
End Sub

'���ܣ��رմ��壬����pblninuseΪFalse
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
'    Set pobj�շѹ��� = Nothing
'    Set pobjҵ������ = Nothing
'    Set pobj��λ��λ = Nothing
    Set mobjGUI = Nothing
    Set rs���Ҽ�¼ = Nothing
End Sub
'���ܣ����ݲ�ѯ�����������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub sub�����1()
 
Dim i As Integer
    On Error GoTo errhandle
    cprg����.Visible = True
    cprg����.Value = 15
    
    
    
    'mstrSQL = func��ѯ����
    'Set rs���Ҽ�¼ = func���Ҽ�¼(mstrSQL)
    
    '���²�ѯ���ݽӿ�
    'ʱ�䣺2002/08/05
    '���ߣ��켽��
    Dim lstr�洢���� As String      '���������¼ִ�д洢�������
    mstrSQL = func��ѯ����
    lstr�洢���� = "exec �շѹ���_�����շ���Ϣ " + mstrSQL
    Set rs���Ҽ�¼ = dafuncGetData(lstr�洢����)
    
    
    If rs���Ҽ�¼.RecordCount <= 0 Then
        MsgBox "δ�ҵ�ƥ���¼�����ܸ���Ŀδ����" & Chr(13) & Chr(10) & "�����������Ƿ���ȷ��", vbInformation, "�˷�"
        ctmr��ʱ.Enabled = False
        ctxt��ʾ.Visible = False
        cprg����.Value = 0
        cprg����.Visible = False
        Exit Sub
    End If
    cgrd������Ϣ.Clear       '��ձ��
    cgrd������Ϣ.Cols = 7    'ֻ��ʾ7��
    'cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |���۱���"
    cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |�վݺ� |���۱���"
    'rs���Ҽ�¼.MoveLast      'Ϊ��ȷ���õ���ȷ��.RecordCount���
    cgrd������Ϣ.Rows = rs���Ҽ�¼.RecordCount + 1
    rs���Ҽ�¼.MoveFirst
    i = 1
    Do Until rs���Ҽ�¼.EOF
        '��������Ϣ���
        cgrd������Ϣ.TextMatrix(i, 0) = IIf(IsNull(rs���Ҽ�¼.Fields("�շ�����")), "", rs���Ҽ�¼.Fields("�շ�����"))
        cgrd������Ϣ.TextMatrix(i, 1) = IIf(IsNull(rs���Ҽ�¼.Fields("�շѱ��")), "", rs���Ҽ�¼.Fields("�շѱ��"))
        'cgrd������Ϣ.TextMatrix(i, 2) = rs���Ҽ�¼.Fields("�շ���Ŀ���")
        cgrd������Ϣ.TextMatrix(i, 3) = IIf(IsNull(rs���Ҽ�¼.Fields("����")), "0", rs���Ҽ�¼.Fields("����"))
        cgrd������Ϣ.TextMatrix(i, 4) = IIf(IsNull(rs���Ҽ�¼.Fields("���")), "0.00", rs���Ҽ�¼.Fields("���"))
        cgrd������Ϣ.TextMatrix(i, 5) = IIf(IsNull(rs���Ҽ�¼.Fields("������")), "", rs���Ҽ�¼.Fields("������"))
        cgrd������Ϣ.TextMatrix(i, 7) = IIf(IsNull(rs���Ҽ�¼.Fields("���۱���")), "1", rs���Ҽ�¼.Fields("���۱���"))
        'ת����λ���Ϊ����
        Dim lstrNum As String
        Dim lstrUnitName As String
        Dim lrdsTemp As ADODB.Recordset
        lstrNum = IIf(IsNull(rs���Ҽ�¼.Fields("���ѵ�λ���")), "", rs���Ҽ�¼.Fields("���ѵ�λ���"))
        lstrUnitName = IIf(IsNull(rs���Ҽ�¼.Fields("���ѵ�λ����")), "", rs���Ҽ�¼.Fields("���ѵ�λ����"))
        If lstrUnitName = vbNullString Then
            If lstrNum <> vbNullString Then
                Set lrdsTemp = pobj�շѹ���.funcExecute("select ��λ���� from ��λ����_��λ������Ϣ�� where upper(������)=upper('" & lstrNum & "')", "cls������Ϣ")
                If Not (lrdsTemp Is Nothing) Then
                    If lrdsTemp.RecordCount = 1 Then
                        cgrd������Ϣ.TextMatrix(i, 6) = lrdsTemp("��λ����")
                    Else
                        cgrd������Ϣ.TextMatrix(i, 6) = lstrNum
                    End If
                Else
                    cgrd������Ϣ.TextMatrix(i, 6) = lstrNum
                End If
                Set lrdsTemp = Nothing
            Else
                cgrd������Ϣ.TextMatrix(i, 6) = lstrNum
            End If
        Else
            cgrd������Ϣ.TextMatrix(i, 6) = lstrUnitName
        End If
        
        'ת���շ���Ŀ���Ϊ����
        lstrNum = rs���Ҽ�¼.Fields("�շ���Ŀ���")
        If lstrNum <> vbNullString Then
            Set lrdsTemp = pobj�շѹ���.func��ѯ�շ���Ŀ("upper(�շ���Ŀ���)=upper('" & lstrNum & "')")
            If Not (lrdsTemp Is Nothing) Then
                If lrdsTemp.RecordCount = 1 Then
                    cgrd������Ϣ.TextMatrix(i, 2) = lrdsTemp("�շ���Ŀ����")
                Else
                    cgrd������Ϣ.TextMatrix(i, 2) = lstrNum
                End If
            Else
                cgrd������Ϣ.TextMatrix(i, 2) = lstrNum
            End If
                Set lrdsTemp = Nothing
        Else
            cgrd������Ϣ.TextMatrix(i, 2) = lstrNum
        End If
        
        rs���Ҽ�¼.MoveNext
        i = i + 1
    Loop
    Call subѡ��ͬ������
    Call ��̬����TextBox
    
    ctmr��ʱ.Enabled = False
    cprg����.Value = cprg����.Max
    cprg����.Visible = False
    ctxt��ʾ.Visible = False
    
    Exit Sub
errhandle:
        ctmr��ʱ.Enabled = False
        ctxt��ʾ.Visible = False
        cprg����.Value = 0
        cprg����.Visible = False
        Call sfsub������("�շѽ������", "frm�˷�", "sub�˷�", Err.Number, Err.Description, True)
End Sub

'���ܣ��޸ķ����ڽ����ϵ���ʾ
'ʱ�䣺2002/08/01
'���ߣ��켽��
Private Sub sub�����()
Dim i As Integer
    On Error GoTo errhandle
    Dim lint�˷Ѽ�¼�� As Long
    '*********************
    '�ж�������
    If Trim(cinb�˷�����(�շ�����).Text) = "" And Trim(cinb�˷�����(�վݺ�).Text) = "" And Trim(cinb�˷�����(������).Text) = "" And Trim(cinb�˷�����(���ѵ�λ).Text) = "" Then
        If cchk��ʱ���ѯ.Value = 0 Then
            MsgBox "��������һ��������ָ��ʱ�䷶Χ��", vbInformation, "�˷�"
            Exit Sub
        Else
            If cdtp����(��ʼ����).Value > cdtp����(��������).Value Then
                MsgBox "��ʼ���ڲ��ܴ��ڽ������ڡ�", vbInformation, "�˷�"
                Exit Sub
            ElseIf DateDiff("d", cdtp����(��ʼ����).Value, cdtp����(��������).Value) > 90 Then
                MsgBox "���ڷ�Χ���ܴ���90�졣", vbInformation, "�˷�"
                Exit Sub
            End If
        End If
    Else
        If cdtp����(��ʼ����).Value > cdtp����(��������).Value Then
            MsgBox "��ʼ���ڲ��ܴ��ڽ������ڡ�", vbInformation, "�˷�"
            Exit Sub
        End If
    End If
    '*********************
    ctxt��ʾ.Visible = True
    ctmr��ʱ.Enabled = True
    ctxt��ʾ.Text = "���Ժ�..."
    cprg����.Visible = True
    cprg����.Value = 15
    Me.Refresh
    
    '���ܣ����´����ݿ��л�ȡ������Ϣ
    'ע�⣺���ô洢���̻�ȡ�����з��ô������˷ѷ��õĴ���
    'ʱ�䣺2002/08/02
    '���ߣ��켽��
    Dim lstr�洢���� As String
    mstrSQL = func��ѯ����
    lstr�洢���� = "exec �շѹ���_�����շ���Ϣ " + mstrSQL
    Set rs���Ҽ�¼ = dafuncGetData(lstr�洢����)
    
    Dim lobjRec As Object       '������ʱ��¼����
    Dim lInt As Long            '����ѭ������
    
    lstr�洢���� = "exec �շѹ���_���ط��ô��� " + mstrSQL
    Set lobjRec = dafuncGetData(lstr�洢����)
    
    If lobjRec.RecordCount > 0 Then
        lobjRec.MoveFirst
        For lInt = 0 To lobjRec.RecordCount - 1
            If lobjRec("��Ŀ") = "�ܴ���" Then
                clab��¼��.Caption = IIf(IsNull(lobjRec("����")), "0", lobjRec("����"))
            End If
                        
            If lobjRec("��Ŀ") = "�˷Ѵ���" Then
                clab�˷Ѽ�¼��.Caption = IIf(IsNull(lobjRec("����")), "0", lobjRec("����"))
            End If
                        
            lobjRec.MoveNext
        Next
    End If
    
    If rs���Ҽ�¼.RecordCount <= 0 Then
        ctxt��ʾ.Visible = False
        ctmr��ʱ.Enabled = False
        cprg����.Value = 0
        cprg����.Visible = False
        MsgBox "δ�ҵ�ƥ���¼�����ܸ���Ŀδ����" & Chr(13) & Chr(10) & "���������Ƿ���ȷ��", vbInformation, "�˷�"
        
        Exit Sub
    End If
    
   
    cgrd������Ϣ.Clear       '��ձ��
    cgrd������Ϣ.Cols = 10   'ֻ��ʾ9��
    cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |�վݺ� |���۱���|��������     "
    'rs���Ҽ�¼.MoveLast      'Ϊ��ȷ���õ���ȷ��.RecordCount���
    cgrd������Ϣ.Rows = rs���Ҽ�¼.RecordCount + 1
    rs���Ҽ�¼.MoveFirst
    i = 1
    Do Until rs���Ҽ�¼.EOF
        '��������Ϣ���
        
        If rs���Ҽ�¼("��ʶ") = 2 Then
            'cvfg���ּƻ�.Cell(flexcpBackColor, Row, 2, Row, cvfg���ּƻ�.Cols - 1) = M_CON_ʱ�䵽
            cgrd������Ϣ.Cell(flexcpBackColor, i, 0, i, 9) = &HFFC0C0
        Else
            
        End If
        
        cgrd������Ϣ.TextMatrix(i, 0) = IIf(IsNull(rs���Ҽ�¼.Fields("�շ�����")), "", rs���Ҽ�¼.Fields("�շ�����"))
        cgrd������Ϣ.TextMatrix(i, 1) = IIf(IsNull(rs���Ҽ�¼.Fields("�շѱ��")), "", rs���Ҽ�¼.Fields("�շѱ��"))
        cgrd������Ϣ.TextMatrix(i, 2) = IIf(IsNull(rs���Ҽ�¼.Fields("�շ���Ŀ����")), "", rs���Ҽ�¼.Fields("�շ���Ŀ����"))
        cgrd������Ϣ.TextMatrix(i, 3) = IIf(IsNull(rs���Ҽ�¼.Fields("����")), "0", rs���Ҽ�¼.Fields("����"))
        cgrd������Ϣ.TextMatrix(i, 4) = IIf(IsNull(rs���Ҽ�¼.Fields("���")), "0.00", rs���Ҽ�¼.Fields("���"))
        cgrd������Ϣ.TextMatrix(i, 5) = IIf(IsNull(rs���Ҽ�¼.Fields("������")), "", rs���Ҽ�¼.Fields("������"))
        cgrd������Ϣ.TextMatrix(i, 6) = IIf(IsNull(rs���Ҽ�¼.Fields("���ѵ�λ����")), "", rs���Ҽ�¼.Fields("���ѵ�λ����"))
        cgrd������Ϣ.TextMatrix(i, 7) = IIf(IsNull(rs���Ҽ�¼.Fields("�վݺ�")), "", rs���Ҽ�¼.Fields("�վݺ�"))
        cgrd������Ϣ.TextMatrix(i, 8) = IIf(IsNull(rs���Ҽ�¼.Fields("���۱���")), "1", rs���Ҽ�¼.Fields("���۱���"))
        cgrd������Ϣ.TextMatrix(i, 9) = IIf(IsNull(rs���Ҽ�¼.Fields("��������")), "", rs���Ҽ�¼.Fields("��������"))
        
        rs���Ҽ�¼.MoveNext
        i = i + 1
    Loop
    
    Call subѡ��ͬ������
    Call ��̬����TextBox
    ctmr��ʱ.Enabled = False
    'cstu״̬��.Panels(1).Text = "���"
    cprg����.Value = cprg����.Max
    cprg����.Visible = False
    ctxt��ʾ.Visible = False
    Exit Sub
errhandle:
        ctmr��ʱ.Enabled = False
        ctxt��ʾ.Visible = False
        cprg����.Value = 0
        cprg����.Visible = False
        Call sfsub������("�շѽ������", "frm�˷�", "sub�˷�", Err.Number, Err.Description, True)
End Sub

'���ܣ���ȡ������ϵͳ��ǰʱ��ֵ
'���룺��
'�������
'���أ�ʱ��ֵ����ȷ����
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Function func��ȡ����������() As Date
Dim lrsTemp As New ADODB.Recordset '��ʱ���ִ�д洢���̻�õĽ��RecordSet
    On Error GoTo errhandle
    Set lrsTemp = dafuncGetData("Select getdate() as ����")
    func��ȡ���������� = Format(lrsTemp("����"), "yyyy-mm-dd hh:mm:ss")  'ȡ����
    Set lrsTemp = Nothing
    Exit Function
errhandle:
    Call sfsub������("�շѽ������", "frm�˷�", "func��ȡ����������", Err.Number, Err.Description, True)
End Function
 


'���ܣ���Ӧ�û��Ĺ���������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29

Private Sub mobjGUI_Operate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandle
    Select Case Operate
        Case "��ѯ"
            ctlb������.Buttons(1).Enabled = False
            cgrd������Ϣ.Clear
            cgrd������Ϣ.Rows = 1
            'cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |���۱���"
            cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |�վݺ� |���۱���|�շ�����  "
            sub�����
            ctlb������.Buttons(1).Enabled = True
        Case "���"
            'ɾ����ѯ����
            cinb�˷�����(3).Text = ""
            'cinb�˷�����(4).Text = ""
            cinb�˷�����(1).Text = ""
            cinb�˷�����(0).Text = ""
            
            cgrd������Ϣ.Clear
            cgrd������Ϣ.Rows = 1
            'cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |���۱���"
            cgrd������Ϣ.FormatString = "�շ�����        |�շѱ��       |�շ���Ŀ     |���� |���   |������  |���ѵ�λ          |�վݺ� |���۱���|�շ�����  "
        Case "�˷�"
            Call sub�˷�
            
            '���˷Ѻ󣬽����ϵ�����ͨ����ѯ���
            'ʱ�䣺2002/08/05 �켽��
            mobjGUI_Operate "��ѯ", False
        Case "��ӡ"
            If cgrd������Ϣ.Row < 1 Then Exit Sub
            '�޸ģ�2002-9-29�������ѡ���˷Ѽ�¼�����ӡ�˷�Ʊ�ݡ�
            If cgrd������Ϣ.Cell(flexcpBackColor, cgrd������Ϣ.Row, 0) = Label7.BackColor Then
                '�˷ѡ�
                Call sub��ӡ�˷�Ʊ��(cgrd������Ϣ.TextMatrix(cgrd������Ϣ.Row, 0))
            Else
                Call sub��ӡƱ��
            End If
    End Select
    Exit Sub
errhandle:
    sffuncMsg Operate & "���ɹ���" & Err.Description, sf����
End Sub

'���ܣ����Զ�δ�ӡ�˷���Ϣ
'ʱ�䣺2002/08/02
'���ߣ��켽��
Private Sub subѡ��ͬ������()
Dim i As Integer
Dim lcur��� As Currency
Dim lbln As Boolean
Dim lInt As Long
    ctxt�ܽ��.Text = "0.00 Ԫ"
    ctxt�˷�����.Text = ""
    
    lInt = cgrd������Ϣ.RowSel
    If lInt > 0 Then
         If cgrd������Ϣ.TextMatrix(lInt, 4) >= 0 Then
            lbln = True
         Else
            lbln = False
         End If
    End If
    With cgrd������Ϣ
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
        
            If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) Then
                .IsSelected(i) = True
            End If
        Next i
        
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                .TopRow = i
                Exit For
            End If
        Next i
                
        lcur��� = 0
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
               lcur��� = lcur��� + CCur(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 8))
            End If
        Next i
        ctxt�ܽ��.Text = CStr(lcur���) & " Ԫ"
        ctxt�˷�����.Text = .TextMatrix(.Row, 0)
    Else
        ctxt�ܽ��.Text = "0.00 Ԫ"
        ctxt�˷�����.Text = ""
    End If
    End With
    
    
    '����:����Ԫ������,�������û��ֵ,����û��Ȩ��,�˷�Ϊ������
    '����:�켽��
    'ʱ��:2002/08/14
    '�˷�Ȩ�޵Ŀ���
    If umfuncУ���û�Ȩ��("�շѹ���_�˷�") Then
        If ctxt�ܽ��.Text = "0 Ԫ" Then
            ctlb������.Buttons(4).Enabled = False
        Else
            ctlb������.Buttons(4).Enabled = True
        End If
    Else
        ctlb������.Buttons(4).Enabled = False
    End If
    
    '��ӡȨ�޵Ŀ���
    If umfuncУ���û�Ȩ��("�շѹ���_Ʊ�ݴ�ӡ") Then
        ctlb������.Buttons(5).Enabled = True
    Else
        ctlb������.Buttons(5).Enabled = False
    End If
    
    With cgrd������Ϣ
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
            If lbln = True Then
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, 4) >= 0 Then
                    .IsSelected(i) = True
                End If
            Else
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, 4) < 0 Then
                    .IsSelected(i) = True
                End If
            End If
        Next i
        
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                .TopRow = i
                Exit For
            End If
        Next i
                
        lcur��� = 0
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
               lcur��� = lcur��� + CCur(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 8))
            End If
        Next i
        ctxt�ܽ��.Text = CStr(lcur���) & " Ԫ"
        ctxt�˷�����.Text = .TextMatrix(.Row, 0)
    Else
        ctxt�ܽ��.Text = "0.00 Ԫ"
        ctxt�˷�����.Text = ""
    End If
    End With
End Sub

Private Sub ��̬����TextBox()
   
End Sub


'���ܣ�У�������Ϣ״̬
'˵�������շ�ǰ�������ݿ���У�������Ϣ��״̬��ֻ����
'      û���˷ѵ���Ϣ�����˷ѡ�
'����ֵ��funУ�����ϢΪtrue,��ʾ�����˷�,���򲻿����˷ѡ�
'���ߣ��켽��
'ʱ�䣺2002/08/05

Private Function funУ�����Ϣ(ByVal para�շ����� As String) As Boolean
On Error GoTo errhandler
    Dim lstrSql As String           '���������¼sql���
    Dim lobjRec As Object           '��������¼��ʱ�����
    
    '��ʼ����������ֵ
    funУ�����Ϣ = False
    
    If para�շ����� = "" Then
        Exit Function
    Else
        lstrSql = "select distinct(�շ�����),�շ�״̬ from �շѹ���_������Ϣ��  where �շ�����='" & para�շ����� & "'"
        Set lobjRec = dafuncGetData(lstrSql)
        
        If lobjRec.RecordCount > 0 Then
            If lobjRec("�շ�״̬") = 1 Then
                funУ�����Ϣ = True
            Else
                funУ�����Ϣ = False
            End If
        End If
    End If
Exit Function
errhandler:
    funУ�����Ϣ = False
End Function
