VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmQueryStatis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ְҵ�������-��ѯͳ��"
   ClientHeight    =   10680
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "ͳ�ƽ��"
      ForeColor       =   &H000080FF&
      Height          =   2175
      Left            =   7440
      TabIndex        =   49
      Top             =   4080
      Width           =   5775
      Begin VSFlex8Ctl.VSFlexGrid cgrdStatic 
         Height          =   1095
         Left            =   0
         TabIndex        =   50
         Top             =   240
         Width           =   5775
         _cx             =   2088773578
         _cy             =   2088765323
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
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
         AutoSearch      =   0
         AutoSearchDelay =   2
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.CommandButton CmdAction 
      BackColor       =   &H00C0FFFF&
      Caption         =   "��ѯ���"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton ccmdStatic 
      BackColor       =   &H0080FF80&
      Caption         =   "ͳ��"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "ͼ�� "
      ForeColor       =   &H000080FF&
      Height          =   4215
      Left            =   7440
      TabIndex        =   44
      Top             =   6360
      Width           =   5775
      Begin VB.PictureBox picChart 
         AutoSize        =   -1  'True
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3795
         ScaleWidth      =   5475
         TabIndex        =   45
         Top             =   240
         Width           =   5535
         Begin VB.Image Image1 
            Height          =   3495
            Left            =   240
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5175
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ѯ���� "
      ForeColor       =   &H000080FF&
      Height          =   3135
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Width           =   6255
      Begin VB.TextBox ctxt��λ���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1200
         TabIndex        =   31
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton ccmd��λ��λ 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox ccmb��ѯ���� 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox ccmb��ѯ���� 
         Height          =   300
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2280
         Width           =   3855
      End
      Begin VB.ComboBox ccmb��ѯ���� 
         Height          =   300
         Index           =   2
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox ccmb������ 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   3855
      End
      Begin VB.OptionButton cop������ 
         Caption         =   "���ϸ�"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   25
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton cop������ 
         Caption         =   "�ϸ�"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox ccmb������ 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1080
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTP��ʼ 
         Height          =   300
         Left            =   1200
         TabIndex        =   32
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��"
         Format          =   59637763
         CurrentDate     =   40986
      End
      Begin MSComCtl2.DTPicker DTP��ֹ 
         Height          =   300
         Left            =   3480
         TabIndex        =   33
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��"
         Format          =   59637763
         CurrentDate     =   40986
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "������"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "��λ����"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Σ������"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "��ҵ���"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�ֹ���"
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   39
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label7 
         Caption         =   "�������"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "������"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "��"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   36
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "�Ǽ����ڴ�"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "������"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame fraChartConfig 
      Caption         =   "ͳ��ͼ�����"
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   2640
      Width           =   6015
      Begin VB.TextBox YAxisTitle 
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "FrmQueryStatis.frx":0000
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox XAxisTitle 
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "FrmQueryStatis.frx":0008
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox ChartTitle 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "FrmQueryStatis.frx":0010
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox combͼ����ʽ 
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Text            =   "ͼ����ʽ"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox combɫ����ʽ 
         Height          =   300
         Left            =   2400
         TabIndex        =   17
         Text            =   "ɫ����ʽ"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox combͼ���� 
         Height          =   300
         Left            =   4200
         TabIndex        =   16
         Text            =   "ͼ����"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraChartStatic 
      Caption         =   "ͳ�����"
      ForeColor       =   &H000080FF&
      Height          =   1695
      Left            =   6480
      TabIndex        =   3
      Top             =   840
      Width           =   6015
      Begin VB.PictureBox tmpPicture 
         Height          =   495
         Left            =   4680
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox cchkRowColSwap 
         Caption         =   "���н���"
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox combYAxis 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox combXAxis 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton ccmdExportCrt 
         Caption         =   "����ͼ��"
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton ccmdDrawCrt 
         Caption         =   "����ͼ��"
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox CombInterval 
         Height          =   300
         Left            =   2160
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox ctxtInterval 
         Height          =   270
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox cchkInterval 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ÿ��"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "����"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "����"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frame4 
      Caption         =   "��ѯ���"
      ForeColor       =   &H000080FF&
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   7335
      Begin VSFlex8Ctl.VSFlexGrid cgrdInfo 
         Height          =   3495
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   6855
         _cx             =   2088775483
         _cy             =   2088769557
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
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
         AutoSearchDelay =   2
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSComDlg.CommonDialog ccdg 
         Left            =   8160
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   2760
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "ע���Ȳ�ѯ�������ͳ�ơ�"
      Height          =   180
      Left            =   9840
      TabIndex        =   48
      Top             =   3600
      Width           =   2160
   End
End
Attribute VB_Name = "FrmQueryStatis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���壺ְҵ����ѯͳ�ƽ���
'���ܣ���ְҵ�������Ϣ����ϸ��ѯ��ͳ��
'���ߣ�����
'ʱ�䣺2012-04-18
'��ע������
Option Explicit
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Public mblnInUse As Boolean
Dim lojb���� As Collection '���п���
Dim lobj��ѯͳ�ƺ��� As Object    '��ѯͳ�ƺ���
Dim mobj���� As Object      '���没���ϼ�Ŀ¼
Dim isHistory As Boolean    '�Ƿ�Ϊ������ѯ
'2012-04-19 �ڵ�� ��
'���ͳ�Ʋ���excel��ر���
Private hasStatPerm As Boolean
Private initChart As Boolean
Private xlApp As Object     'Excel.Application
Private xlBook As Object    'Excel.Workbook
Private xlSheet As Object   'Excel.Worksheet
Private xlChart As Object   'Excel.Chart
'2012-04-19 �ڵ�� ��

Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property

'Private Sub ccmdBack_Click()
'
'    Set cgrdInfo.DataSource = mobj����
'    ccmdBack.Visible = False
'    Set mobj���� = Nothing
'
'End Sub

'2012-05-20 �ڵ��
'��Ӽ���ͳ�ƺ��������յ�ǰ����ͳ��
Private Sub ccmdStatic_Click()
    Dim XSelected As Integer
    Dim YSelected As Integer
    '�޸��ˣ����� 2012.12.04
    'bug�ţ�0000055
    '˵������ʼʱ�䲻���ڽ���ʱ��֮�󣬸�����ʾ������
    If DTP��ʼ.Value > DTP��ֹ.Value Then
        MsgBox "��ʼ���ڲ����ڽ�������֮��"
        DTP��ʼ.Value = DTP��ֹ.Value
        Exit Sub
    End If
    '2012.12.04     ����
    If cgrdInfo.rows = 1 Then Exit Sub
    If isHistory Then Exit Sub
    
    ccmdStatic.Caption = "ͳ����..."
    
    XSelected = combXAxis.ItemData(combXAxis.ListIndex)
    YSelected = combYAxis.ItemData(combYAxis.ListIndex)
    
    '2012-05-29 �ڵ�� ��
    '�޸�ͳ���Ӻ�����ÿ�δ���ͳ�Ʒ����ͳ������
    '��ʼ��
'    SSTabGrid.Tab = 1
    Select Case XSelected
    Case 0  '������ͳ��
        sub������ͳ�� XSelected, YSelected
    Case 1  '����ҵͳ��
        sub����ҵͳ�� XSelected, YSelected
    Case 2  '����λͳ��
        sub����λͳ�� XSelected, YSelected
    Case 3  '��������ͳ��
        sub��������ͳ�� XSelected, YSelected
    '2012-08-08 �ڵ�� ��
    Case 4  '��Σ������ͳ��
        sub��Σ������ͳ�� XSelected, YSelected
    '2012-08-08 �ڵ�� ��
    '2013-03-31 ������ ��
    Case 5
        sub��ʱ������ͳ�� XSelected, YSelected
    '2013-03-31 ������ ��
    End Select
    '2012-05-29 �ڵ�� ��
    
    ccmdStatic.Caption = "ͳ��"
    
    '����cgrdStatic�����ݣ���excel chartͼ��
    SelectData (cchkRowColSwap)
    DrawChart

End Sub

'Private Sub ccmd������ѯ_Click()
'
'    Dim lobjRec As Object, lobjNo As Object
'    Dim str As String, lstrNo As String
'    Dim i As Integer
'    Dim date1 As String
'    Dim date2 As String
''    str = "select a.ϵͳ���,b.����,a.���ֽ���,a.ҽ�����,a.�������� " _
''    & "from ְҵ�����_���ҽ��۱� a join ϵͳ����_�ֵ�_�ֵ����ݱ� b on a.���� = b.��� and b.id = 84"
'    If Trim(ctxt����(0).Text) = "" And Trim(ctxt����(1).Text) = "" Then
'        MsgBox "��ָ�������Ա��������ϵͳ��ţ�"
'        Exit Sub
'    End If
'
'    str = "select ϵͳ��� from ְҵ�����_���������ݿ� where 1 = 1 "
'
'    '�޸��ˣ����� 2012.12.18  ����
'    '�޸�˵����ȥ��һ�������š�
'    If Trim(ctxt����(0).Text) <> "" Then
''        str = str & " and ����=''" & Trim(ctxt����(0).Text) & "''"
'        str = str & " and ����='" & Trim(ctxt����(0).Text) & "'"
'    End If
'
'    If Trim(ctxt����(1).Text) <> "" Then
'        '�޸��ˣ����� 2013.01.04   ����
'        '�޸�˵��������ϵͳ��Ų�Ψһ�����԰���ϵͳ��Ų�ѯ���ѯ�������ݣ�
'        '          ���Ծ��ñ�Ų�ѯ��ǰ��Ŷ�Ӧ��������Ȼ�����������ѯ��Ӧ��ʷ���ݡ�
''        str = str & " and ϵͳ��� =''" & Trim(ctxt����(1).Text) & "''"
''        str = str & " and ϵͳ��� ='" & Trim(ctxt����(1).Text) & "'"
'        Set lobjNo = dafuncGetData("select ���� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & Trim(ctxt����(1).Text) & "'")
'        If Not (lobjNo.EOF Or lobjNo.BOF) Then
'            str = str & " and ����='" & lobjNo(0) & "'"
'        Else
'            MsgBox "������ϵͳ��Ų�������������ѯ���ݡ�"
'        End If
'        Set lobjNo = Nothing
'        '�޸��ˣ����� 2013.01.04   ����
'    End If
'
'    If Trim(ccmb������λ.Text) <> "" Then
''        str = str & " and ��λ����=''" & Trim(ccmb������λ.Text) & "''"
'        str = str & " and ��λ����='" & Trim(ccmb������λ.Text) & "'"
'    End If
'
'    If Trim(ccmbΣ������.Text) <> "����" Then
''        str = str & " and Σ������=''" & Trim(ccmbΣ������.Text) & "''"
'        str = str & " and Σ������='" & Trim(ccmbΣ������.Text) & "'"
'    End If
'    '�޸��ˣ����� 2012.12.18  ����
'    date1 = DTP����begin.Value
'    date2 = DTP����end.Value
'
'    '�޸��ˣ����� 2012.12.18  ����
'    '�޸�˵����֮ǰ��select���ֱ��Ƕ����exec��ִ�У��﷨�����⣬���������ִ��select�����ֵ����exec��ִ�С�
''    str = str & "'"
'    '�����޸��ˣ����� 2013.01.04  ����
'    '�޸�˵��������һ���˿����ж����첡ʷ�����Բ�ѯ������Ҫ������ϵͳ��ż����ѯ������
'    Set lobjNo = dafuncGetData(str)
'    lstrNo = ""
'    For i = 0 To lobjNo.RecordCount - 1
'        If Not lstrNo = "" Then
'            lstrNo = lstrNo & ",''" & lobjNo(0) & "''"
'            lobjNo.MoveNext
'        Else
'            lstrNo = "''" & lobjNo(0) & "''"
'            lobjNo.MoveNext
'        End If
'    Next
'    '�޸��ˣ����� 2013.01.04   ����
'    If Not lstrNo = "" Then
''    Set lobjRec = dafuncGetData("exec ְҵ�����_������ز�����Ϣ " & str & ",'" & date1 & "','" & date2 & "'")
'        Set lobjRec = dafuncGetData("exec ְҵ�����_������ز�����Ϣ '" & lstrNo & "','" & date1 & "','" & date2 & "'")
'        '�޸��ˣ����� 2012.12.18  ����
'        cgrdInfo.rows = 1
'        If Not (lobjRec.BOF Or lobjRec.EOF) Then
'            cgrdInfo.SelectionMode = 3
'            Set cgrdInfo.DataSource = lobjRec
'    '        With cgrdInfo
'    '            .Cols = .Cols + 1: .TextMatrix(0, .Cols - 1) = IIf()
'    '        End With
'            cgrdInfo.AutoResize = True
'            cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, , True
'            'cgrdInfo.AutoSizeMode = flexAutoSizeColWidth
'            cgrdInfo.AllowSelection = False
'            isHistory = True
'            Set mobj���� = lobjRec
'        End If
'    Else
'        cgrdInfo.rows = 1
'    End If
'End Sub

'���ǣ�2012-10-22
Private Sub ccmd��λ��λ_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��
    Dim lobj��λ As Object
    Dim lobj��λ��Ϣ As Object
    Dim mstr��λ������ As String
    '������λ��λ���档
'    Set lobjRec = pobjҵ�����.func��λ��λ        'ԭ���õ�λ��λע�͵�  2016-1-21 by Ĳ��
    frmQueryCompanyLocation.Show 1, Me    '���õ�λ��λ��ѯ 2016-1-21 by Ĳ��
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxt��λ����.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    
    '�ѽ���ص���λ¼��򡣱����ܱ����µ�λ��λ��Ϣ��
    ctxt��λ����.SetFocus
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "��ѯͳ�ƽ���", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub

Private Sub cgrdInfo_DblClick()
    Dim lobjRec As Object
    Dim tempDepart As String
    Dim tempNo As String
    Dim tempDNo As String
    Dim tempDate As String
    On Error GoTo errHandler
    
    If Not isHistory Then Exit Sub
    
    If Not mobj���� Is Nothing Then Exit Sub
    
    If cgrdInfo.Row > 0 Then
        tempNo = cgrdInfo.TextMatrix(cgrdInfo.Row, 0)
        tempDepart = cgrdInfo.TextMatrix(cgrdInfo.Row, 1)
        If Trim(tempDepart) = "���ս���¼��" Then
            MsgBox "û�пɲ鿴������"
            Exit Sub
        End If
        tempDNo = cgrdInfo.TextMatrix(cgrdInfo.Row, 3)
        tempDate = cgrdInfo.TextMatrix(cgrdInfo.Row, 4)
        Set lobjRec = dafuncGetData("select * from ְҵ�����_�����Ϣ_" & tempDepart & " where " _
        & "ϵͳ���='" & tempNo & "' and ���ҽʦ='" & tempDNo & "' and convert(varchar(10),��дʱ��,120)='" & tempDate & "'")
        
        If Not lobjRec.EOF Then
'            ccmdBack.Visible = True
            Set mobj���� = cgrdInfo.DataSource
            Set cgrdInfo.DataSource = lobjRec
        End If
        
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

'��ѯͳ�ƣ����ǣ�2012-10-22
Private Sub CmdAction_Click()
    Dim i As Integer
    Dim sql As String
    Dim lobjRec As Object
    On Error GoTo errHandler

    sql = "select * from ְҵ�����_��ѯͳ����ͼ where 1=1"
    For i = 0 To ccmb��ѯ����.Count - 1
        If ccmb��ѯ����(i).Text <> "����" Then
            sql = sql & " and " & Label2(i).Caption & " = '" & ccmb��ѯ����(i).Text & "'"
        End If
    Next
    
    If Trim(ctxt��λ����.Text) <> "" Then
        sql = sql & " and ��λ���� = '" & Trim(ctxt��λ����.Text) & "'"
    End If
    
    If ccmb������.Text <> "����" Then
        sql = sql & " and ���� = '" & ccmb������.Text & "'"
    End If
    '�޸��ˣ����� 2012.12.18
    '�޸�˵����������Ū���ˣ�Ӧ���ǡ������Ա���͡�����
    If ccmb������.Text <> "����" Then
        sql = sql & " and �����Ա���� = '" & ccmb������.Text & "'"
    End If
    '�޸��ˣ����� 2012.12.18  ����
    Dim dtpTimeTo As Date
    
    dtpTimeTo = Format(DateAdd("m", 1, DTP��ֹ.Value), "yyyy-mm-01")
    dtpTimeTo = Format(DateAdd("d", -1, dtpTimeTo), "yyyy-mm-dd")
    sql = sql & " and (������� between '" & Format(DTP��ʼ.Value, "yyyy-mm" & "-01 00:00:00") & "' and '" & Format(dtpTimeTo, "yyyy-mm-dd" & " 23:59:59") & "')"
    '2013-03-04 ������
    '����Ҫ�ϸ��벻�ϸ�
'    If cop������(0) Then
'        sql = sql & " and ���״̬ = '�ѷ�����'"
'    Else
'        sql = sql & " and ���״̬ = '������'"
'    End If
    sql = sql & " and ���״̬ in('�ѷ�����','������','�Ѹ���')"
    '2013-03-04 ������
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(sql)
'    If SSTabGrid.TabIndex = 1 Then
'        SSTabGrid.TabIndex = 0
'    End If
    cgrdInfo.rows = 1
    Label8.Caption = "������" & cgrdInfo.rows - 1
    If Not (lobjRec.EOF Or lobjRec.BOF) Then
        Set cgrdInfo.DataSource = lobjRec
        cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
        isHistory = False
        
    Else
        MsgBox "û�з��ϲ�ѯ�����Ľ����", vbInformation, "ϵͳ��ʾ"
    End If
    Label8.Caption = "������" & cgrdInfo.rows - 1  'ͳ������ 2016-1-20 by Ĳ��
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume

End Sub

'2012-05-30 �ڵ��
'��x��Ϊ���״��ʱ�������������������΢��һ����
'��ʱ��������ѡ��ϸ��������������ϸ������������ϸ��ʡ������޽��������������
Private Sub combXAxis_Click()
    If combXAxis.ListIndex = 3 Then
        combYAxis.Enabled = False
    Else
        'combYAxis.Enabled = True
    End If
End Sub

'2012-05-20 �ڵ�� �ж�ʱ�������ָ�ʽ
Private Sub ctxtInterval_LostFocus()
    If ctxtInterval.Text = "" And cchkInterval.Value = 1 Then MsgBox ("����������"): Exit Sub
    If IsNumeric(ctxtInterval.Text) = False Then MsgBox ("����������"): Exit Sub
End Sub

Private Sub DTP��ֹ_Change()
    If (DTP��ֹ.Value - DTP��ʼ.Value) / 30 < 1 Then
        MsgBox "������һ�������ϡ�"
        DTP��ʼ.Value = DateAdd("m", -1, Format(DTP��ֹ.Value, "yyyy/MM"))
'        Exit Sub
    End If
    DTP��ʼ.Value = Format(DTP��ʼ.Value, "yyyy/MM")
'    DTP��ֹ.Value = Format(DTP��ֹ.Value, "yyyy/MM")
    DTP��ֹ.Value = DateAdd("d", -1, Format(DateAdd("M", 1, Format(DTP��ֹ.Value, "yyyy/MM")), "yyyy/MM"))
End Sub

Private Sub DTP��ʼ_Change()
    If (DTP��ֹ.Value - DTP��ʼ.Value) / 30 < 1 Then
        MsgBox "������һ�������ϡ�"
        DTP��ʼ.Value = DateAdd("m", -1, Format(DTP��ֹ.Value, "yyyy/MM"))
'        Exit Sub
    End If
'    DTP��ʼ.Value = Format(DTP��ʼ.Value, "yyyy/MM")
'    DTP��ֹ.Value = Format(DTP��ֹ.Value, "yyyy/MM")
'    DTP��ֹ.Value = DateAdd("d", -1, Format(DateAdd("M", 1, Format(DTP��ֹ.Value, "yyyy/MM")), "yyyy/MM"))
End Sub

Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    Dim i As Integer
    On Error GoTo errHandler
    Dim lojbRec As Object   '���ݿ�������

    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    '���ù�����������Ҫ�ĸ��ְ�ť��
    '�޸ģ�2002-7-1�������ȡ�����۵Ĳ�������Ϊ������ѡ��
    With lcol��������ť
        '2012-04-19 �ڵ�� ��
        '�޸����ݣ�ֻ֧�ֵ��뵼����ʽΪexcel��ʽ��
        .Add "����Excel(&O)113"
        .Add "|"
        .Add "����ͼ��(&T)102"
        .Add "Ԥ������(&L)108"
        .Add "��ӡ����(&P)107"
        '2012-04-19 �ڵ�� ��
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""
    
    '2012-05-23 ���� ������
    '����Ȩ������
'    Dim lobjTmp As Object
'    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
'    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_ְҵ����ѯͳ��_����") = False Then
'        ctlb������.Buttons(1).Visible = False
'    End If
'    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_ְҵ����ѯͳ��_����") = False Then
'        ctlb������.Buttons(2).Visible = False
'    End If
'    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_ְҵ����ѯͳ��_��ӡ") = False Then
'        ctlb������.Buttons(3).Visible = False
'        ctlb������.Buttons(4).Visible = False
'    End If
'    Set lobjTmp = Nothing
    '2012-05-23 ������
    
    '��ѯ������ʼ��
    '������ѯͳ�ƺ�������
    Set lobj��ѯͳ�ƺ��� = CreateObject("ְҵ������.clsQueryStatis")
    
    '2012-04-19 �ڵ�� ��
    'Ӧ����ӵ��Ȩ�޵�ǰ���£�����ͳ�ƵĲ���
    'if ��ͳ�Ƶ�Ȩ��=true then
    hasStatPerm = True
'    Else
'        hasStatPerm = False
    'end if
    '��һ����ʱ��excel�ļ�
    OpenTempExcel
    
    '��ʼ��ͳ�Ʋ��������б�
    subInitChartList
    subInitStaticList
    
    '2012-04-19 �ڵ�� ��
    
    '2012-04-23 �ڵ�� ��
    '����cgrdInfo�������
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    '2012-04-23 �ڵ�� ��
    
    '2012-05-20 �ڵ�� ��
    '���ͳ�Ʋ��ֿؼ���ʼ����Ϣ
'    DTP��ʼ.Value = Format(DateAdd("M", -1, Date), "yyyy/MM")
'    DTP��ֹ.Value = DateAdd("d", -1, Format(DateAdd("M", 1, Date), "yyyy/MM"))
'    chkStart.Value = 1
'    chkEnd.Value = 1
'    DTP����begin.Value = DateAdd("d", -30, Date)
'    DTP����end.Value = Date
    
    cchkInterval.Value = 1
    ctxtInterval.Text = "1"
    With CombInterval   'û�ж��·ݡ����ȡ��������ϸ�ж�
        .Clear
'        .AddItem "��": .ItemData(.NewIndex) = 1
        .AddItem "��": .ItemData(.NewIndex) = 1
        .AddItem "��": .ItemData(.NewIndex) = 3
        .AddItem "��": .ItemData(.NewIndex) = 12
        .ListIndex = 0
    End With
'    SSTab��ѯͳ�ƽ��.TabIndex = 0
'    SSTabGrid.TabIndex = 0
    '2012-05-20 �ڵ�� ��
    'û�л���ͼ�����ܴ�ӡ����͵���excel
    If picChart.Picture = 0 Then
'        ctlb������.Buttons(4).Enabled = False
'        ctlb������.Buttons(5).Enabled = False
'        ctlb������.Buttons(4).Enabled = False
    End If
    
    '����ʼ�������ǣ�2012��
    cgrdInfo.cols = 0
    With cgrdInfo
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ϵͳ���"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�շ�����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "Σ������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��ҵ���"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����ϵͳ���"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�½�������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ҽʦ����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�շѽ��"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "���״̬"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    If picChart.Picture = 0 Then
        ccmdExportCrt.Enabled = False
    End If
    cop������(0).Value = True
    '�޸��ˣ����� 2012.12.04     ����
    'bug�ţ�0000054
    '˵��������Ԥ���ʹ�ӡ���档
'    ctlb������.Buttons(2).Visible = False
    ctlb������.Buttons(5).Visible = False
    ctlb������.Buttons(4).Visible = False
    '2012.12.04     ����
    '����ʱ��ؼ�ִ�����´���
    Timer1.Enabled = True
    '2012-05-21 ��¶
    '����Ȩ������
'    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
'    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_��ѯͳ��_����") = False Then
'        ctlb������.Buttons(1).Visible = False
'    End If
'    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_��ѯͳ��_����") = False Then
'        ctlb������.Buttons(2).Visible = False
'    End If
'    Set lobjTmp = Nothing
    '2012-05-21
    
    'MsgBox (DTP��ʼ.Value & " || " & Format(DTP��ʼ.Value, "yyyy-mm-dd") & " " & Format("00:00:00", "hh:mm:ss")) ''''''''''test
    'MsgBox (DTP��ֹ.Value & " || " & Format(DTP��ֹ.Value, "yyyy-mm-dd") & " " & Format("00:00:00", "hh:mm:ss")) ''''''''''test
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    mblnInUse = False
    
    '2012-05-23 �ڵ�� ��
    '�˳�����ʱ��ǿ�йرս���
    If hasStatPerm = True Then CloseTempExcel
    '2012-05-23 �ڵ�� ��
    
    Set mobjGUI = Nothing
    Exit Sub

    '2012-05-24 �ڵ�� ��
errHandler:
    mblnInUse = False
    Set xlChart = Nothing
    Set xlSheet = Nothing
    xlApp.Workbooks.Close
    xlApp.Quit
    Set xlBook = Nothing
    If Not xlApp Is Nothing Then
        Shell "cmd.exe /c taskkill /f /im excel.exe"
    End If
    Set xlApp = Nothing
    Set mobjGUI = Nothing
    '2012-05-24 �ڵ�� ��
End Sub

Private Sub ccmdDrawCrt_Click()
    If cgrdInfo.rows = 1 Then Exit Sub
    If isHistory Then Exit Sub
    SelectData (cchkRowColSwap.Value)
    DrawChart
End Sub

Private Sub cchkRowColSwap_Click()
'    SelectData (cchkRowColSwap.Value)
    selectdataRC (cchkRowColSwap.Value)  '2016-1-27 by Ĳ��
    DrawChart
End Sub

Private Sub ccmdExportCrt_Click()
    '
    Dim lstrFile As String
    ccdg.FileName = ""
    ccdg.Filter = "JPEG(*.jpg)|*.jpg|" & _
                "97-03 Excel(*.xls)|*.xls|" & _
                "07 Excel(*.xlsx)|*.xlsx"
    ccdg.ShowSave
    lstrFile = ccdg.FileName
    If lstrFile <> "" Then
        If ccdg.FilterIndex = 1 Then        '��chart��Ϊjpg
            VB.SavePicture picChart.Picture, lstrFile
        ElseIf ccdg.FilterIndex = 2 Then    '��chart�����ݴ�Ϊxls��xlsx(��֧�ֵ�ǰ�汾��)
            Select Case xlApp.Application.Version
            Case "12.0"
                xlBook.SaveCopyAs (Replace(lstrFile, ".xls", ".xlsx"))
            Case Else
                xlBook.SaveCopyAs (lstrFile)
            End Select
        ElseIf ccdg.FilterIndex = 3 Then
            Select Case xlApp.Application.Version
            Case "12.0"
                xlBook.SaveCopyAs (lstrFile)
            Case Else
                xlBook.SaveCopyAs (Replace(lstrFile, ".xlsx", ".xls"))
            End Select
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    Picture1.Left = 0
'    Picture1.Width = Me.ScaleWidth - Picture1.Left
'    Picture1.Height = Me.ScaleHeight - Picture1.Top
'    Frame1.Width = Picture1.Width - Frame1.Left
'    Frame1.Height = Picture1.Height - Frame1.Top
'    ctlb������.Width = Frame1.Width - ctlb������.Left
'
''    SSTab��ѯͳ�ƽ��.Width = Frame1.Width - SSTab��ѯͳ�ƽ��.Left - 60
''    SSTab��ѯͳ�ƽ��.Height = Frame1.Height - SSTab��ѯͳ�ƽ��.Top - 120
''    Frame4.Width = SSTab��ѯͳ�ƽ��.Width - cgrdInfo.Left - 60
'    frame4.Height = Frame1.Height - frame4.Top - 120
'    cgrdInfo.Width = frame4.Width - cgrdInfo.Left - 60
'    cgrdInfo.Height = frame4.Height - cgrdInfo.Top - 120
''    Frame6.Width = SSTab��ѯͳ�ƽ��.Width - Frame6.Left - 60
''    Frame6.Height = SSTab��ѯͳ�ƽ��.Height - Frame6.Top - 120
'    cgrdStatic.Width = Frame6.Width - cgrdStatic.Left - 60
'    cgrdStatic.Height = Frame6.Height - cgrdStatic.Top - 120
    ctlb������.Width = Me.ScaleWidth
    Frame3.Width = Me.ScaleWidth * 2 / 5
    fraChartStatic.Left = Frame3.Left + Frame3.Width + 20
    fraChartStatic.Width = Me.ScaleWidth - fraChartStatic.Left - Frame3.Left
    fraChartConfig.Left = fraChartStatic.Left
    fraChartConfig.Width = fraChartStatic.Width
    Frame4.Left = Frame3.Left
    Frame1.Left = fraChartStatic.Left
    Frame4.Width = Frame3.Width
    Frame1.Width = fraChartStatic.Width
    Frame4.Height = Me.ScaleHeight - Frame4.Top - 20
    Frame1.Height = Frame4.Height
    cgrdInfo.Width = Frame4.Width - cgrdInfo.Left * 2
    cgrdInfo.Height = Frame4.Height - cgrdInfo.Top - 10
    picChart.Width = Frame1.Width - picChart.Left * 2
    picChart.Height = Frame1.Height - picChart.Top - 10
    Image1.Width = FrmQueryStatis.Width - Frame4.Width - 700 '��image�Ĵ�С����
'    Image1.Width = 8000
'    Image1.Height = Frame4.Height - Frame2.Height - 1000
    CmdAction.Left = Frame3.Left + Frame3.Width + 100
    ccmdStatic.Left = CmdAction.Left + CmdAction.Width + 100
    Label9.Left = ccmdStatic.Left + ccmdStatic.Width + 100
    Frame2.Left = Frame1.Left
    Frame2.Width = Frame1.Width
    cgrdStatic.Height = Frame2.Height - cgrdStatic.Top - 10
    cgrdStatic.Width = Frame2.Width - cgrdStatic.Left * 2
    
End Sub



Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Dim paraԤ������ As Boolean
    Dim para��ӡ���� As Boolean
    Cancel = True
    
    Select Case Operate
    '2012-04-19 �ڵ�� ��
    '���excel����vsflexgrid��vsflexgrid����excel��
    Case "����Excel"
        '����ط������ж���жϣ���ʱֻ����Excel�ġ�
        '��Ϻ���Ķ������ݲ��֣������жϲ�ͬ�ļ����͵����롣
        ccdg.Filter = "Excel file" & "(*.xls)|*.xls" & _
                    "|Batch Files (*.bat)|*.bat|" & _
                    "All Files (*.*)|*.*"
        ccdg.FileName = ""
        ccdg.ShowOpen
        '2012-05-20 �ڵ�� ��
'        SSTab��ѯͳ�ƽ��.Tab = 1
'        SSTabGrid.Tab = 1
        'subInitChartList
        If ccdg.FileName = "" Then Exit Sub
        '2012-05-20 �ڵ�� ��
        sub��ʾ������Ϣ
        If hasStatPerm = True Then SelectData (cchkRowColSwap): DrawChart
    Case "����ͼ��"
        ccmdExportCrt_Click
    Case "����Excel"
        'ע��᲻�����ccdgδ�������⡣
        Dim lstrFile As String
        ccdg.Filter = "Excel�ļ� (*.xls)|*.xls" & "|Excel 2007 files (*.xlsx)|*.xlsx"
        ccdg.ShowSave
        lstrFile = ccdg.FileName 'Replace(ccdg.FileName, ".xls", "") & "_" & Date & ".xls"
        If lstrFile <> "" Then
            'cgrdMain.ColDataType(0) = flexDTString '����ʱ�����еĸ�ʽΪ�ַ��� flexFileExcel
            xlSheet.SaveAs lstrFile
'            cgrdStatic.SaveGrid lstrFile, flexFileData, True 'trueʱ��������ͷ��false��������ͷ
        End If
    '2012-04-19 �ڵ�� ��
    '2012-06-06 �ڵ�� ��
    'Ԥ������ӡͳ��ͼ���棬ˮ�������ʽ
    Case "Ԥ������"
        paraԤ������ = True
        para��ӡ���� = False
        sub��ӡͳ�Ʊ��� paraԤ������, para��ӡ����
    Case "��ӡ����"
        paraԤ������ = False
        para��ӡ���� = True
        sub��ӡͳ�Ʊ��� paraԤ������, para��ӡ����
    '2012-06-06 �ڵ�� ��
    Case "�˳�"
        '2012-04-19 �ڵ�� ��
'        If hasStatPerm = True Then
'            CloseTempExcel
'        End If
        '2012-04-19 �ڵ�� ��
        Set mobj���� = Nothing
        Unload Me
    End Select
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub


'�޸ģ���ʼ����ѯ���������ǣ�2012-10-22
Private Sub Timer1_Timer()
    Dim lobjRec As Object   '���ݿ�������
    Dim i As Integer
    On Error GoTo errHandler
    
    Timer1.Enabled = False

    '����ʱ������
    DTP��ʼ.Value = DateAdd("M", -1, Now)
    DTP��ֹ.Value = Now
    
    '��ȡ��ҵ���
    Set lobjRec = lobj��ѯͳ�ƺ���.func��ȡ��ҵ���
    If lobjRec.RecordCount > 0 Then
        ccmb��ѯ����(0).AddItem "����"
        For i = 1 To lobjRec.RecordCount
            ccmb��ѯ����(0).AddItem lobjRec("����")
            lobjRec.MoveNext
        Next
    End If
    
    '��ȡΣ������
    Set lobjRec = lobj��ѯͳ�ƺ���.funcΣ������
    If lobjRec.RecordCount > 0 Then
        ccmb��ѯ����(1).AddItem "����"
'        ccmbΣ������.AddItem "����"
        For i = 1 To lobjRec.RecordCount
            ccmb��ѯ����(1).AddItem lobjRec("����")
'            ccmbΣ������.AddItem lobjRec("����")
            lobjRec.MoveNext
        Next
    End If
    
    '��ȡ����
    Set lobjRec = lobj��ѯͳ�ƺ���.func��ȡ����
    If lobjRec.RecordCount > 0 Then
        ccmb��ѯ����(2).AddItem "����"
        For i = 1 To lobjRec.RecordCount
            ccmb��ѯ����(2).AddItem lobjRec("����")
            lobjRec.MoveNext
        Next
    End If
    
'    ccmbΣ������.ListIndex = 0
    ccmb��ѯ����(0).ListIndex = 0
    ccmb��ѯ����(1).ListIndex = 0
    ccmb��ѯ����(2).ListIndex = 0
    
    '�������ȡ
    Set lobjRec = lobj��ѯͳ�ƺ���.func��ȡ�������
    If lobjRec.RecordCount > 0 Then
        ccmb������.AddItem "����"
        For i = 1 To lobjRec.RecordCount
            ccmb������.AddItem lobjRec("����")
            lobjRec.MoveNext
        Next i
    End If
    ccmb������.Text = "����"

    '��ȡ��������
    Set lobjRec = lobj��ѯͳ�ƺ���.func��ȡ����
    If lobjRec.RecordCount > 0 Then
        ccmb������.AddItem "����"
        For i = 1 To lobjRec.RecordCount
            ccmb������.AddItem lobjRec("��������")
            lobjRec.MoveNext
        Next i
    End If
    ccmb������.Text = "����"
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetMedicalExamTemplate", "Timer1_Timer", 6666, lstrError, False
    MousePointer = 0
    '�ָ�������Բ�����
    Me.Enabled = True
End Sub

'2012-04-19 �ڵ��
'�������ܣ�����excel�ĵ�ʱ����ʾ�����ݡ�
Sub sub��ʾ������Ϣ()
    On Error GoTo errHandler
    
    With cgrdStatic
        '����cgrdStatic����ʽ
        .FixedRows = 0: .FixedCols = 0
        
        '�����֮ǰ���е���Ϣ
        .Clear
        
        '����ط������ж���жϣ���ʱֻ����Excel�ġ�
        '��Ϻ���Ķ������ݲ��֣������жϲ�ͬ�ļ����͵����롣
        .LoadGrid ccdg.FileName, flexFileExcel
        .AutoSize 1, .cols - 1, 0, 0
        
        '����cgrdStatic����ʽ
        .FixedRows = 1: .FixedCols = 1
        
        '2012-05-23 ��¶
        'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
        cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
        cgrdStatic.ExplorerBar = flexExSort
        cgrdStatic.DataMode = flexDMFree
        '2012-05-23
    End With

    Exit Sub
errHandler:
    '���û�е����ļ�������ʾ�κ���Ϣ��
    If ccdg.FileName = "" Then Exit Sub
    MsgBox ("����Excel�ļ�����")
End Sub


'2012-05-20 �ڵ��
'�ж��ַ������Ƿ��С����족�������顱�������ϸ񡱡����У��򲻺ϸ�
Private Function sub���պϸ�(ByVal paraCon As String) As Integer
    If paraCon = "" Then sub���պϸ� = 1: Exit Function
    If InStr(paraCon, "����") > 0 Or InStr(paraCon, "����") > 0 Or InStr(paraCon, "���ϸ�") > 0 Then
        sub���պϸ� = 0
        Exit Function
    Else
        sub���պϸ� = 1
        Exit Function
    End If
End Function

'2012-04-19 �ڵ��
Sub SelectData(ByVal ifSwap As Integer)
    ccmdStatic.Caption = "����������..."
    
'    xlSheet.Activate
    xlSheet.rows.Clear
    Dim i, j As Integer
    For i = 0 To cgrdStatic.rows - 1
         For j = 0 To cgrdStatic.cols - 1
            If ifSwap = 0 Then
                xlSheet.Cells(i + 1, j + 1) = cgrdStatic.TextMatrix(i, j)
            Else
                xlSheet.Cells(j + 1, i + 1) = cgrdStatic.TextMatrix(i, j)
            End If
        Next j
    Next i
    xlSheet.Activate  '��ѡ��Ԫ��ǰҪ�������ڹ�������Ȼ���˳�����ڶ��β�ѯͳ��ʱ�����  2016-1-28 by Ĳ��
    xlSheet.Cells.Select
'    xlSheet.Shapes.addchart.Select
    xlSheet.Shapes.SelectAll    '2015-12-7 by Ĳ��
'    Set xlChart = xlApp.ActiveChart   '�ڶ��μ�֮��ʹ��ʱ���ᱨ��Զ�̷�����������"
    Set xlChart = xlBook.Charts.Add   '2016-1-27 by Ĳ��
    initChart = True
    
    ccmdStatic.Caption = "ͳ��"
End Sub

'���н���������������������selectdata 2016-1-27 by Ĳ�� ��
Sub selectdataRC(ByVal ifSwap As Integer)
    ccmdStatic.Caption = "����������..."
    
'    xlSheet.Activate
    xlSheet.rows.Clear
    Dim i, j As Integer
    For i = 0 To cgrdStatic.rows - 1
         For j = 0 To cgrdStatic.cols - 1
            If ifSwap = 0 Then
                xlSheet.Cells(i + 1, j + 1) = cgrdStatic.TextMatrix(i, j)
            Else
                xlSheet.Cells(j + 1, i + 1) = cgrdStatic.TextMatrix(i, j)
            End If
        Next j
    Next i
'    xlSheet.Cells.Select
    xlSheet.Shapes.SelectAll
    Set xlChart = xlBook.Charts.Add
    initChart = True
    
    ccmdStatic.Caption = "ͳ��"
End Sub
'2016-1-27 by Ĳ�� ��

'2012-04-19 �ڵ��
'��һ����ʱ��excel�ļ�������ͳ��ʱ���в�����
Sub OpenTempExcel()
    Set xlApp = CreateObject("excel.application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets("sheet1")
    xlSheet.Activate
End Sub

'2012-04-19 �ڵ��
'�ر�ͳ�ƹ������õ���excel�ļ������ͷ��ڴ档
'������ر��ļ�����һֱ���ڴ���ռ����Դ
Sub CloseTempExcel()
    On Error GoTo errHandler

    '2012-05-22 �ڵ�� ��
    '���������ʱexcelд�����ݣ���������������˳�����ʱ�������̡�
    '�����ʱexcel����д�������ݣ��ڵ�ǰ�����˳�ʱ�����̲��ܽ�����
    'ֻ�ܵ�����ϵͳ�˳�ʱ��VB�Źرս��̡�
    Set xlChart = Nothing
    Set xlSheet = Nothing
    xlApp.DisplayAlerts = False
    xlBook.Save
    xlBook.Close (True)
    xlApp.Workbooks.Close
    xlApp.Quit
    Set xlBook = Nothing
    '    2012-05-23 �ڵ�� ��
    '    ǿ�ƹر�һ��excel���̣����ڹرյ��ĸ���Ҫ�Ȼ�����ԡ�
    '    ���м������޷�ִ�е���һ������Ȼ�޷��ر�excel���̡�
'    If Not xlApp Is Nothing Then
'        Shell "cmd.exe /c taskkill /f /im excel.exe"
'    End If
    '    2012-05-23 �ڵ�� ��
    Set xlApp = Nothing
    '2012-05-22 �ڵ�� ��
    
    Exit Sub
errHandler:
'    xlApp.Quit
'    Set xlApp = Nothing
End Sub

'2012-04-19 �ڵ�� ��
'ÿ��ѡ��ͼ�����󣬾�Ҫ���»���
Private Sub combɫ����ʽ_Click()
    If initChart = True Then DrawChart
End Sub

Private Sub combͼ����_Click()
    If initChart = True Then DrawChart
End Sub

Private Sub combͼ����ʽ_Click()
    If initChart = True Then DrawChart
End Sub

Private Sub XAxesTitle_KeyPress(KeyAscii As Integer)
    If initChart = True And KeyAscii = 13 Then DrawChart
End Sub

Private Sub XAxesTitle_LostFocus()
    If initChart = True Then DrawChart
End Sub

Private Sub YAxesTitle_KeyPress(KeyAscii As Integer)
    If initChart = True And KeyAscii = 13 Then DrawChart
End Sub

Private Sub YAxesTitle_LostFocus()
    If initChart = True Then DrawChart
End Sub

Private Sub ChartTitle_KeyPress(KeyAscii As Integer)
    If initChart = True And KeyAscii = 13 Then DrawChart
End Sub

Private Sub ChartTitle_LostFocus()
    If initChart = True Then DrawChart
End Sub
'2012-04-19 �ڵ�� ��

'2012-04-19 �ڵ��
'���û�ͼ��ĸ��������Ȼ��ͼ
Sub DrawChart()
    ccmdDrawCrt.Caption = "������..."
    'picChart.Picture = LoadPicture()
    Clipboard.Clear
    '2012-05-29 �ڵ�� ��
    '�Զ����ͼ����⡢X�ᡢY������
'    If ChartTitle.Text = "ͼ�����" Then ChartTitle.Text = combXAxis.Text & "��" & combYAxis.Text & "ͳ�ƽ��ͼ"
'    If XAxisTitle.Text = "X�����" Then XAxisTitle.Text = combXAxis.Text & "����"
'    If YAxisTitle.Text = "Y�����" Then YAxisTitle.Text = combYAxis.Text
    '2012-05-29 �ڵ�� ��
    ChartTitle.Text = combXAxis.Text & "��" & combYAxis.Text & "ͳ�ƽ��ͼ"
    XAxisTitle.Text = combXAxis.Text & "����"
    YAxisTitle.Text = combYAxis.Text
    
    '2012-05-30 �ڵ�� ��
    '�����������ͳ�ƣ��������y���ǩ�̶�
    If combXAxis.ListIndex = 3 Then
        ChartTitle.Text = combXAxis.Text & "ͳ�ƽ��ͼ"
        YAxisTitle.Text = "����"
    End If
    '2012-05-30 �ڵ�� ��
    
'    xlChart.ClearToMatchStyle

'    xlChart.ActiveChart.ClearToMatchStyle       '2015-12-8 by Ĳ��
    '����ͼ����ʽ
'    xlChart.ActiveChart.ChartType = combͼ����ʽ.ItemData(combͼ����ʽ.ListIndex)   '2015-12-8 by Ĳ��
     xlChart.ChartType = combͼ����ʽ.ItemData(combͼ����ʽ.ListIndex)

    '����ɫ����ʽ
'    xlChart.ActiveChart.ChartStyle = combɫ����ʽ.ItemData(combɫ����ʽ.ListIndex)   '2015-12-8 by Ĳ��
'    xlChart.chartstyle = combɫ����ʽ.ItemData(combɫ����ʽ.ListIndex)

    'xlChart.ClearToMatchStyle
    
    '2012-05-29 �ڵ�� ��
    '�޸��˵�ǰ���е�ͼ����ʽ�벼�ֵĶ�Ӧ��ϵ��ϸ����ÿ��ͼ����
    '����ͼ��X�ᡢY�����
    
    Select Case combͼ����ʽ.ListIndex
    Case 0  '��״ͼ
        '����ͼ����
'        xlChart.applylayout (combͼ����.ItemData(combͼ����.ListIndex))
'
'        Select Case combͼ����.ListIndex + 1 'ͼ�����
'            Case 1, 2, 3, 5, 6, 8, 9, 10
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
'        Select Case combͼ����.ListIndex + 1 'Y�����
'            Case 5, 6, 7, 8, 9
'                xlChart.Axes(xlValue).AxisTitle.Select
'                xlChart.Axes(xlValue, xlPrimary).AxisTitle.Text = YAxisTitle.Text
'        End Select
'        Select Case combͼ����.ListIndex + 1 'X�����
'            Case 7, 8, 9
'                xlChart.Axes(xlCategory).AxisTitle.Select
'                xlChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = XAxisTitle.Text
'        End Select
    Case 1  '����ͼ
        '����ͼ����
'        xlChart.applylayout (combͼ����.ItemData(combͼ����.ListIndex))
'
'        Select Case combͼ����.ListIndex + 1 'ͼ�����
'            Case 1, 2, 3, 5, 6, 8, 9, 10
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
'        Select Case combͼ����.ListIndex + 1 'Y�����
'            Case 1, 5, 6, 7, 10
'                xlChart.Axes(xlValue).AxisTitle.Select
'                xlChart.Axes(xlValue, xlPrimary).AxisTitle.Text = YAxisTitle.Text
'        End Select
'        Select Case combͼ����.ListIndex + 1 'X�����
'            Case 7, 10
'                xlChart.Axes(xlCategory).AxisTitle.Select
'                xlChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = XAxisTitle.Text
'        End Select
    Case 2 '��ͼ
'        Select Case combͼ����.ListIndex + 1 '����ͼ����
'            Case 1, 2, 3, 4, 5, 6, 7
'                xlChart.applylayout (combͼ����.ItemData(combͼ����.ListIndex))
'        End Select
'        Select Case combͼ����.ListIndex + 1 'ͼ�����
'            Case 1, 2, 5, 6, 8, 9, 10, 11
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
    Case 3  '����ͼ
        '����ͼ����
'        If Not (combͼ����.ListIndex + 1 = 11) Then xlChart.applylayout (combͼ����.ItemData(combͼ����.ListIndex))
'
'        Select Case combͼ����.ListIndex + 1 'ͼ�����
'            Case 1, 2, 3, 5, 6, 8, 9, 11
'                xlChart.ChartTitle.Select
'                xlChart.ChartTitle.Text = ChartTitle.Text
'        End Select
'        Select Case combͼ����.ListIndex + 1 'Y�����
'            Case 6, 7, 8
'                xlChart.Axes(xlValue).AxisTitle.Select
'                xlChart.Axes(xlValue, xlPrimary).AxisTitle.Text = YAxisTitle.Text
'        End Select
'        Select Case combͼ����.ListIndex + 1 'X�����
'            Case 7, 8
'                xlChart.Axes(xlCategory).AxisTitle.Select
'                xlChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = XAxisTitle.Text
'        End Select
    End Select
    '2012-05-29 �ڵ�� ��
    
    '������ʾ���
    xlApp.ActiveWindow.Visible = True
    xlChart.ChartArea.Select
    xlChart.ChartArea.Copy
    'Set xlChart = ActiveChart
    picChart.AutoSize = True
'    picChart.Picture = Clipboard.GetData

'��ͼƬ�ŵ�image�У�ͼƬ��С�����壩 2016-1-27 by Ĳ��
    Image1.Picture = Clipboard.GetData
    Image1.Stretch = True

'��ͼƬ�ŵ�frmhuatu�����image��(ͼƬ��Щ�ܿ���) 2016-1-27 by Ĳ��
    frmhuatu.Image1.Picture = Clipboard.GetData
    frmhuatu.Image1.Stretch = True
    frmhuatu.Show
    'picChart.AutoSize = False
    'picChart.PaintPicture picChart.Picture, 0, 0, picChart.ScaleWidth, picChart.ScaleHeight
    ccmdDrawCrt.Caption = "����ͼ��"
    
    If picChart.Picture <> 0 Or cgrdStatic.rows >= 1 Then
        ctlb������.Buttons(1).Enabled = True
        ctlb������.Buttons(3).Enabled = True
        ctlb������.Buttons(4).Enabled = True

        ccmdExportCrt.Enabled = True
    End If
    
End Sub

'2012-04-19 �ڵ��
'��ʼ��ͼ�������list
Sub subInitChartList()
    initChart = False
        
    'ͼ����ʽlist
    combͼ����ʽ.Clear
    combͼ����ʽ.AddItem "��״ͼ": combͼ����ʽ.ItemData(combͼ����ʽ.NewIndex) = 51    'xlColumnClustered
    combͼ����ʽ.AddItem "����ͼ": combͼ����ʽ.ItemData(combͼ����ʽ.NewIndex) = 4     'xlLine
    combͼ����ʽ.AddItem "��״ͼ": combͼ����ʽ.ItemData(combͼ����ʽ.NewIndex) = 69    'xlPieExploded
    combͼ����ʽ.AddItem "����ͼ": combͼ����ʽ.ItemData(combͼ����ʽ.NewIndex) = 57    'xlBarClustered
    combͼ����ʽ.ListIndex = 0
    
    'ɫ����ʽlist
    combɫ����ʽ.Clear
    combɫ����ʽ.AddItem "��ɫ1": combɫ����ʽ.ItemData(combɫ����ʽ.NewIndex) = 2
    combɫ����ʽ.AddItem "��ɫ2": combɫ����ʽ.ItemData(combɫ����ʽ.NewIndex) = 10
    combɫ����ʽ.AddItem "��ɫ3": combɫ����ʽ.ItemData(combɫ����ʽ.NewIndex) = 18
    combɫ����ʽ.AddItem "��ɫ4": combɫ����ʽ.ItemData(combɫ����ʽ.NewIndex) = 26
    combɫ����ʽ.AddItem "��ɫ5": combɫ����ʽ.ItemData(combɫ����ʽ.NewIndex) = 34
    combɫ����ʽ.AddItem "��ɫ6": combɫ����ʽ.ItemData(combɫ����ʽ.NewIndex) = 42
    combɫ����ʽ.ListIndex = 0
     
    'ͼ����list
    combͼ����.Clear
    combͼ����.AddItem "����1": combͼ����.ItemData(combͼ����.NewIndex) = 1
    combͼ����.AddItem "����2": combͼ����.ItemData(combͼ����.NewIndex) = 2
    combͼ����.AddItem "����3": combͼ����.ItemData(combͼ����.NewIndex) = 3
    combͼ����.AddItem "����4": combͼ����.ItemData(combͼ����.NewIndex) = 4
    combͼ����.AddItem "����5": combͼ����.ItemData(combͼ����.NewIndex) = 5
    combͼ����.AddItem "����6": combͼ����.ItemData(combͼ����.NewIndex) = 6
    combͼ����.AddItem "����7": combͼ����.ItemData(combͼ����.NewIndex) = 7
    combͼ����.AddItem "����8": combͼ����.ItemData(combͼ����.NewIndex) = 8
    combͼ����.AddItem "����9": combͼ����.ItemData(combͼ����.NewIndex) = 9
    combͼ����.AddItem "����10": combͼ����.ItemData(combͼ����.NewIndex) = 10
    combͼ����.AddItem "����11": combͼ����.ItemData(combͼ����.NewIndex) = 11
    combͼ����.ListIndex = 0
    
End Sub

'2012-05-29 �ڵ��
'��ʼ��ͳ�Ʋ��գ����ᣩ��ͳ�ƣ����ᣩ���������б�
Sub subInitStaticList()
    'X��ѡ��
    combXAxis.Clear
    combXAxis.AddItem "������": combXAxis.ItemData(combXAxis.NewIndex) = 0
    combXAxis.AddItem "����ҵ": combXAxis.ItemData(combXAxis.NewIndex) = 1
    combXAxis.AddItem "����λ": combXAxis.ItemData(combXAxis.NewIndex) = 2
    combXAxis.AddItem "��������": combXAxis.ItemData(combXAxis.NewIndex) = 3
    '2012-08-08 �ڵ�� ��
    combXAxis.AddItem "��Σ������": combXAxis.ItemData(combXAxis.NewIndex) = 4
    '2012-08-08 �ڵ�� ��
    '2013-03-31 ������ ��
    combXAxis.AddItem "������/����": combXAxis.ItemData(combXAxis.NewIndex) = 5
    '2013-03-31 ������ ��
    combXAxis.ListIndex = 5
    
    'Y��ѡ��
    combYAxis.Clear
    combYAxis.AddItem "�ϸ�����": combYAxis.ItemData(combYAxis.NewIndex) = 0
    combYAxis.AddItem "���ϸ�����": combYAxis.ItemData(combYAxis.NewIndex) = 1
    combYAxis.AddItem "�ϸ���": combYAxis.ItemData(combYAxis.NewIndex) = 2
    combYAxis.AddItem "�޽������": combYAxis.ItemData(combYAxis.NewIndex) = 3
    combYAxis.AddItem "���": combYAxis.ItemData(combYAxis.NewIndex) = 4
    '2013-03-31 ������ ��
    combYAxis.AddItem "����": combYAxis.ItemData(combYAxis.NewIndex) = 5
    '2013-03-31 ������ ��
    combYAxis.ListIndex = 5
    
End Sub

'2012-05-29 �ڵ��
'������ͳ�ƺ�����ͳ�����ݰ������ϸ����������ϸ��������ϸ��ʡ��޽�����������
Sub sub������ͳ��(ByVal XSelected As Integer, ByVal YSelected As Integer)

    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '����ʱ�仮���˶����У���ǰ����ʱ����
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no����, no���� As Integer
    Dim cur���� As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    'ȥ���ظ���ϵͳ���
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP��ֹ.Year - DTP��ʼ.Year) * 12 + (DTP��ֹ.Month - DTP��ʼ.Month) + 1
    
    '��ʼ��ͳ����Ϣ
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP��ֹ.Value - DTP��ʼ.Value < 0 Then MsgBox ("�������ڱ�����ڳ�ʼ����"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "������" & cgrdInfo.rows - 1
'    Label9.Caption = "ϵͳ��Ų��ظ���" & TotalLines
    
    no���� = 0
    Set cur���� = New Collection
    cur����.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        strSQL = "select �ֹ��� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & strSysNo & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("�ֹ���"))
        
        ifAdd = True
        For j = 1 To cur����.Count
            If cur����.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then cur����.Add queryinfo(i, 0)
        
        strSQL = "select * from ְҵ�����_���ҽ��۱� where ����='16' and ϵͳ���='" & strSysNo & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP��ֹ.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub���պϸ�(IIf(IsNull(lobjRec("���ֽ���")), "", lobjRec("���ֽ���")))
            queryinfo(i, 2) = lobjRec("��������")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '�н��ۣ����ǽ��
            
            strSQL = "select * from ְҵ�����_��������Ϣ�� where ϵͳ���='" & strSysNo & "'"
            dasubSetQueryTimeout 6000
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("�շѽ��")), "0", lobjRec("�շѽ��"))
        Else
            no���� = no���� + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '�޽���
            queryinfo(i, 4) = "0"
        End If
        
    Next
    cur����.Remove (1)
    cur����.Add ""  '�޹�������Ϊδ���࣬���ڼ��������
'    Label10.Caption = "�н��������" & (TotalLines - no����)
    
    '����ϸ�������������
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur����.Count + 1) * 4) As Double
    no���� = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP��ʼ.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP��ʼ.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        For j = 0 To cur����.Count - 1
            If queryinfo(i, 0) = cur����.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then no���� = no���� + 1
        Next j
    Next i
        
    '��ʼ��cgrdStatic������ݺͱ�ͷ��ʽ
    With cgrdStatic
        .Clear
        .cols = cur����.Count + IIf(no���� > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To cur����.Count - 1
            .TextMatrix(0, i) = cur����.Item(i)
        Next
        If no���� > 0 Then .TextMatrix(0, i) = "δ����"
        tmpDate = DTP��ʼ.Value
        i = 1
        While tmpDate < DTP��ֹ.Value
            '�޸��ˣ����� 2012.12.10
            '˵�����������ΪCombInterval��ֵ���磺�գ��£�������ȡ�����
'            .TextMatrix(i, 0) = "��" & i & "����"
            .TextMatrix(i, 0) = "��" & i & CombInterval.Text
            '2012.12.11   ����
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '������������cgrdStatic��
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur����.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur����.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur����.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur����.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur����.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    
    dafuncGetData ("update ְҵ�����_ҵ��������Ϣ�� set ������Ŀ='ͳ������_������',����ֵ='ͳ�����_" & combYAxis.Text & "',ö����Դ='" & Xnum & "',˵��='" & Ynum & "' where left(������Ŀ,5)='ͳ������_'")
End Sub

Sub sub����ҵͳ��(ByVal XSelected As Integer, ByVal YSelected As Integer)

    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '����ʱ�仮���˶����У���ǰ����ʱ����
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no��ҵ, no���� As Integer
    Dim cur��ҵ As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    'ȥ���ظ���ϵͳ���
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP��ֹ.Year - DTP��ʼ.Year) * 12 + (DTP��ֹ.Month - DTP��ʼ.Month) + 1
    
    '��ʼ��ͳ����Ϣ
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP��ֹ.Value - DTP��ʼ.Value < 0 Then MsgBox ("�������ڱ�����ڳ�ʼ����"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "������" & cgrdInfo.rows - 1
'    Label9.Caption = "ϵͳ��Ų��ظ���" & TotalLines
    
    
    no���� = 0
    Set cur��ҵ = New Collection
    cur��ҵ.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        strSQL = "select ��λ���� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("��λ����"))
        
        If Not queryinfo(i, 0) = "" Then
            strSQL = "select * from ��λ����_��λ������Ϣ�� where ��λ����='" & queryinfo(i, 0) & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("��ҵ���"))
        End If
        
        ifAdd = True
        For j = 1 To cur��ҵ.Count
            If cur��ҵ.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then cur��ҵ.Add queryinfo(i, 0)
        
        strSQL = "select * from ְҵ�����_���ҽ��۱� where ����='16' and ϵͳ���='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP��ֹ.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub���պϸ�(IIf(IsNull(lobjRec("���ֽ���")), "", lobjRec("���ֽ���")))
            queryinfo(i, 2) = lobjRec("��������")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '�н��ۣ����ǽ��
            
            strSQL = "select * from ְҵ�����_��������Ϣ�� where ϵͳ���='" & strSysNo & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("�շѽ��")), "0", lobjRec("�շѽ��"))
        Else
            no���� = no���� + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '�޽���
            queryinfo(i, 4) = "0"
        End If
    Next
    cur��ҵ.Remove (1)
    cur��ҵ.Add ""  '����ҵ����Ϊδ���࣬���ڼ��������
'    Label10.Caption = "�н��������" & (TotalLines - no����)
    
    '����ϸ�������������
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur��ҵ.Count + 1) * 4) As Double
    no��ҵ = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP��ʼ.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP��ʼ.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        
        For j = 0 To cur��ҵ.Count - 1
            If queryinfo(i, 0) = cur��ҵ.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then no��ҵ = no��ҵ + 1
        Next j
    Next i
        
        
    '��ʼ��cgrdStatic������ݺͱ�ͷ��ʽ
    With cgrdStatic
        .Clear
        .cols = cur��ҵ.Count + IIf(no��ҵ > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To cur��ҵ.Count - 1
            .TextMatrix(0, i) = cur��ҵ.Item(i)
        Next
        If no��ҵ > 0 Then .TextMatrix(0, i) = "δ����"
        tmpDate = DTP��ʼ.Value
        i = 1
        While tmpDate < DTP��ֹ.Value
            '�޸��ˣ����� 2012.12.10
            '˵�����������ΪCombInterval��ֵ���磺�գ��£�������ȡ�����
'            .TextMatrix(i, 0) = "��" & i & "����"
            .TextMatrix(i, 0) = "��" & i & CombInterval.Text
            '2012.12.11   ����
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '������������cgrdStatic��
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��ҵ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��ҵ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��ҵ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��ҵ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��ҵ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    dafuncGetData ("update ְҵ�����_ҵ��������Ϣ�� set ������Ŀ='ͳ������_����ҵ',����ֵ='ͳ�����_" & combYAxis.Text & "',ö����Դ='" & Xnum & "',˵��='" & Ynum & "' where left(������Ŀ,5)='ͳ������_'")
End Sub

'2012-05-30 �ڵ��
'����쵥λͳ��
Sub sub����λͳ��(ByVal XSelected As Integer, ByVal YSelected As Integer)
    
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '����ʱ�仮���˶����У���ǰ����ʱ����
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no��λ, no���� As Integer
    Dim cur��λ As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    'ȥ���ظ���ϵͳ���
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    Month = (DTP��ֹ.Year - DTP��ʼ.Year) * 12 + (DTP��ֹ.Month - DTP��ʼ.Month) + 1
    
    '��ʼ��ͳ����Ϣ
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP��ֹ.Value - DTP��ʼ.Value < 0 Then MsgBox ("�������ڱ�����ڳ�ʼ����"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "������" & cgrdInfo.rows - 1
'    Label9.Caption = "ϵͳ��Ų��ظ���" & TotalLines
    
    
    no���� = 0
    Set cur��λ = New Collection
    cur��λ.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows -1
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 0)
            
        strSQL = "select ��λ���� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("��λ����"))
        
        ifAdd = True
        For j = 1 To cur��λ.Count
            If cur��λ.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then cur��λ.Add queryinfo(i, 0)
        
        strSQL = "select * from ְҵ�����_���ҽ��۱� where ����='16' and ϵͳ���='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP��ֹ.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub���պϸ�(IIf(IsNull(lobjRec("���ֽ���")), "", lobjRec("���ֽ���")))
            queryinfo(i, 2) = lobjRec("��������")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '�н��ۣ����ǽ��
            
            strSQL = "select * from ְҵ�����_��������Ϣ�� where ϵͳ���='" & strSysNo & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("�շѽ��")), "0", lobjRec("�շѽ��"))
        Else
            no���� = no���� + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '�޽���
            queryinfo(i, 4) = "0"
        End If
    Next
    cur��λ.Remove (1)
    cur��λ.Add ""  '�޵�λ����Ϊδ���࣬���ڼ��������
'    Label10.Caption = "�н��������" & (TotalLines - no����)
    
    '����ϸ�������������
    ReDim StaticGrid(1 To Timex + 1, 1 To (cur��λ.Count + 1) * 4) As Double
    no��λ = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP��ʼ.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP��ʼ.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        
        For j = 0 To cur��λ.Count - 1
            If queryinfo(i, 0) = cur��λ.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then no��λ = no��λ + 1
        Next j
    Next i
        
        
    '��ʼ��cgrdStatic������ݺͱ�ͷ��ʽ
    With cgrdStatic
        .Clear
        .cols = cur��λ.Count + IIf(no��λ > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To cur��λ.Count - 1
            .TextMatrix(0, i) = cur��λ.Item(i)
        Next
        If no��λ > 0 Then .TextMatrix(0, i) = "δ����"
        tmpDate = DTP��ʼ.Value
        i = 1
        While tmpDate < DTP��ֹ.Value
            '�޸��ˣ����� 2012.12.10
            '˵�����������ΪCombInterval��ֵ���磺�գ��£�������ȡ�����
'            .TextMatrix(i, 0) = "��" & i & "����"
            .TextMatrix(i, 0) = "��" & i & CombInterval.Text
            '2012.12.11   ����
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '������������cgrdStatic��
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��λ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��λ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��λ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��λ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To cur��λ.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    dafuncGetData ("update ְҵ�����_ҵ��������Ϣ�� set ������Ŀ='ͳ������_����λ',����ֵ='ͳ�����_" & combYAxis.Text & "',ö����Դ='" & Xnum & "',˵��='" & Ynum & "' where left(������Ŀ,5)='ͳ������_'")
End Sub

'2012-05-30 �ڵ��
'��������ͳ��
Sub sub��������ͳ��(ByVal XSelected As Integer, ByVal YSelected As Integer)
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '����ʱ�仮���˶����У���ǰ����ʱ����
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim no����, numClassify As Integer
    Dim Xnum, Ynum, Znum As Integer
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    'ȥ���ظ���ϵͳ���
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP��ֹ.Year - DTP��ʼ.Year) * 12 + (DTP��ֹ.Month - DTP��ʼ.Month) + 1
    
    '��ʼ��ͳ����Ϣ
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP��ֹ.Value - DTP��ʼ.Value < 0 Then MsgBox ("�������ڱ�����ڳ�ʼ����"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "������" & cgrdInfo.rows - 1
'    Label9.Caption = "ϵͳ��Ų��ظ���" & TotalLines
    
    ReDim queryinfo(1 To TotalLines, 1 To 3) As String
    
    no���� = TotalLines
    For i = 1 To TotalLines
        strSysNo = SysNo(i)
        strSQL = "select * from ְҵ�����_���ҽ��۱� where ����='16' and ϵͳ���='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        On Error Resume Next
        queryinfo(i, 1) = "0"
        queryinfo(i, 2) = DTP��ֹ.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub���պϸ�(IIf(IsNull(lobjRec("���ֽ���")), "", lobjRec("���ֽ���")))
            queryinfo(i, 2) = lobjRec("��������")
            If lobjRec("��������").RecordCount = 1 Then
                no���� = no���� - 1
                queryinfo(i, 3) = "1"   '�н��ۣ����ǽ��
            End If
        Else
            queryinfo(i, 1) = ""
            queryinfo(i, 2) = ""
            queryinfo(i, 3) = "0"   '�޽���
        End If
    Next
'    Label10.Caption = "�н��������" & (TotalLines - no����)
    
    '����ϸ�������������
    numClassify = 2 + IIf(no���� > 0, 1, 0) '�������ϸ��������������ϸ������������ܰ������޽��������
    ReDim StaticGrid(1 To Timex + 1, 1 To numClassify) As Double
    For i = 1 To TotalLines
        Month = (Mid(queryinfo(i, 2), 1, 4) - DTP��ʼ.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP��ʼ.Month) + 1
        curTimex = Round(Month / TimeInterval)
        If Round(Month / TimeInterval) < Month / TimeInterval Then
            curTimex = Round(Month / TimeInterval) + 1
        End If
        
        If queryinfo(i, 1) = "1" Then
            StaticGrid(curTimex, 1) = StaticGrid(curTimex, 1) + 1
        ElseIf queryinfo(i, 1) = "0" Then
            StaticGrid(curTimex, 2) = StaticGrid(curTimex, 2) + 1
        Else
            StaticGrid(curTimex, 3) = StaticGrid(curTimex, 3) + 1
        End If
'        If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, 3) = StaticGrid(curTimex, 3) + 1
    Next i
        
    '��ʼ��cgrdStatic������ݺͱ�ͷ��ʽ
    With cgrdStatic
        .Clear
        .cols = numClassify + 1
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        .TextMatrix(0, 1) = "�ϸ�����"
        .TextMatrix(0, 2) = "���ϸ�����"
        If no���� > 0 Then .TextMatrix(0, 3) = "�޽������"
        tmpDate = DTP��ʼ.Value
        i = 1
        While tmpDate < DTP��ֹ.Value
            '�޸��ˣ����� 2012.12.10
            '˵�����������ΪCombInterval��ֵ���磺�գ��£�������ȡ�����
'            .TextMatrix(i, 0) = "��" & i & "����"
            .TextMatrix(i, 0) = "��" & i & CombInterval.Text
            '2012.12.11   ����
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '������������cgrdStatic��
    Xnum = 0
    Ynum = 0
    Znum = 0
    With cgrdStatic
        For i = 1 To Timex
            For j = 1 To numClassify
                .TextMatrix(i, j) = StaticGrid(i, j)
            Next
            Xnum = Xnum + StaticGrid(i, 1)
            Ynum = Ynum + StaticGrid(i, 2)
            If numClassify = 3 Then Znum = Znum + StaticGrid(i, 3)
        Next
        .AutoSize 0, .cols - 1, 0, 0
    End With
    dafuncGetData ("update ְҵ�����_ҵ��������Ϣ�� set ����ֵ='" & Xnum & "',ö����Դ='" & Ynum & "',˵��='" & Znum & "' where ������Ŀ='ͳ������-��������'")
End Sub

'2012-08-08 �ڵ��
'��Σ������ͳ��
Sub sub��Σ������ͳ��(ByVal XSelected As Integer, ByVal YSelected As Integer)
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '����ʱ�仮���˶����У���ǰ����ʱ����
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim noΣ��, no���� As Integer
    Dim curΣ�� As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 4) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j As Integer
    
    Dim Month As String
    
    'ȥ���ظ���ϵͳ���
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
        If flag(i) = True Then
            For j = i + 1 To cgrdInfo.rows - 1
                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
            Next j
        End If
        If flag(i) = True Then SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
    Next i
    
    TotalLines = TotalLines - 1
    
    Month = (DTP��ֹ.Year - DTP��ʼ.Year) * 12 + (DTP��ֹ.Month - DTP��ʼ.Month) + 1
    
    '��ʼ��ͳ����Ϣ
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP��ֹ.Value - DTP��ʼ.Value < 0 Then MsgBox ("�������ڱ�����ڳ�ʼ����"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "������" & cgrdInfo.rows - 1
'    Label9.Caption = "ϵͳ��Ų��ظ���" & TotalLines
    
    
    no���� = 0
    Set curΣ�� = New Collection
    curΣ��.Add ""
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        strSQL = "select Σ������ from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        queryinfo(i, 0) = IIf(lobjRec.RecordCount = 0, "", lobjRec("Σ������"))
        
        ifAdd = True
        For j = 1 To curΣ��.Count
            If curΣ��.Item(j) = queryinfo(i, 0) Then ifAdd = False
        Next j
        If ifAdd Then curΣ��.Add queryinfo(i, 0)
        
        strSQL = "select * from ְҵ�����_���ҽ��۱� where ����='16' and ϵͳ���='" & strSysNo & "'"
        Set lobjRec = dafuncGetData(strSQL)
        'On Error Resume Next
        queryinfo(i, 1) = "1"
        queryinfo(i, 2) = DTP��ֹ.Value
        If lobjRec.RecordCount > 0 Then
            queryinfo(i, 1) = sub���պϸ�(IIf(IsNull(lobjRec("���ֽ���")), "", lobjRec("���ֽ���")))
            queryinfo(i, 2) = lobjRec("��������")
            If lobjRec.RecordCount = 1 Then queryinfo(i, 3) = "1"   '�н��ۣ����ǽ��
            
            strSQL = "select * from ְҵ�����_��������Ϣ�� where ϵͳ���='" & strSysNo & "'"
            Set lobjRec = dafuncGetData(strSQL)
            queryinfo(i, 4) = IIf(IsNull(lobjRec("�շѽ��")), "0", lobjRec("�շѽ��"))
        Else
            no���� = no���� + 1
            queryinfo(i, 2) = "0"
            queryinfo(i, 3) = "0"   '�޽���
            queryinfo(i, 4) = "0"
        End If
    Next
    curΣ��.Remove (1)
    curΣ��.Add ""  '��Σ������Ϊδ���࣬���ڼ��������
'    Label10.Caption = "�н��������" & (TotalLines - no����)
    
    '����ϸ�������������
    ReDim StaticGrid(1 To Timex + 1, 1 To (curΣ��.Count + 1) * 4) As Double
    noΣ�� = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
        If queryinfo(i, 2) = "0" Then
            curTimex = 1
        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP��ʼ.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP��ʼ.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
        End If
        
        For j = 0 To curΣ��.Count - 1
            If queryinfo(i, 0) = curΣ��.Item(j + 1) Then
                If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then StaticGrid(curTimex, j * 4 + 2) = StaticGrid(curTimex, j * 4 + 2) + 1
                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 4 + 1) = StaticGrid(curTimex, j * 4 + 1) + 1
                StaticGrid(curTimex, j * 4 + 3) = StaticGrid(curTimex, j * 4 + 3) + 1
                StaticGrid(curTimex, j * 4 + 4) = StaticGrid(curTimex, j * 4 + 4) + CDbl(queryinfo(i, 4))
            End If
            If queryinfo(i, 0) = "" Then noΣ�� = noΣ�� + 1
        Next j
                
    Next i
        
        
    '��ʼ��cgrdStatic������ݺͱ�ͷ��ʽ
    With cgrdStatic
        .Clear
        .cols = curΣ��.Count + IIf(noΣ�� > 0, 1, 0)
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
        For i = 1 To curΣ��.Count - 1
            .TextMatrix(0, i) = curΣ��.Item(i)
        Next
        If noΣ�� > 0 Then .TextMatrix(0, i) = "δ����"
        tmpDate = DTP��ʼ.Value
        i = 1
        While tmpDate < DTP��ֹ.Value
            '�޸��ˣ����� 2012.12.10
            '˵�����������ΪCombInterval��ֵ���磺�գ��£�������ȡ�����
'            .TextMatrix(i, 0) = "��" & i & "����"
            .TextMatrix(i, 0) = "��" & i & CombInterval.Text
            '2012.12.11   ����
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '������������cgrdStatic��
    Xnum = 0
    Ynum = 0
    Select Case YSelected
    Case 0
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To curΣ��.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 2)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 1
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To curΣ��.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + (StaticGrid(i, 4 * j + 3) - StaticGrid(i, 4 * j + 2) - StaticGrid(i, 4 * j + 1))
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 2
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To curΣ��.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = Format(100# * CDbl(StaticGrid(i, 4 * j + 2)) / IIf(CDbl(StaticGrid(i, 4 * j + 3)) = 0#, 1, CDbl(StaticGrid(i, 4 * j + 3))), "#0.000") & "%"
                    Xnum = Xnum + StaticGrid(i, 4 * j + 2)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 3
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To curΣ��.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 1)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 1)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    Case 4
        With cgrdStatic
            For i = 1 To Timex
                '�޸��ˣ����� 2012.12.10
                '˵�����˴�cur���ֵ������ñ���л�ȡ���������档����
'                For j = 0 To curΣ��.Count - 1
                For j = 0 To cgrdStatic.cols - 2
                '2012.12.10   ����
                    .TextMatrix(i, j + 1) = StaticGrid(i, 4 * j + 4)
                    Xnum = Xnum + StaticGrid(i, 4 * j + 4)
                    Ynum = Ynum + StaticGrid(i, 4 * j + 3)
                Next
            Next
        End With
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    dafuncGetData ("update ְҵ�����_ҵ��������Ϣ�� set ������Ŀ='ͳ������_��Σ��',����ֵ='ͳ�����_" & combYAxis.Text & "',ö����Դ='" & Xnum & "',˵��='" & Ynum & "' where left(������Ŀ,5)='ͳ������_'")

End Sub

'2012-06-06 �ڵ��
'��Ӻ�������ӡͳ�Ʊ���
Sub sub��ӡͳ�Ʊ���(ByVal paraԤ������ As Boolean, ByVal para��ӡ���� As Boolean)
    Dim lcolID As Collection
    Dim para�������� As String
    
    Set lcolID = New Collection
    lcolID.Add "27"
    
    Dim tmplobj As Object
    Set tmplobj = CreateObject("ְҵ������.cls����")
    
    VB.SavePicture picChart.Picture, "C:\ͳ��ͼ��.bmp"
    
    cgrdStatic.PictureType = flexPictureColor
    tmpPicture.AutoSize = True
    Clipboard.Clear
    Clipboard.SetData cgrdStatic.Picture
    tmpPicture.Picture = Clipboard.GetData
    VB.SavePicture tmpPicture.Picture, "C:\ͳ������.bmp"
    
    If combXAxis.ListIndex <> 3 And combYAxis.ListIndex <> 2 And combYAxis.ListIndex <> 4 Then
        para�������� = "ְҵ�����_ͳ�Ʊ���(����)"
    ElseIf combXAxis.ListIndex <> 3 And combYAxis.ListIndex = 2 Then
        para�������� = "ְҵ�����_ͳ�Ʊ���(�ϸ���)"
    ElseIf combXAxis.ListIndex <> 3 And combYAxis.ListIndex = 4 Then
        para�������� = "ְҵ�����_ͳ�Ʊ���(���)"
    ElseIf combXAxis.ListIndex = 3 Then
        para�������� = "ְҵ�����_ͳ�Ʊ���(��������)"
    End If
    
    If paraԤ������ = True Then
        tmplobj.Sub��ӡ���� para��������, lcolID, True, True, "tmp��������", False
    Else
        tmplobj.Sub��ӡ���� para��������, lcolID, False, False, "tmp��������", False
    End If
    
    Set lcolID = Nothing

End Sub

'2013-03-31 ������
'��ʱ������ͳ��
Sub sub��ʱ������ͳ��(ByVal XSelected As Integer, ByVal YSelected As Integer)
    Dim TimeInterval As Integer
    Dim Timex, curTimex As Integer  '����ʱ�仮���˶����У���ǰ����ʱ����
    Dim tmpDate As Date
    Dim TotalLines As Integer
    Dim TotalLinesF As Integer
    Dim lobjRec As Object
    Dim strSQL, strSysNo As String
    Dim noΣ��, no���� As Integer
    Dim curΣ�� As Collection
    Dim ifAdd As Boolean
    Dim Xnum, Ynum As Integer
    ReDim queryinfo(1 To cgrdInfo.rows - 1, 0 To 5) As String
    ReDim SysNo(1 To cgrdInfo.rows - 1) As String
    ReDim SysNoF(1 To cgrdInfo.rows - 1) As String
    ReDim flag(1 To cgrdInfo.rows - 1) As Boolean
    Dim i, j, k As Integer
    
    Dim Month As String
    Dim peopleCount As Double
    'ȥ���ظ���ϵͳ���
    cgrdInfo.Sort = flexSortStringAscending
    TotalLines = 1
    TotalLinesF = 1
    'For i = 1 To cgrdInfo.rows - 1: flag(i) = True: Next
    For i = 1 To cgrdInfo.rows - 1
'        If flag(i) = True Then
'            For j = i + 1 To cgrdInfo.rows - 1
'                If cgrdInfo.TextMatrix(i, 0) = cgrdInfo.TextMatrix(j, 0) Then flag(j) = False
'            Next j
'        End If
        'If flag(i) = True Then
        'If Right(cgrdInfo.TextMatrix(i, 0), 1) <> "F" Then
            SysNo(TotalLines) = cgrdInfo.TextMatrix(i, 0): TotalLines = TotalLines + 1
        'Else
        '    SysNoF(TotalLinesF) = cgrdInfo.TextMatrix(i, 0): TotalLinesF = TotalLinesF + 1
        'End If
    Next i
    
    TotalLines = TotalLines - 1
    TotalLinesF = TotalLinesF - 1
    
    Month = (DTP��ֹ.Year - DTP��ʼ.Year) * 12 + (DTP��ֹ.Month - DTP��ʼ.Month) + 1
    
    '��ʼ��ͳ����Ϣ
    TimeInterval = Val(ctxtInterval.Text) * CombInterval.ItemData(CombInterval.ListIndex)
    If DTP��ֹ.Value - DTP��ʼ.Value < 0 Then MsgBox ("�������ڱ�����ڳ�ʼ����"): Exit Sub

    Timex = Round(Month / TimeInterval)
    If Round(Month / TimeInterval) < Month / TimeInterval Then
        Timex = Round(Month / TimeInterval) + 1
    End If
    Label8.Caption = "������" & cgrdInfo.rows - 1
'    Label9.Caption = "ϵͳ��Ų��ظ���" & TotalLines
    
    
    'no���� = 0
    Set curΣ�� = New Collection
    curΣ��.Add "����"
    curΣ��.Add "����"
    For k = 1 To cgrdInfo.cols
        If cgrdInfo.TextMatrix(0, k) = "�������" Then
            Exit For
        End If
    Next k
    For i = 1 To TotalLines 'cgrdInfo.Rows
        strSysNo = SysNo(i) 'cgrdInfo.TextMatrix(i, 1)
            
        If Right(strSysNo, 1) = "F" Then
            queryinfo(i, 0) = "����"
        Else
            queryinfo(i, 0) = "����"
        End If
        
        If Len(Trim(cgrdInfo.TextMatrix(i, k))) = 10 Then
            queryinfo(i, 2) = cgrdInfo.TextMatrix(i, k)
        Else
            queryinfo(i, 2) = "2013-01-01"
        End If
        
        
    Next
   
    '����ϸ�������������
    'ReDim StaticGrid(1 To Timex + 1, 1 To (curΣ��.Count + 1) * 4) As Double
    ReDim StaticGrid(1 To Timex + 1, 1 To (curΣ��.Count + 1) * 5) As Double
    'noΣ�� = 0
    'peopleCount = 0
    For i = 1 To TotalLines 'cgrdInfo.Rows
'        If queryinfo(i, 2) = "0" Then
'            curTimex = 1
'        Else
            Month = (Mid(queryinfo(i, 2), 1, 4) - DTP��ʼ.Year) * 12 + (Mid(queryinfo(i, 2), 6, 2) - DTP��ʼ.Month) + 1
            curTimex = Round(Month / TimeInterval)
            If Round(Month / TimeInterval) < Month / TimeInterval Then
                curTimex = Round(Month / TimeInterval) + 1
            End If
'        End If
        
        For j = 0 To curΣ��.Count - 1
        'j = 0
            If queryinfo(i, 0) = curΣ��.Item(j + 1) Then
                'If queryinfo(i, 1) = "1" And queryinfo(i, 3) = "1" Then
'                StaticGrid(curTimex, j * 5 + 2) = StaticGrid(curTimex, j * 5 + 2) + 1
'                If queryinfo(i, 3) = "0" Then StaticGrid(curTimex, j * 5 + 1) = StaticGrid(curTimex, j * 5 + 1) + 1
'                StaticGrid(curTimex, j * 5 + 3) = StaticGrid(curTimex, j * 5 + 3) + 1
'                StaticGrid(curTimex, j * 5 + 4) = StaticGrid(curTimex, j * 5 + 4) + CDbl(queryinfo(i, 4))
                'If Right(SysNo(i), 1) = "F" Then
                '    StaticGrid(curTimex, j * 5 + 0) = StaticGrid(curTimex, j * 5 + 0) + 1
                'Else
                    StaticGrid(curTimex, j * 5 + 5) = StaticGrid(curTimex, j * 5 + 5) + 1
                'End If
                
                
            End If
            'If queryinfo(i, 0) = "" Then noΣ�� = noΣ�� + 1
            'If queryinfo(i, 5) = "1" Then peopleCount = peopleCount + 1
        Next j
                
    Next i
        
        
    '��ʼ��cgrdStatic������ݺͱ�ͷ��ʽ
    With cgrdStatic
        .Clear
'        .cols = curΣ��.Count + IIf(noΣ�� > 0, 1, 0)
        .cols = 3
        .rows = Timex + 1
        .FixedCols = 0
        .FixedRows = 0
'        For i = 1 To curΣ��.Count - 1
'            .TextMatrix(0, i) = curΣ��.Item(i)
'        Next
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����"
        'If noΣ�� > 0 Then .TextMatrix(0, i) = "δ����"
        tmpDate = DTP��ʼ.Value
        i = 1
        While tmpDate < DTP��ֹ.Value
            '�޸��ˣ����� 2012.12.10
            '˵�����������ΪCombInterval��ֵ���磺�գ��£�������ȡ�����
'            .TextMatrix(i, 0) = "��" & i & "����"
            .TextMatrix(i, 0) = "��" & i & CombInterval.Text
            '2012.12.11   ����
            tmpDate = DateAdd("m", TimeInterval, tmpDate)
            i = i + 1
        Wend
        .FixedCols = 1
        .FixedRows = 1
    End With
        
    '������������cgrdStatic��
    Xnum = 0
    Ynum = 0
    Select Case YSelected

    '2013-03-31 ������
    Case 5
        With cgrdStatic
            For i = 1 To Timex
                For j = 0 To cgrdStatic.cols - 2
                    .TextMatrix(i, j + 1) = StaticGrid(i, 5 * j + 5)
                    
                    'Xnum = Xnum + StaticGrid(i, 5 * j + 4)
                    'Ynum = Ynum + StaticGrid(i, 5 * j + 3)
                Next
            Next
        End With
    '2013-03-31
    End Select
    cgrdStatic.AutoSize 0, cgrdStatic.cols - 1, 0, 0
    
    dafuncGetData ("update ְҵ�����_ҵ��������Ϣ�� set ������Ŀ='ͳ������_��Σ��',����ֵ='ͳ�����_" & combYAxis.Text & "',ö����Դ='" & Xnum & "',˵��='" & Ynum & "' where left(������Ŀ,5)='ͳ������_'")

End Sub
