VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm���ݵ��뵼�� 
   BorderStyle     =   0  'None
   Caption         =   "�շ����ݵ��뵼��"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame5 
      Caption         =   "����ѡ��"
      Height          =   645
      Left            =   4785
      TabIndex        =   31
      Top             =   6225
      Width           =   2805
      Begin VB.OptionButton copt�������� 
         Caption         =   "��������"
         Height          =   240
         Left            =   1740
         TabIndex        =   33
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton coptҵ������ 
         Caption         =   "ҵ������"
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   255
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   1440
         X2              =   1440
         Y1              =   120
         Y2              =   615
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   1425
         X2              =   1425
         Y1              =   105
         Y2              =   750
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "��������"
      Height          =   765
      Left            =   4755
      TabIndex        =   29
      Top             =   4710
      Width           =   5685
      Begin VB.CheckBox cchkϵͳ��Ϣ 
         Caption         =   "ϵͳ��Ϣ"
         Enabled         =   0   'False
         Height          =   270
         Left            =   180
         TabIndex        =   30
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "ע�����뵼������Ϣ������Ҫ�ϳ���ʱ��"
         Height          =   240
         Left            =   2235
         TabIndex        =   34
         Top             =   375
         Width           =   3360
      End
   End
   Begin VB.Frame cfra�������� 
      Caption         =   "��������:"
      Height          =   555
      Left            =   4755
      TabIndex        =   27
      Top             =   5655
      Width           =   5670
      Begin MSComctlLib.ProgressBar cprg�������� 
         Height          =   285
         Left            =   75
         TabIndex        =   28
         Top             =   195
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VSFlex6Ctl.vsFlexGrid Cgrd��¼��ʾ 
      Height          =   5985
      Left            =   75
      TabIndex        =   22
      Top             =   945
      Width           =   4575
      _cx             =   4202374
      _cy             =   4204861
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorAlternate=   12640511
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
      Rows            =   27
      Cols            =   5
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
   Begin VB.Frame Frame1 
      Caption         =   "����ѡ��"
      Height          =   630
      Left            =   7710
      TabIndex        =   13
      Top             =   6225
      Width           =   2715
      Begin VB.OptionButton copt����ѡ�� 
         Caption         =   "���ݵ���"
         Height          =   180
         Index           =   1
         Left            =   1500
         TabIndex        =   15
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton copt����ѡ�� 
         Caption         =   "���ݵ���"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   285
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   1335
         X2              =   1335
         Y1              =   105
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1320
         X2              =   1320
         Y1              =   105
         Y2              =   600
      End
   End
   Begin VB.Frame cfraҵ������ 
      Caption         =   "ҵ������"
      Height          =   3750
      Left            =   4725
      TabIndex        =   1
      Top             =   855
      Width           =   5700
      Begin VB.CheckBox cchk��Ŀѡ�� 
         Caption         =   "������Ϣ"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   16
         Top             =   3345
         Width           =   1110
      End
      Begin VB.CheckBox cchk��Ŀѡ�� 
         Caption         =   "������Ϣ"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   1110
      End
      Begin VB.CheckBox cchk��Ŀѡ�� 
         Caption         =   "�շ���Ŀ"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1845
         Width           =   1110
      End
      Begin VB.CheckBox cchk��Ŀѡ�� 
         Caption         =   "�շѱ�׼"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   2145
         Width           =   1110
      End
      Begin VB.CheckBox cchk��Ŀѡ�� 
         Caption         =   "Ʊ�ݸ�ʽ"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   2445
         Width           =   1110
      End
      Begin VB.CheckBox cchk��Ŀѡ�� 
         Caption         =   "�������"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   2745
         Width           =   1110
      End
      Begin VB.CheckBox cchk��Ŀѡ�� 
         Caption         =   "ϵͳ����"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   7
         Top             =   3045
         Width           =   1110
      End
      Begin VB.Frame Frame3 
         Caption         =   "����"
         Height          =   1470
         Index           =   1
         Left            =   1245
         TabIndex        =   2
         Top             =   240
         Width           =   4320
         Begin VB.CheckBox cchk��ʱ���ѯ 
            Height          =   210
            Left            =   1170
            TabIndex        =   26
            Top             =   270
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.TextBox ctxtʱ�� 
            Height          =   300
            Index           =   1
            Left            =   2715
            MaxLength       =   8
            TabIndex        =   19
            Text            =   "00:00:00"
            Top             =   1005
            Width           =   1455
         End
         Begin VB.TextBox ctxtʱ�� 
            Height          =   300
            Index           =   0
            Left            =   945
            MaxLength       =   8
            TabIndex        =   18
            Text            =   "00:00:00"
            Top             =   1005
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker cdtp���� 
            Height          =   300
            Index           =   0
            Left            =   945
            TabIndex        =   3
            Top             =   570
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   36951
         End
         Begin MSComCtl2.DTPicker cdtp���� 
            Height          =   300
            Index           =   1
            Left            =   2715
            TabIndex        =   4
            Top             =   570
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   36951
         End
         Begin VB.Label Label4 
            Caption         =   "��ʱ���ѯ"
            Height          =   225
            Left            =   135
            TabIndex        =   25
            Top             =   285
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "ʱ�䷶Χ"
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   1065
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   2460
            TabIndex        =   20
            Top             =   1065
            Width           =   180
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   2460
            TabIndex        =   6
            Top             =   630
            Width           =   180
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "���ڷ�Χ"
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   630
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar cstu״̬�� 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   7410
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "���ݵ��뵼��"
            TextSave        =   "���ݵ��뵼��"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   3945
      Top             =   7335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Ctlb������ 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin VB.Label fdfsfd 
      Caption         =   "��¼����"
      Height          =   300
      Left            =   90
      TabIndex        =   24
      Top             =   6990
      Width           =   750
   End
   Begin VB.Label clbl��¼�� 
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   930
      TabIndex        =   23
      Top             =   7005
      Width           =   1485
   End
End
Attribute VB_Name = "frm���ݵ��뵼��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************frm���ݵ��뵼��******************************************************************************
'����ʱ�䣺                 2001-3-30
'�����ˣ�                   ����
'�޸�ʱ�䣺
'�޸��ˣ�
'***************************BEGIN*****************************************************************************************
Option Explicit
Public pblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
'���峣��
Private Const ���ݵ��� = 0
Private Const ���ݵ��� = 1
Private Const ������Ϣ = 1
Private Const �շ���Ŀ = 2
Private Const �շѱ�׼ = 3
Private Const Ʊ�ݸ�ʽ = 4
Private Const ������� = 5
Private Const ϵͳ���� = 6
Private Const ������Ϣ = 7
Private Const ��ʼ = 0
Private Const ���� = 1

Private mstr���� As String   '��ѯ�����ַ���

'����򵼳��ļ�¼���������Ŀ����
Private mint�������ܼ�¼�� As Integer
Private mint�����ķ�����Ϣ��¼�� As Integer
Private mint�������շ���Ŀ��¼�� As Integer
Private mint�������շѱ�׼��¼�� As Integer
Private mint������Ʊ�ݸ�ʽ��¼�� As Integer
Private mint�����Ĵ��������¼�� As Integer
Private mint������ϵͳ���ü�¼�� As Integer

Private pobj�շѹ��� As Object
Private pobjҵ������ As Object
Private pobj��λ��λ As Object  '��λ�����ӿ�
Private Const mstrMDBFile = "\�շѹ���2001.mdb"

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


'���ܣ� ѡ���Ƿ�ʱ�䷶Χ����ѯ
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-4-13
Private Sub cchk��ʱ���ѯ_Click()
Dim i As Integer
    For i = 1 To 7
        cchk��Ŀѡ��(i).Value = 0
    Next i
    Cgrd��¼��ʾ.Clear
    Cgrd��¼��ʾ.Rows = 27
    Cgrd��¼��ʾ.Cols = 5
    clbl��¼��.Caption = ""
End Sub

'���ܣ� ѡ��Ҫ����򵼳�����Ŀ��ˢ�±����ʾ������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cchk��Ŀѡ��_Click(Index As Integer)
Dim i As Integer
    If Not func��֤ʱ��(ctxtʱ��(��ʼ)) Or Not func��֤ʱ��(ctxtʱ��(����)) Then '�û������ʱ�䲻��ȷ�����˳�������
        Exit Sub
    End If

    '�����ַ���
    If cchk��ʱ���ѯ Then
        If copt����ѡ��(���ݵ���) Then
            mstr���� = "�������� between '" & cdtp����(��ʼ) & " " & ctxtʱ��(��ʼ) & "' and '" & cdtp����(����) & " " & ctxtʱ��(����) & "'" & " and �շ�״̬='1'"
        Else
            mstr���� = "�������� between #" & cdtp����(��ʼ) & " " & ctxtʱ��(��ʼ) & "# and #" & cdtp����(����) & " " & ctxtʱ��(����) & "#" & " and �շ�״̬='1'"
        End If
    Else
            mstr���� = "�շ�״̬='1'"
    End If
    If Index = ������Ϣ And cchk��Ŀѡ��(Index) Then   'ѡ���ˡ�������Ϣ����
        For i = ������Ϣ To ϵͳ����
            cchk��Ŀѡ��(i).Value = 1
        Next i
        Call sub�����("������Ϣ", copt����ѡ��(���ݵ���), mstr����) '���ѡ����������Ŀ�����ڱ���н���ʾҪ����򵼳��ķ�����Ϣ
        Exit Sub
    End If
    
    If cchk��Ŀѡ��(Index) = 0 Then       '����û�ԭ���Ѿ�ѡ���˸���Ŀ������Ҫȡ�������ˢ�±��
        cchk��Ŀѡ��(������Ϣ) = 0
        Cgrd��¼��ʾ.Clear
        Cgrd��¼��ʾ.Rows = 27
        Cgrd��¼��ʾ.Cols = 5
        clbl��¼��.Caption = ""
        For i = ������Ϣ To ϵͳ����
            If cchk��Ŀѡ��(i) Then
                Call sub�����(cchk��Ŀѡ��(i).Caption, copt����ѡ��(���ݵ���), mstr����)
                Exit Sub
            End If
        Next i
    Else
        Call sub�����(cchk��Ŀѡ��(Index).Caption, copt����ѡ��(���ݵ���), mstr����)
    End If
End Sub




'���ܣ� ��ѯ���������ı䣬��Ļ�ϵ����������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-4-13

Private Sub cdtp����_Change(Index As Integer)
Dim i As Integer
    
    If cchk��ʱ���ѯ = 0 Then Exit Sub
    Cgrd��¼��ʾ.Clear
    Cgrd��¼��ʾ.Rows = 27
    Cgrd��¼��ʾ.Cols = 5
    clbl��¼��.Caption = ""
    For i = 1 To 7
        cchk��Ŀѡ��(i).Value = 0
    Next i
End Sub

'���ܣ� �㡰����ѡ�񡱰�ť
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub copt����ѡ��_Click(Index As Integer)
Dim i As Integer
    If Index = ���ݵ��� Then
          If ctlb������.Buttons.Item(3).Caption = "����(&E)" Then
             For i = 1 To 7
                 cchk��Ŀѡ��(i).Value = 0
             Next i
             Cgrd��¼��ʾ.Clear
             Cgrd��¼��ʾ.Rows = 27
             Cgrd��¼��ʾ.Cols = 5
             clbl��¼��.Caption = ""
          End If
          ctlb������.Buttons.Item(3).Caption = "����(&I)"
    Else
          If ctlb������.Buttons.Item(3).Caption = "����(&I)" Then
             For i = 1 To 7
                 cchk��Ŀѡ��(i).Value = 0
             Next i
             Cgrd��¼��ʾ.Clear
             Cgrd��¼��ʾ.Rows = 27
             Cgrd��¼��ʾ.Cols = 5
             clbl��¼��.Caption = ""
          End If
          ctlb������.Buttons.Item(3).Caption = "����(&E)"
    End If
End Sub

Private Sub copt��������_Click()
Call subEnabled(False)
End Sub

Private Sub coptҵ������_Click()
    Call subEnabled(True)
End Sub

'���ܣ��㹤������ť
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub ctlb������_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim i As Integer
On Error GoTo errhandle
    Select Case Button.Caption
           Case "���(&C)1"
                Cgrd��¼��ʾ.Clear
                Cgrd��¼��ʾ.Rows = 27
                Cgrd��¼��ʾ.Cols = 5
                clbl��¼��.Caption = ""
                For i = 1 To 7
                   cchk��Ŀѡ��(i).Value = 0
                Next i
           Case "����(&I)"
                If coptҵ������ Then
                    Call subBegin
                Else
                If cchkϵͳ��Ϣ Then
                    If MsgBox("�������ݣ�������Ϣ��Ʊ�����ý���գ�Ҫ�����𣿣�", vbInformation + vbYesNo, Mid(ctlb������.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                    Me.Enabled = False
                    'If copt����ѡ��(���ݵ���) Then
                    
                    cprg��������.Max = 1
                    dafuncGetData "exec ϵͳ����_�����������"
                    dasubBeginTran
                    umsub���ݵ��� App.Path & mstrMDBFile, False, cprg��������
                    dasubCommitTran
                    'Else
                    '    umsub���ݵ��� App.Path & mstrMDBFile, cprg��������
                    'End If
                    
                    cprg��������.Value = 0
                    MsgBox "��ɣ�", vbInformation, Left(ctlb������.Buttons.Item(3).Caption, 2)
                    cprg��������.Max = 100
                    Me.Enabled = True
                    
                End If
             End If
                
           Case "����(&E)"
                If coptҵ������ Then
                    Call subBegin
                Else
                If cchkϵͳ��Ϣ Then
                    If MsgBox("�����Ҫ�����𣿣�", vbInformation + vbYesNo, Mid(ctlb������.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                    Me.Enabled = False
                    'If copt����ѡ��(���ݵ���) Then
                    '    umsub���ݵ��� App.Path & mstrMDBFile, True, cprg��������
                    'Else
                    cprg��������.Max = 1
                    umsub���ݵ��� App.Path & mstrMDBFile, cprg��������
                    'End If
                    
                    cprg��������.Value = 0
                    MsgBox "��ɣ�", vbInformation, Left(ctlb������.Buttons.Item(3).Caption, 2)
                    cprg��������.Max = 100
                    Me.Enabled = True
                End If
             End If
           
    End Select
    Exit Sub
errhandle:
    'MsgBox Err.Number & " " & Err.Description
    MsgBox "����ʧ�ܣ�", vbInformation, Left(ctlb������.Buttons.Item(3).Caption, 2)
    cprg��������.Value = 0
    cprg��������.Max = 100
    cfra��������.Caption = "��������:"
    Me.Enabled = True

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
    lsngW = ctlb������.ButtonWidth
    lsngH = ctlb������.ButtonHeight
    lsngSepW = ctlb������.Buttons(2).Width
    
    With cstu״̬��
    If X <= lsngW And Y <= lsngH Then
       .Panels(1).Text = " �������ϵ�����"
    Else
       .Panels(1).Text = ""
    End If
    
    If X <= 2 * lsngW + lsngSepW And X > lsngW + lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "���룺���������ݿ�������ӵ������������������������е����ݵ��������ؿ�"
       If copt����ѡ��(���ݵ���).Value = True Then
       
             ctlb������.Buttons(3).ToolTipText = "����(&I)"
       Else
             ctlb������.Buttons(3).ToolTipText = "����(&E)"
       End If
    End If
    
    If X <= 3 * lsngW + 2 * lsngSepW And X > 2 * lsngW + 2 * lsngSepW And Y <= lsngH Then
       .Panels(1).Text = "�رյ��뵼������"
    End If
    
    End With
End Sub

'���ܣ� ��ѯʱ�������ı䣬����մ����ϵ�����
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-4-13
Private Sub ctxtʱ��_Change(Index As Integer)
Dim i As Integer
    
    If cchk��ʱ���ѯ = 0 Then Exit Sub
    Cgrd��¼��ʾ.Clear
    Cgrd��¼��ʾ.Rows = 27
    Cgrd��¼��ʾ.Cols = 5
    clbl��¼��.Caption = ""
    For i = 1 To 7
        cchk��Ŀѡ��(i).Value = 0
    Next i
End Sub




'���ܣ� �ڽ���ʱ����������ʱ����ֹͨ������Ҽ�����ɾ���Ȳ���
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-4-13

Private Sub ctxtʱ��_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ctxtʱ��(Index).Locked = True
End Sub

'���ܣ� �ڽ���ʱ����������ʱ����ֹͨ������Ҽ�����ɾ���Ȳ���
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-4-13

Private Sub ctxtʱ��_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctxtʱ��(Index).Locked = False
End Sub

'���ܣ� ����Delete����Del��
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-4-13
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    If KeyCode = 46 Or KeyCode = 110 Then
       KeyCode = 0
    End If
    'If Ctlb������.Buttons.Item(3).Caption = "����(&E)" Then
        If Shift = 4 And KeyCode = vbKeyI Then
            'Call subBegin
            If coptҵ������ Then
                Call subBegin
             Else
                If cchkϵͳ��Ϣ Then
                    If copt����ѡ��(���ݵ���) Then
                        If MsgBox("�������ݽ���գ�Ҫ�����𣿣�", vbInformation + vbYesNo, Mid(ctlb������.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                        Me.Enabled = False
                        cprg��������.Max = 1
                        umsub���ݵ��� App.Path & mstrMDBFile, True, cprg��������
                        
                        cprg��������.Value = 0
                        MsgBox "��ɣ�", vbInformation, Left(ctlb������.Buttons.Item(3).Caption, 2)
                        cprg��������.Max = 100
                        Me.Enabled = True
                    End If
                End If
             End If
        ElseIf Shift = 4 And KeyCode = vbKeyE Then
             If coptҵ������ Then
                Call subBegin
             Else
                If cchkϵͳ��Ϣ Then
                    If copt����ѡ��(���ݵ���).Value = False Then
                        If MsgBox("�����Ҫ�����𣿣�", vbInformation + vbYesNo, Mid(ctlb������.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
                        Me.Enabled = False
                        cprg��������.Max = 1
                        umsub���ݵ��� App.Path & mstrMDBFile, cprg��������
                        
                        cprg��������.Value = 0
                        MsgBox "��ɣ�", vbInformation, Left(ctlb������.Buttons.Item(3).Caption, 2)
                        cprg��������.Max = 100
                        Me.Enabled = True
                    End If
                End If
             End If
            
            
            
        End If
    'End If
    Exit Sub
errhandle:
    MsgBox "����ʧ�ܣ�", vbInformation, Left(ctlb������.Buttons.Item(3).Caption, 2)
    cprg��������.Value = 0
    cprg��������.Max = 100
    cfra��������.Caption = "��������:"
    Me.Enabled = True
End Sub

'���ܣ���״̬������ʾ"���ݵ��뵼��"
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cstu״̬��.Panels(1).Text = "���ݵ��뵼��"
End Sub


'���ܣ���ʽ����У���û������ʱ��ֵ
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub ctxtʱ��_KeyPress(Index As Integer, KeyAscii As Integer)
        With ctxtʱ��(Index)
    If (KeyAscii >= 48 And KeyAscii <= 58) Then
        
        Select Case .SelStart
               Case 0
                    If Val(Mid(.Text, 2, 1)) <= 4 Then
                        If KeyAscii > 50 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 49 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 0
                        .SelLength = 1
                        .SetFocus
               Case 1
                    If Val(Mid(.Text, 1, 1)) = 2 Then
                        If KeyAscii > 52 Then
                            KeyAscii = 0
                        End If
                    End If
                    .SelStart = 1
                    .SelLength = 1
                    .SetFocus
               Case 2
                    If Val(Mid(.Text, 5, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 3
                        .SelLength = 1
                        .SetFocus
               Case 3
                    If Val(Mid(.Text, 5, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                    End If
                    .SelStart = 3
                    .SelLength = 1
                    .SetFocus
                   
               Case 4
                    If Val(Mid(.Text, 4, 1)) = 6 Then
                        If KeyAscii > 48 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 4
                        .SelLength = 1
                        .SetFocus
             
               Case 5
                    If Val(Mid(.Text, 8, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                     End If
                        .SelStart = 6
                        .SelLength = 1
                        .SetFocus
                     
               Case 6
                    If Val(Mid(.Text, 8, 1)) = 0 Then
                        If KeyAscii > 54 Then
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii > 53 Then
                            KeyAscii = 0
                        End If
                    End If
                        .SelStart = 6
                        .SelLength = 1
                        .SetFocus
                   
               Case 7
                    If Val(Mid(.Text, 7, 1)) = 6 Then
                        If KeyAscii > 48 Then
                            KeyAscii = 0
                        End If
                    Else
                    End If
                        .SelStart = 7
                        .SelLength = 1
                        .SetFocus
                    
        End Select
        
    ElseIf KeyAscii = 8 And .SelStart > 0 Then
             KeyAscii = 0
            .SelStart = .SelStart - 1
            .SelLength = 0
    Else
        KeyAscii = 0
    End If
       End With
End Sub

'���ܣ�װ�ش��壬��ʼ������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
'�޸ģ�2001-4-26����
Private Sub Form_Load()
    'Dim lobjҵ������ As Object
Dim i As Integer
    If pblnInUse Then Exit Sub

        pblnInUse = True
        Set pobj�շѹ��� = CreateObject("�շ�ҵ�����.cls�շѹ���")
        Set pobjҵ������ = CreateObject("�շ�ҵ�����.clsҵ������")
        Set pobj��λ��λ = CreateObject("��λ����ҵ��.ClsUnitInterface")
   
       Dim lcol��������ť As Collection

     '��ʼ��������
       Set mobjGUI = New cls����ͨ�ö���
       Set mobjGUI.Form = Me
       Set mobjGUI.c������ = ctlb������
       Set lcol��������ť = New Collection
       lcol��������ť.Add "���"
       lcol��������ť.Add "|"
       lcol��������ť.Add "����(&I)112"
       lcol��������ť.Add "|"
       lcol��������ť.Add "�˳�"
       mobjGUI.subInitialize lcol��������ť, ""
       Set lcol��������ť = Nothing
       
       cdtp����(��ʼ) = Date
       cdtp����(����) = Date
       
      'subCopyFile
      
      Dim llngattri As Long   '�ļ�������ֵ
      llngattri = GetFileAttributes(App.Path & "\�շѹ���2001.mdb") '���ļ�������
      If Dir(App.Path & "\�շѹ���2001.mdb") <> "�շѹ���2001.mdb" Then
            MsgBox "δ�ҵ�Ҫ����ġ��շѹ���2001.mdb���ļ��������������ݻ����˳����ԡ�", vbInformation, "���ݵ��뵼��"
            'For i = 1 To 7
            '    cchk��Ŀѡ��(i).Enabled = False
            'Next i
            copt����ѡ��(0).Enabled = False
            copt����ѡ��(1).Value = True
            ctlb������.Buttons(3).Caption = "����(&E)"
            'copt����ѡ��(1).Enabled = False
            'ctlb������.Buttons(3).Enabled = False
      ElseIf llngattri = 33 Or llngattri = 35 Or llngattri = 3 Or llngattri = 1 Then
            MsgBox "��" & App.Path & "\�շѹ���2001.mdb���ļ�Ϊֻ�����ԣ����˳��޸����ԡ�", vbInformation, "���ݵ��뵼��"
            'For i = 1 To 7
            '    cchk��Ŀѡ��(i).Enabled = False
            'Next i
            copt����ѡ��(0).Enabled = False
            copt����ѡ��(1).Value = True
            ctlb������.Buttons(3).Caption = "����(&E)"
            'copt����ѡ��(1).Enabled = False
            'Ctlb������.Buttons(3).Enabled = False
      Else
          'For i = 1 To 7
          '      cchk��Ŀѡ��(i).Enabled = True
          'Next i
          copt����ѡ��(0).Enabled = True
          copt����ѡ��(1).Enabled = True
          ctlb������.Buttons(3).Enabled = True
     End If
End Sub

'���ܣ����ݲ�ѯ�����Ͳ�ѯ��Ŀ����ѯ���ݿ�򱾵����ݿ�
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Function Func�ɼ�����(ByVal Str�ɼ���Ŀ As String, ByVal bln�����־ As Boolean, Optional ByVal Str���� As String) As ADODB.Recordset
Dim lobjTemp As Object
    On Error GoTo errhandle
    Select Case bln�����־
            Case True  '����
                Select Case Str�ɼ���Ŀ
                       Case "������Ϣ"
                            Set lobjTemp = pobj�շѹ���.func��ѯ������Ϣ(Str����)  '�϶���ʱ�䷶Χ������
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�����ķ�����Ϣ��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�����ķ�����Ϣ��¼�� = 0
                                End If
                            Else
                                mint�����ķ�����Ϣ��¼�� = 0
                            End If
                            
                       Case "�շ���Ŀ"
                            Set lobjTemp = pobj�շѹ���.func��ѯ�շ���Ŀ("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�������շ���Ŀ��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�������շ���Ŀ��¼�� = 0
                                End If
                            Else
                                mint�������շ���Ŀ��¼�� = 0
                            End If
                
                            
                       Case "�շѱ�׼"
                            Set lobjTemp = pobjҵ������.func��ѯ�շѱ�׼("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�������շѱ�׼��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�������շѱ�׼��¼�� = 0
                                End If
                            Else
                                mint�������շѱ�׼��¼�� = 0
                            End If
                
                            
                       Case "Ʊ�ݸ�ʽ"
                           Set lobjTemp = pobjҵ������.func��ѯƱ����Ϣ("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint������Ʊ�ݸ�ʽ��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint������Ʊ�ݸ�ʽ��¼�� = 0
                                End If
                            Else
                                mint������Ʊ�ݸ�ʽ��¼�� = 0
                            End If
                       
                    
                       Case "�������"
                            Set lobjTemp = pobjҵ������.func��ѯ������Ϣ("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�����Ĵ��������¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�����Ĵ��������¼�� = 0
                                End If
                            Else
                                mint�����Ĵ��������¼�� = 0
                            End If
                
                       
                       Case "ϵͳ����"
                            Set lobjTemp = pobjҵ������.func��ѯҵ��������Ϣ("")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint������ϵͳ���ü�¼�� = 1
                                Else
                                    mint������ϵͳ���ü�¼�� = 0
                                End If
                            Else
                                mint������ϵͳ���ü�¼�� = 0
                            End If
                                        
                End Select
                
            Case False  '����
                Select Case Str�ɼ���Ŀ
                       Case "������Ϣ"
                            
                            Set lobjTemp = pobj�շѹ���.func��ȡ�ⲿ����("�շѹ���_������Ϣ��", Str����)   '�϶���ʱ�䷶Χ������
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�����ķ�����Ϣ��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�����ķ�����Ϣ��¼�� = 0
                                End If
                            Else
                                mint�����ķ�����Ϣ��¼�� = 0
                            End If
                            
                       Case "�շ���Ŀ"
                            Set lobjTemp = pobj�շѹ���.func��ȡ�ⲿ����("�շѹ���_�շ���Ŀ�ֵ��")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�������շ���Ŀ��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�������շ���Ŀ��¼�� = 0
                                End If
                            Else
                                mint�������շ���Ŀ��¼�� = 0
                            End If
                            
                       Case "�շѱ�׼"
                            Set lobjTemp = pobj�շѹ���.func��ȡ�ⲿ����("�շѹ���_�շѱ�׼��Ϣ��")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�������շѱ�׼��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�������շѱ�׼��¼�� = 0
                                End If
                            Else
                                mint�������շѱ�׼��¼�� = 0
                            End If
                            
                       Case "Ʊ�ݸ�ʽ"
                            Set lobjTemp = pobj�շѹ���.func��ȡ�ⲿ����("�շѹ���_Ʊ��������Ϣ��")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint������Ʊ�ݸ�ʽ��¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint������Ʊ�ݸ�ʽ��¼�� = 0
                                End If
                            Else
                                mint������Ʊ�ݸ�ʽ��¼�� = 0
                            End If
                                                        
                       Case "�������"
                            Set lobjTemp = pobj�շѹ���.func��ȡ�ⲿ����("�շѹ���_������Ϣ��")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint�����Ĵ��������¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint�����Ĵ��������¼�� = 0
                                End If
                            Else
                                mint�����Ĵ��������¼�� = 0
                            End If
                       
                       Case "ϵͳ����"
                            Set lobjTemp = pobj�շѹ���.func��ȡ�ⲿ����("�շѹ���_ҵ�����ñ�")
                            If Not (lobjTemp Is Nothing) Then
                                Set Func�ɼ����� = lobjTemp
                                If Func�ɼ�����.RecordCount > 0 Then
                                    mint������ϵͳ���ü�¼�� = Func�ɼ�����.RecordCount
                                Else
                                    mint������ϵͳ���ü�¼�� = 0
                                End If
                            Else
                                mint������ϵͳ���ü�¼�� = 0
                            End If
                          
               End Select
                 
    End Select
Exit Function
errhandle:
sfsub������ "�շѽ������", "frm���ݵ��뵼��", "Func�ɼ�����", Err.Number, Err.Description, True
End Function

'���ܣ����ݵ���򵼳�����Ŀ�������շѹ��������Ӧ�ĵ��뵼������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub Sub���ݵ��뵼��(ByVal Str��Ŀ As String, ByVal bln�����־ As Boolean, Optional ByVal Str���� As String)
On Error GoTo errhandle
    Select Case bln�����־
           Case False '����
                Select Case Str��Ŀ
                       Case "������Ϣ"
                            pobj�շѹ���.func����������Ϣ Str����
                         
                       Case "�շ���Ŀ"
                            pobj�շѹ���.func�����շ���Ŀ

                       Case "�շѱ�׼"
                            pobjҵ������.func�����շѱ�׼
                            
                       Case "Ʊ�ݸ�ʽ"
                            pobjҵ������.func����Ʊ����Ϣ
                            
                       Case "�������"
                            pobjҵ������.func����������Ϣ
                            
                       Case "ϵͳ����"
                            pobjҵ������.func����ҵ��������Ϣ
                            
                End Select
           Case True '����
                Select Case Str��Ŀ
                       Case "������Ϣ"
                            pobj�շѹ���.func���������Ϣ Str����
                         
                       Case "�շ���Ŀ"
                            pobj�շѹ���.func�����շ���Ŀ
                            
                       Case "�շѱ�׼"
                            pobjҵ������.func�����շѱ�׼
                            
                       Case "Ʊ�ݸ�ʽ"
                            pobjҵ������.func����Ʊ����Ϣ
                            
                       Case "�������"
                            pobjҵ������.func���������Ϣ
                            
                       Case "ϵͳ����"
                            pobjҵ������.func����ҵ��������Ϣ
                            
                End Select
           
    End Select
Exit Sub
errhandle:
    sfsub������ "�շѽ������", "frm���ݵ��뵼��", "Sub���ݵ��뵼��", Err.Number, Err.Description, True
End Sub

'���ܣ���ʼ���뵼����������������ʾ��������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub subBegin()
    'On Error GoTo errhandle
    If Not func��֤ʱ��(ctxtʱ��(��ʼ)) Or Not func��֤ʱ��(ctxtʱ��(����)) Then
        Exit Sub
    End If
    subͳ�Ƽ�¼��
    If mint�������ܼ�¼�� <= 0 Then Exit Sub
    If MsgBox("ȷ��Ҫ������", vbYesNo + vbQuestion, Mid(ctlb������.Buttons(3).Caption, 1, 2)) = vbNo Then Exit Sub
    Me.Enabled = False    '��ʼ����򵼳�����ʱ,��ֹ����������ݡ�
    With cprg��������
        .Max = mint�������ܼ�¼��
        .Min = 0
        .Value = 0
    On Error Resume Next
    Select Case copt����ѡ��(���ݵ���)
           Case False '����ʱ�����ļ�
               subCopyFile
               
               If cchk��ʱ���ѯ Then
                    mstr���� = "�������� between '" & cdtp����(��ʼ) & " " & ctxtʱ��(��ʼ) & "' and '" & cdtp����(����) & " " & ctxtʱ��(����) & "'" & " and �շ�״̬='1'"
               Else
                    mstr���� = "�շ�״̬='1'"
               End If
               If cchk��Ŀѡ��(������Ϣ).Value = 1 Then
                    Sub���ݵ��뵼�� "������Ϣ", False, mstr����
                    If Err.Number = 0 Then
                        .Value = .Value + mint�����ķ�����Ϣ��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�����ķ�����Ϣ��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk��Ŀѡ��(�շ���Ŀ).Value = 1 Then

                    Sub���ݵ��뵼�� "�շ���Ŀ", False
                
                    If Err.Number = 0 Then
                        .Value = .Value + mint�������շ���Ŀ��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�������շ���Ŀ��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk��Ŀѡ��(�շѱ�׼).Value = 1 Then
                    Sub���ݵ��뵼�� "�շѱ�׼", False
                   
                    If Err.Number = 0 Then
                        .Value = .Value + mint�������շѱ�׼��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�������շѱ�׼��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk��Ŀѡ��(Ʊ�ݸ�ʽ).Value = 1 Then
                    Sub���ݵ��뵼�� "Ʊ�ݸ�ʽ", False
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint������Ʊ�ݸ�ʽ��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint������Ʊ�ݸ�ʽ��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk��Ŀѡ��(�������).Value = 1 Then
                    Sub���ݵ��뵼�� "�������", False
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint�����Ĵ��������¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�����Ĵ��������¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk��Ŀѡ��(ϵͳ����).Value = 1 Then
                    Sub���ݵ��뵼�� "ϵͳ����", False
                    .Value = .Value + mint������ϵͳ���ü�¼��
                   
                    cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                    Me.Refresh
                    'Err.Number = 0
               End If
           Case True
               If cchk��ʱ���ѯ Then
                    mstr���� = "�������� between #" & cdtp����(��ʼ) & " " & ctxtʱ��(��ʼ) & "# and #" & cdtp����(����) & " " & ctxtʱ��(����) & "#" & " and �շ�״̬='1'"
               Else
                    mstr���� = "�շ�״̬='1'"
               End If
               If cchk��Ŀѡ��(������Ϣ).Value = 1 Then
                    Sub���ݵ��뵼�� "������Ϣ", True, mstr����
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint�����ķ�����Ϣ��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�����ķ�����Ϣ��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk��Ŀѡ��(�շ���Ŀ) Then
                    Sub���ݵ��뵼�� "�շ���Ŀ", True
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint�������շ���Ŀ��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�������շ���Ŀ��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk��Ŀѡ��(�շѱ�׼).Value = 1 Then
                    Sub���ݵ��뵼�� "�շѱ�׼", True
                   
                   If Err.Number = 0 Then
                        .Value = .Value + mint�������շѱ�׼��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�������շѱ�׼��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk��Ŀѡ��(Ʊ�ݸ�ʽ).Value = 1 Then
                    Sub���ݵ��뵼�� "Ʊ�ݸ�ʽ", True
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint������Ʊ�ݸ�ʽ��¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint������Ʊ�ݸ�ʽ��¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
            
               If cchk��Ŀѡ��(�������).Value = 1 Then
                    Sub���ݵ��뵼�� "�������", True
                    
                    If Err.Number = 0 Then
                        .Value = .Value + mint�����Ĵ��������¼��
                
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                    Else
                        If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - mint�����Ĵ��������¼��
                        cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                        Me.Refresh
                        Err.Number = 0
                    End If
               End If
               
               If cchk��Ŀѡ��(ϵͳ����).Value = 1 Then
               Dim lobjTempIN  As Object
               Dim lobjTempOUT As Object
                    Set lobjTempIN = pobjҵ������.func��ѯҵ��������Ϣ("")
                    Set lobjTempOUT = pobj�շѹ���.func��ȡ�ⲿ����("�շѹ���_ҵ�����ñ�")
                    If Not (lobjTempIN Is Nothing) And Not (lobjTempOUT Is Nothing) Then
                        If lobjTempIN.RecordCount > 0 And lobjTempOUT.RecordCount > 0 Then
                            If lobjTempIN("��Ŀ����") <> lobjTempOUT("��Ŀ����") Then
                                If MsgBox("��Ҫ����Ŀ�Ŀ��������������ݲ������Ƿ������", vbInformation + vbYesNo, "����ҵ��������Ϣ") = vbNo Then
                                    mint������ϵͳ���ü�¼�� = 0
                                    
                                    If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - 1
                                    cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                                    Me.Refresh
                                    GoTo WayOut
                                End If
                            Else
                                Sub���ݵ��뵼�� "ϵͳ����", True
                                .Value = .Value + mint������ϵͳ���ü�¼��
                                    
                                cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                                Me.Refresh
                            End If
                        
                       End If
                   Else
                       mint������ϵͳ���ü�¼�� = 0
                       
                       If mint�������ܼ�¼�� > 0 Then mint�������ܼ�¼�� = mint�������ܼ�¼�� - 1
                       cfra��������.Caption = "��������:��" & .Value & "����¼/�ܹ�" & mint�������ܼ�¼�� & "����¼"
                       Me.Refresh
                   End If
               
               
               End If
    End Select
    End With
WayOut:
    cstu״̬��.Panels(1).Text = "����ɣ�"
    '�ָ�����
    Me.Enabled = True
    'cprg��������.Value = 0
    'cfra��������.Caption = "��������:"
    Set lobjTempIN = Nothing
    Set lobjTempOUT = Nothing
    MsgBox "��ɣ�", vbInformation, Left(ctlb������.Buttons.Item(3).Caption, 2)
    cprg��������.Value = 0
    cfra��������.Caption = "��������:"
    Exit Sub
errhandle:
    'sffuncMsg Err.Description
    'MsgBox "����ʧ�ܣ�", vbInformation, Left(Ctlb������.Buttons.Item(3).Caption, 2)
    cprg��������.Value = 0
    cfra��������.Caption = "��������:"
    Me.Enabled = True
End Sub


'���ܣ���������ʱ��ֵ�Ƿ���ȷ
'���룺��
'�������
'���أ�����һ������ֵ
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Function func��֤ʱ��(strTime As String) As Boolean
    strTime = Trim(strTime)
    If Not IsDate(strTime) Then
        func��֤ʱ�� = False
        Exit Function
    End If
    If InStr(1, strTime, ":") = 0 Then
        func��֤ʱ�� = False
        Exit Function
    Else
        func��֤ʱ�� = True
    End If
    
End Function


'���ܣ� ���ݲ�ѯ�Ľ�������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub sub�����(ByVal Str�ɼ���Ŀ As String, ByVal bln�����־ As Boolean, Optional ByVal Str���� As String)
Dim lrsTemp As ADODB.Recordset
Dim i As Integer
Dim j As Integer

    On Error GoTo errhandle
    Set lrsTemp = Func�ɼ�����(Str�ɼ���Ŀ, bln�����־, Str����)
    If Not (lrsTemp Is Nothing) Then
        If lrsTemp Is Nothing Then Exit Sub
        If lrsTemp.RecordCount <= 0 Then Exit Sub
        Cgrd��¼��ʾ.Clear   '��ձ��
        Cgrd��¼��ʾ.Rows = lrsTemp.RecordCount + 1
        Cgrd��¼��ʾ.Cols = lrsTemp.Fields.Count
        If Cgrd��¼��ʾ.Rows < 27 Then Cgrd��¼��ʾ.Rows = 27
        If Cgrd��¼��ʾ.Cols < 5 Then Cgrd��¼��ʾ.Cols = 5
        '�����
        For i = 0 To lrsTemp.Fields.Count - 1
            Cgrd��¼��ʾ.TextMatrix(0, i) = lrsTemp.Fields(i).Name
        Next i
        lrsTemp.MoveFirst
        For i = 1 To lrsTemp.RecordCount
            For j = 0 To lrsTemp.Fields.Count - 1
                Cgrd��¼��ʾ.TextMatrix(i, j) = IIf(IsNull(lrsTemp(j).Value), "", lrsTemp(j).Value)
            Next j
            lrsTemp.MoveNext
        Next i
        lrsTemp.MoveFirst
        'Set Cgrd��¼��ʾ.DataSource = lrsTemp '�����
         
        clbl��¼��.Caption = lrsTemp.RecordCount
        Set lrsTemp = Nothing
    End If
Exit Sub
errhandle:

End Sub

'���ܣ� ͳ��Ҫ�������ܼ�¼������Ӧ�ļ�¼��
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub subͳ�Ƽ�¼��()
If cchk��Ŀѡ��(������Ϣ).Value = 0 Then mint�����ķ�����Ϣ��¼�� = 0
If cchk��Ŀѡ��(�շ���Ŀ).Value = 0 Then mint�������շ���Ŀ��¼�� = 0
If cchk��Ŀѡ��(�շѱ�׼).Value = 0 Then mint�������շѱ�׼��¼�� = 0
If cchk��Ŀѡ��(Ʊ�ݸ�ʽ).Value = 0 Then mint������Ʊ�ݸ�ʽ��¼�� = 0
If cchk��Ŀѡ��(�������).Value = 0 Then mint�����Ĵ��������¼�� = 0
If cchk��Ŀѡ��(ϵͳ����).Value = 0 Then mint������ϵͳ���ü�¼�� = 0
mint�������ܼ�¼�� = mint�����ķ�����Ϣ��¼�� + mint�������շ���Ŀ��¼�� + mint�������շѱ�׼��¼�� + mint������Ʊ�ݸ�ʽ��¼�� + _
                       mint�����Ĵ��������¼�� + mint������ϵͳ���ü�¼��
End Sub



'���ܣ� �رմ��ڡ�
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub Form_Unload(Cancel As Integer)
    Set mobjGUI = Nothing
    pblnInUse = False
    Set pobj�շѹ��� = Nothing
    Set pobjҵ������ = Nothing
    Set pobj��λ��λ = Nothing
End Sub

'���ܣ���Ӧ�û��������������
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
Dim i As Integer
Select Case Operate
        Case "���"
        
            Cgrd��¼��ʾ.Clear
            Cgrd��¼��ʾ.Rows = 27
            Cgrd��¼��ʾ.Cols = 5
            clbl��¼��.Caption = ""
            
            For i = 1 To 7
               cchk��Ŀѡ��(i).Value = 0
            Next i
            
End Select
End Sub



Private Sub subEnabled(ByVal lbln�Ƿ�ҵ������ As Boolean)
Dim i As Integer
    If lbln�Ƿ�ҵ������ Then
        For i = ������Ϣ To ������Ϣ
            cchk��Ŀѡ��(i).Enabled = True
            cchkϵͳ��Ϣ.Enabled = False
            cchk��ʱ���ѯ.Enabled = True
            cdtp����(��ʼ).Enabled = True
            cdtp����(����).Enabled = True
            ctxtʱ��(��ʼ).Enabled = True
            ctxtʱ��(����).Enabled = True
        Next i
    Else
        For i = ������Ϣ To ������Ϣ
            cchk��Ŀѡ��(i).Enabled = False
            cchkϵͳ��Ϣ.Enabled = True
            cchk��ʱ���ѯ.Enabled = False
            cdtp����(��ʼ).Enabled = False
            cdtp����(����).Enabled = False
            ctxtʱ��(��ʼ).Enabled = False
            ctxtʱ��(����).Enabled = False
        Next i
        
    End If
End Sub

Private Sub subCopyFile()
On Error GoTo errhandle
    Dim lstrFile As String
    lstrFile = Replace(App.Path, "�շѹ���", "�������") & mstrMDBFile
    CopyFile lstrFile, App.Path & mstrMDBFile, 0
    Exit Sub
errhandle:
    
End Sub
