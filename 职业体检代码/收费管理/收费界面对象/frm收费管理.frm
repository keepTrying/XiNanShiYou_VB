VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm�շѹ��� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�շѹ���"
   ClientHeight    =   7620
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm�շѹ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Cchk�˷Ѵ�ӡ��ʶ 
      Caption         =   "�˷�ʱ��ӡƱ��"
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   7080
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   6960
      Width           =   10995
      Begin VB.TextBox ctxt�ϼ� 
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
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   120
         Width           =   2415
      End
      Begin VB.Timer ctmr��ʱ 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   5640
         Top             =   120
      End
      Begin VB.TextBox ctxt�˷����� 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   120
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ܽ��"
         Height          =   180
         Index           =   1
         Left            =   7680
         TabIndex        =   17
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "�շ�����"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   0
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10980
      Begin VB.OptionButton coptType 
         Caption         =   "����Ʊ��"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "δ�շѼ�¼"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "�˷Ѽ�¼"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "�շѼ�¼"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
         Height          =   5565
         Left            =   75
         TabIndex        =   1
         Top             =   540
         Width           =   7935
         _cx             =   25376428
         _cy             =   25372248
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
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
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
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
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   5565
         Left            =   8040
         TabIndex        =   22
         Top             =   540
         Width           =   2895
         _cx             =   162862066
         _cy             =   162866776
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
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
         FormatString    =   ""
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
      End
      Begin VB.Label clblMark 
         BackColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   5640
         TabIndex        =   20
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Width           =   540
      End
      Begin VB.Label clab�˷Ѽ�¼�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   9900
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˷Ѵ�����"
         Height          =   180
         Left            =   8880
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label clab��¼�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   8200
         TabIndex        =   11
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շѴ�����"
         Height          =   180
         Left            =   7200
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkԤ�� 
         Caption         =   "��ӡǰԤ��"
         Height          =   255
         Left            =   6960
         TabIndex        =   21
         Top             =   120
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSComCtl2.DTPicker cdtp���� 
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   529
      _Version        =   393216
      Format          =   21299201
      CurrentDate     =   36951
   End
   Begin VB.Menu cmnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu cmnuItemView 
         Caption         =   "��ѯ(&Q)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "ˢ��(&R)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "���ٲ���"
         Index           =   3
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "�˳�(&X)"
         Index           =   5
      End
   End
   Begin VB.Menu cmnuFee 
      Caption         =   "�շ�(&F)"
      Begin VB.Menu cmnuItemFee 
         Caption         =   "����(&N)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "�շ�(&E)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu cmnuItemFee 
         Caption         =   "�޸�"
         Index           =   5
      End
   End
   Begin VB.Menu cmnuBackFee 
      Caption         =   "����"
      Begin VB.Menu cmnuItemBackFee 
         Caption         =   "�˷�(&R)"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu cmnuItemBackFee 
         Caption         =   "����(&B)"
         Index           =   2
      End
   End
   Begin VB.Menu cmnuPrint 
      Caption         =   "��ӡ"
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "Ʊ��"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "��ӡ���ж��ʵ�"
         Index           =   2
      End
   End
   Begin VB.Menu cmenuLocate 
      Caption         =   "��λ"
      Visible         =   0   'False
      Begin VB.Menu cmenuItemLocate 
         Caption         =   "���ٶ�λ"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm�շѹ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

Private mstr��λ��� As String

'��ѯ����
Private mstr�շ����� As String
Private mstr�վݺ� As String
Private mstr��λ���� As String
Private mstr������ As String
Private mstr��ʼ���� As String
Private mstr��ֹ���� As String
Private mstrҵ����� As String

Private mobjQueryResult As Object
Private mcolIndex As Collection

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
Private Sub cgrdMain_Click()
    On Error GoTo errHandler
    sub��ʾһ��������ϸ
    
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "cgrdMain_Click", Err.Number, Err.Description, False)
End Sub

 



'���ܣ��ڷ�����Ϣ�������β��ְ���
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cgrdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
    End If
End Sub


Private Sub cgrdMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbRightButton Then
        PopupMenu cmenuLocate, vbPopupMenuRightButton ', X, Y
    End If
End Sub

Private Sub cmenuItemLocate_Click(Index As Integer)
    Dim lstr�վݺ� As String
    Dim i As Long
    
    lstr�վݺ� = InputBox("�վݺţ�", "���ٶ�λ", "")
    If lstr�վݺ� <> "" Then
        cgrdMain.Row = 0
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.TextMatrix(i, mcolIndex("�վݺ�")) = lstr�վݺ� Then
                cgrdMain.Row = i
                
                Exit For
            End If
        Next
        cgrdMain_Click
    End If
End Sub


Private Sub cmnuItemBackFee_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '�˷�
        '****************�˷���Ϣ����*****************
        With cgrdMain
            If .Row = 0 Then
                MsgBox "��ѡ��Ҫ�˷ѵķ�����Ϣ��", vbInformation, "ϵͳ��ʾ"
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, 0, .Row, 0) = clblMark.BackColor Then
                MsgBox "�ü�¼���˷ѣ�", vbInformation, "ϵͳ��ʾ"
                Exit Sub
            End If
            If MsgBox("���ѵ�λ��" & .TextMatrix(.Row, mcolIndex("���ѵ�λ")) & Chr(13) & Chr(10) & "   �����ˣ�" & .TextMatrix(.Row, mcolIndex("������")) & Chr(13) & Chr(10) & "�շѱ�� ��" & .TextMatrix(.Row, mcolIndex("�շѱ��")) & Chr(13) & Chr(10) & "  �ܽ�� ��" & ctxt�ܽ�� & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "   ��ȷ��Ҫ����ʷ�����", vbYesNo, "ϵͳ��ʾ") = vbYes Then
                pobj�շѹ���.sub�˷� .TextMatrix(.Row, mcolIndex("�շѱ��")), um�û����, Format(Date, "yyyy-mm-dd")
            
                If Cchk�˷Ѵ�ӡ��ʶ.Value <> 1 Then
                    MsgBox "�˷��ѳɹ���", vbOKOnly + vbInformation, "ϵͳ��ʾ"
                End If
          
                If Cchk�˷Ѵ�ӡ��ʶ.Value = 1 Then
                     pobj�շѹ���.sub��ӡ�˷�Ʊ�� .TextMatrix(.Row, mcolIndex("�շѱ��")), IIf(cchkԤ��.Value = 1, True, False)
                End If
                
                'ˢ�½��档
                sub��ѯ����ʾ��¼
                
            End If
        End With
    
    Case 2 '����
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫ�˷ѵķ�����Ϣ��", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If MsgBox("��ȷ��Ҫ����ѡ���еķ�����Ϣ�����ϲ������ָܻ���", vbQuestion + vbYesNo + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
            pobj�շѹ���.sub���Ϸ�����Ϣ cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
            'ˢ�½��档
            sub��ѯ����ʾ��¼
        End If
    End Select
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "cmnuItemBackFee_Click", Err.Number, Err.Description, False)
End Sub

Private Sub cmnuItemFee_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 'ֱ���շ�-����
        frmֱ���շ�.pstr�շѱ�� = ""
        frmֱ���շ�.Show 1, Me
        
        sub��ѯ����ʾ��¼
    Case 2 '�շ�
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫ�շѵķ��ü�¼��", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
        frmֱ���շ�.pstr�շѱ�� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
        frmֱ���շ�.Show 1, Me
        
        sub��ѯ����ʾ��¼
    Case 3 'ɾ��
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫɾ���ķ��ü�¼��", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
        With cgrdMain
            If MsgBox("���ѵ�λ��" & .TextMatrix(.Row, mcolIndex("���ѵ�λ")) & Chr(13) & Chr(10) & "   �����ˣ�" & .TextMatrix(.Row, mcolIndex("������")) & Chr(13) & Chr(10) & "�շ����� ��" & .TextMatrix(.Row, mcolIndex("�շ�����")) & Chr(13) & Chr(10) & "  �ܽ�� ��" & ctxt�ܽ�� & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "   ��ȷ��Ҫɾ����ʷ�����", vbYesNo, "ϵͳ��ʾ") = vbYes Then
                pobj�շѹ���.subɾ�� .TextMatrix(.Row, mcolIndex("�շѱ��"))
            
                'ˢ�½��档
                sub��ѯ����ʾ��¼
                
            End If
        End With
    Case 5 '�޸Ľ��ѷ�ʽ
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵķ��ü�¼��", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
        frm�޸��շ�.pstr�շ����� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
        frm�޸��շ�.Show 1, Me
        If frm�޸��շ�.pblnOk Then
            sub��ѯ����ʾ��¼
        End If
    End Select
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "cmnuItemFee_Click", Err.Number, Err.Description, False)
End Sub

Private Sub cmnuItemPrint_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 'Ʊ��
        '�ж��Ƿ�ѡ�м�¼
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫ��ӡ�ķ�����Ϣ��", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If func¼��Ʊ�ݺ� <> "" Then
            If coptType(1).Value Then
                '�˷�Ʊ�ݡ�
                pobj�շѹ���.sub��ӡ�˷�Ʊ�� cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��")), IIf(cchkԤ��.Value = 1, True, False)
            Else
                
                '�շ�Ʊ�ݡ�
                pobj�շѹ���.sub��ӡƱ�� cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��")), IIf(cchkԤ��.Value = 1, True, False)
                
                sub��ѯ����ʾ��¼
            End If
        End If
    Case 2 '���ж��ʵ�
        frm��ӡ���ж��ʵ�.Show
    End Select
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "cmnuItemPrint_Click", Err.Number, Err.Description, False)
End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '��ѯ
        '�����ѯ����.
        frm��ѯ.pstr�շ����� = mstr�շ�����
        frm��ѯ.pstr�վݺ� = mstr�վݺ�
        frm��ѯ.pstr��λ���� = mstr��λ����
        frm��ѯ.pstr������ = mstr������
        frm��ѯ.pstr��ʼ���� = mstr��ʼ����
        frm��ѯ.pstr��ֹ���� = mstr��ֹ����
        frm��ѯ.pstrҵ����� = mstrҵ�����
        frm��ѯ.Show 1, Me
        If frm��ѯ.pblnOk Then
            
            mstr�շ����� = frm��ѯ.pstr�շ�����
            mstr�վݺ� = frm��ѯ.pstr�վݺ�
            mstr��λ���� = frm��ѯ.pstr��λ����
            mstr������ = frm��ѯ.pstr������
            mstr��ʼ���� = frm��ѯ.pstr��ʼ����
            mstr��ֹ���� = frm��ѯ.pstr��ֹ����
            mstrҵ����� = frm��ѯ.pstrҵ�����
            sub��ѯ����ʾ��¼
        End If
    Case 2 'ˢ��
        sub��ʾ��¼
    Case 3 '���ٶ�λ
        cmenuItemLocate_Click 1
    Case 5
        Unload Me
    End Select

    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "cmnuItemView_Click", Err.Number, Err.Description, False)
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    
    sub��ʾ��¼
    
    'ֻ��δ�շѼ�¼�����շѡ�ɾ��
    ctlb������.Buttons(4).Enabled = coptType(2).Value
    cmnuItemFee(2).Enabled = coptType(2).Value
    cmnuItemFee(3).Enabled = coptType(2).Value
    
    
    'ֻ���շѼ�¼�����˷�,���ϡ�
    cmnuItemBackFee(1).Enabled = coptType(0).Value
    cmnuItemBackFee(2).Enabled = coptType(0).Value
    
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "coptType_Click()", Err.Number, Err.Description, False)
    
End Sub

Private Sub Form_Load()
    If pblnInUse Then Exit Sub
    Dim lcol��������ť As Collection
    Dim lLen As Integer
    On Error GoTo errHandler
    pblnInUse = True                              'ָʾ����������

    '��ʼ��������
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = ctlb������
    Set lcol��������ť = New Collection
    lcol��������ť.Add "��ѯ(&Q)105"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����(&N)102"
    lcol��������ť.Add "�շ�(&F)103"
    lcol��������ť.Add "|"
    lcol��������ť.Add "��ӡ"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����(&T)110"
    lcol��������ť.Add "��ϸ(&D)102"
    lcol��������ť.Add "�˳�"
    mobjGUI.subInitialize lcol��������ť, ""
    If Not umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣɾ��") Then
        cmnuItemFee(3).Visible = False
    End If
    If Not umfuncУ���û�Ȩ��("�շѹ���_ֱ���շ�") Then
        ctlb������.Buttons(3).Visible = False
        ctlb������.Buttons(4).Visible = False
        ctlb������.Buttons(5).Visible = False
        cmnuFee.Visible = False
    End If
    If Not umfuncУ���û�Ȩ��("�շѹ���_�˷�") Then
        cmnuItemBackFee(1).Visible = False
    End If
    If Not umfuncУ���û�Ȩ��("�շѹ���_����") Then
        If cmnuItemBackFee(1).Visible Then
            cmnuItemBackFee(2).Visible = False
        Else
            cmnuBackFee.Visible = False
        End If
    End If
    'If Not umfuncУ���û�Ȩ��("�շѹ���_Ʊ�ݴ�ӡ") Then
        ctlb������.Buttons(6).Visible = False
        ctlb������.Buttons(7).Visible = False
        cmnuPrint.Visible = False
    'End If
    
    'Ĭ����ʾ����������շѼ�¼��
    mstr��ʼ���� = Format(Date, "yyyy-mm-dd")
    mstr��ֹ���� = Format(Date, "yyyy-mm-dd")
    sub��ѯ����ʾ��¼
    
    cmnuItemBackFee(1).Enabled = coptType(0).Value
    cmnuItemBackFee(2).Enabled = coptType(0).Value
    coptType_Click 1
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "Form_Load", Err.Number, Err.Description, False)
End Sub


'���ܣ������������У��ṩ���շ�Ʊ�ݵĴ�ӡ����.
'ʱ��: 2002/02/20
'���ߣ��켽��
Private Sub sub��ӡƱ��()
On Error GoTo errHandle
    
    '�ж��Ƿ�ѡ�м�¼
    If cgrdMain.Row = 0 Then
        MsgBox "��ѡ��Ҫ��ӡ�ķ�����Ϣ��", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    
    pobj�շѹ���.sub��ӡƱ�� cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��")), IIf(cchkԤ��.Value = 1, True, False)

Exit Sub
errHandle:
    sfsub������ "�շѽ��沿��", "frm�˷�", "sub��ӡƱ��", Err.Number, Err.Description, True
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Frame2.Width = Me.ScaleWidth - Frame2.Left - 60
    Frame4.Width = Me.ScaleWidth - Frame4.Left - 60
    Frame4.Top = Me.ScaleHeight - Frame4.Height - 120
    
    Frame2.Height = Frame4.Top - Frame2.Top - 60
    cgrdMain.Width = Frame2.Width * 0.7
    cgrdMain.Height = Frame2.Height - cgrdMain.Top - 60
    
    cgrdDetail.Left = cgrdMain.Left + cgrdMain.Width + 60
    cgrdDetail.Width = Frame2.Width - cgrdDetail.Left - 60
    cgrdDetail.Height = cgrdMain.Height
    
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
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub


Private Sub sub��ѯ����ʾ��¼()
    Dim i As Integer
    Dim lint�˷Ѽ�¼�� As Long
    Dim lobjRec As Object       '������ʱ��¼����
    Dim lInt As Long            '����ѭ������
    
    
    On Error GoTo errHandle
    cgrdDetail.Rows = 1
    
    '��ѯ�շѼ�¼.
    Set mobjQueryResult = pobj�շѹ���.func�շѹ�������ѯ(mstr�շ�����, mstr�վݺ�, mstr��λ����, mstr������, mstr��ʼ����, mstr��ֹ����, mstrҵ�����, lobjRec)
    
    If lobjRec.RecordCount > 0 Then
        lobjRec.MoveFirst
        For lInt = 0 To lobjRec.RecordCount - 1
            If lobjRec("��Ŀ") = "�ܴ���" Then
                clab��¼��.Caption = IIf(IsNull(lobjRec("����")), "0", lobjRec("����"))
            End If
                        
'            If lobjRec("��Ŀ") = "�˷Ѵ���" Then
'                clab�˷Ѽ�¼��.Caption = IIf(IsNull(lobjRec("����")), "0", lobjRec("����"))
'            End If
                        
            lobjRec.MoveNext
        Next
    End If
    
    sub��ʾ��¼
    
    Exit Sub
errHandle:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "sub��ѯ����ʾ��¼()", Err.Number, Err.Description, True)
End Sub

Private Sub sub��ʾ��¼()
    Dim i As Long
    On Error GoTo errHandle
    
    cgrdDetail.Rows = 1
    
    If coptType(0).Value Then
        mobjQueryResult.Filter = "��ʶ=1"
        ctlb������.Buttons(4).Enabled = False
    ElseIf coptType(1).Value Then
        mobjQueryResult.Filter = "��ʶ=2"
    ElseIf coptType(3).Value Then
        mobjQueryResult.Filter = "��ʶ=3"
    Else
        mobjQueryResult.Filter = "��ʶ=0"
    End If
'    mobjQueryResult.Sort = "�շ�����,�շѱ��"
    mobjQueryResult.Sort = "�շѱ��"
    Set cgrdMain.DataSource = mobjQueryResult
    
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.Cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    
    cgrdMain.ColHidden(mcolIndex("�շ�״̬")) = True
    cgrdMain.ColHidden(mcolIndex("��ʶ")) = True
    
    
'    Call subѡ��ͬ������
    If cgrdMain.Row > 0 Then
        sub��ʾһ��������ϸ
    End If
    '��ʾ�ϼơ�
    Dim dblTotal As Double
    For i = 1 To cgrdMain.Rows - 1
        dblTotal = Format(dblTotal + cgrdMain.ValueMatrix(i, mcolIndex("���")), "0.00")
        
        '��ʾ�˷Ѽ�¼����ɫ��
        If cgrdMain.TextMatrix(i, mcolIndex("�շ�״̬")) = 2 Or cgrdMain.TextMatrix(i, mcolIndex("�շ�״̬")) = 3 Then
            cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = clblMark.BackColor
        End If
            
    Next
    ctxt�ϼ� = Format(dblTotal, "0.00")
    cgrdMain.AutoSize 0, cgrdMain.Cols - 1
    Exit Sub
errHandle:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "sub��ʾ��¼()", Err.Number, Err.Description, True)
    
End Sub

Private Sub mobjGUI_Operate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandle
    Select Case Operate
        Case "��ѯ"
            cmnuItemView_Click 1
        Case "����"
            cmnuItemFee_Click 1
        Case "�շ�"
            cmnuItemFee_Click 2
        
        Case "�˷�"
            cmnuItemBackFee_Click 1
        Case "����"
            cmnuItemBackFee_Click 2
        Case "��ϸ"
            If cgrdMain.Row < 1 Then Exit Sub
            frm��ϸ.pNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
            frm��ϸ.Show
        Case "��ӡ"
            cmnuItemPrint_Click 1
        Case "����"
            frm����.Show
    End Select
    Exit Sub
errHandle:
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
    
    lInt = cgrdMain.RowSel
    If lInt > 0 Then
         If cgrdMain.TextMatrix(lInt, mcolIndex("����")) >= 0 Then
            lbln = True
         Else
            lbln = False
         End If
    End If
    With cgrdMain
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
               lcur��� = lcur��� + CCur(.TextMatrix(i, mcolIndex("����"))) * Val(.TextMatrix(i, mcolIndex("����")))
            End If
        Next i
        ctxt�ܽ��.Text = CStr(lcur���) & " Ԫ"
        ctxt�˷�����.Text = .TextMatrix(.Row, 0)
    Else
        ctxt�ܽ��.Text = "0.00 Ԫ"
        ctxt�˷�����.Text = ""
    End If
    End With
    
    
   
    With cgrdMain
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
            If lbln = True Then
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, mcolIndex("����")) >= 0 Then
                    .IsSelected(i) = True
                End If
            Else
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, mcolIndex("����")) < 0 Then
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
               lcur��� = lcur��� + CCur(.TextMatrix(i, mcolIndex("����"))) * Val(.TextMatrix(i, mcolIndex("����")))
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

Private Sub sub��ʾһ��������ϸ()
    Dim lstrNo As String
    Dim lobjRec As Object
    On Error GoTo errHandler
    If cgrdMain.Row < 1 Then Exit Sub
    
    lstrNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
    Set lobjRec = pobj�շѹ���.func��ѯ������ϸ(lstrNo)
    cgrdDetail.FormatString = ""
    Set cgrdDetail.DataSource = lobjRec
    cgrdDetail.AutoResize = True
    cgrdDetail.MergeCol(0) = True
    cgrdDetail.MergeCells = flexMergeRestrictColumns
    ctxt�˷�����.Text = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
    ctxt�ܽ��.Text = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("���"))
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "sub��ʾһ��������ϸ()", Err.Number, Err.Description, True)
End Sub

