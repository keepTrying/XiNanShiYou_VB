VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCareerHstMage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�ܼ��߸�����Ϣ����"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton coptType 
      Caption         =   "δ¼���ܼ��߸�����Ϣ"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton coptType 
      Caption         =   "�����"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton coptType 
      Caption         =   "δ�½���"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "������"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "������"
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   6255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   10095
      _cx             =   2088781198
      _cy             =   2088774425
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
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
   Begin MSComctlLib.Toolbar ctlb������ 
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   953
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ܼ�¼����"
      Height          =   180
      Left            =   8520
      TabIndex        =   8
      Top             =   840
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   9360
      TabIndex        =   7
      Top             =   840
      Width           =   90
   End
End
Attribute VB_Name = "frmCareerHstMage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'���ƣ�ְҵ��ʷ(�ܼ��߸�����Ϣ)�������
'������
'���ܣ�ְҵ��ʷ(�ܼ��߸�����Ϣ)��������Ϲ���������
'���ߣ�Yunle Liu
'ʱ�䣺2012.03
'***************************************

Option Explicit
Public mblninuse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

'��ѯ����
Private mstr��ʼ���� As String
Private mstr��ֹ���� As String
Private mstr�������� As String
Private mstr��λ���� As String
Private mstr���� As String
Private mstr��쵥�� As String
Private mstr�Թܱ�� As String
Private mstrϵͳ��� As String
Private mstr���֤�� As String
'��ѯ���
Private mobjQueryResult As Object

Private mcolIndex As New Collection

'���ܣ����ص�ǰ�����Ƿ��Ѿ����ر�־������ϵͳƽ̨��Ҫ��ġ�
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblninuse
End Property

Private Sub cmnuItemPrint_Click(Index As Integer)
    Dim lcol��� As Collection
    On Error GoTo errHandler
    Set lcol��� = New Collection
    Select Case Index
    Case 1
        '��ӡ����
        lcol���.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
        pobjҵ�����.Sub��ӡ���� "����", lcol���, True
        
    Case 2
        '��ӡ�������
        lcol���.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
        pobjҵ�����.Sub��ӡ���� "�������", lcol���, True
    End Select
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstManage", "cmnuItemPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemRegister_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '����Ǽ�
        '���ʱ�־ = 0 '���ڿ��ƽ���
        frmCareerHstRegt.ctxtsysno = ""
        
        'frmˢ����.Show 1, Me
        frmCareerHstRegt.Show 1, Me
        
        '���²�ѯ��
        sub��ѯ����ʾ
    Case 2 '���Ǽ�
        If cgrdMain.Row >= 1 Then
            'FrmRegister.pstrϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        Else
            'FrmRegister.pstrϵͳ��� = ""
        End If
        'FrmRegister.Show 1, Me
        
        '���²�ѯ��
        sub��ѯ����ʾ
    
    Case 3 '����Ǽ�
        If cgrdMain.Row < 1 Then
            MsgBox "û����Ҫ������ˣ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        'FrmRegisterAgain.pstr��ϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        'FrmRegisterAgain.Show 1, Me
        
    Case 5 '�޸�
        If cgrdMain.Row < 1 Then
            MsgBox "û����Ҫ�޸ĵļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        ���ʼǺ� = 1
        pubϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        'frmCareerHstRegt.ctxtsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        frmCareerHstRegt.Show 1, Me
        '���²�ѯ��
        sub��ѯ����ʾ
    
    Case 6 'ɾ��
        If cgrdMain.Row < 1 Then
            MsgBox "û�п���ɾ���ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        If Not coptType(3) Then
            If MsgBox("��ȷ��Ҫɾ��������¼��һ��ɾ���󽫲��ָܻ���", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
                pobjҵ�����.subɾ�����Ǽ� cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
                    
                cgrdMain.RemoveItem cgrdMain.Row
                clblInfo = cgrdMain.Rows - 1
            End If
        Else
            MsgBox "���������۵ļ�¼������ɾ����", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
    '�޸��ˣ������ 2012-12-3  ��
    '˵�����˳�ʱ���������ѯ
    'bug �� 0000078
     Case 7 '�˳�
        mstr�������� = ""
        mstr��λ���� = ""
        mstr���� = ""
        mstr��쵥�� = ""
        mstr�Թܱ�� = ""
        mstrϵͳ��� = ""
        mstr���֤�� = ""
        
        Unload Me
      ' �����  2012-12-3   ��
    End Select
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstmanage", "cmnuItemRegister_Click", Err.Number, Err.Description, False
End Sub


Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '��ѯ
        With frmQuery
            '��ʾ�ɵĲ�ѯ������
            .pstr��ʼ���� = mstr��ʼ����
            .pstr��ֹ���� = mstr��ֹ����
            .pstr�������� = mstr��������
            .pstr���� = mstr����
            .pstr��λ���� = mstr��λ����
            .pstr��쵥�� = mstr��쵥��
            .pstr�Թܱ�� = mstr�Թܱ��
            .pstr���֤�� = mstr���֤��
            
            '��ȡ�µĲ�ѯ������
            .Show 1, Me
            If .pblnOk Then
                mstr��ʼ���� = .pstr��ʼ����
                mstr��ֹ���� = .pstr��ֹ����
                mstr�������� = .pstr��������
                mstr��λ���� = .pstr��λ����
                mstr���� = .pstr����
                mstr��쵥�� = .pstr��쵥��
                mstr�Թܱ�� = .pstr�Թܱ��
                mstrϵͳ��� = .pstrϵͳ���
                mstr���֤�� = .pstr���֤��
                '���²�ѯ��
                sub��ѯ����ʾ
            End If
        End With
    
    Case 2 'ˢ��
        sub��ʾ��ѯ���
    Case 4
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstmage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub







Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    sub��ʾ��ѯ���
    
    ctlb������.Buttons(4).Enabled = coptType(1).Value
    'cmnuItemRegister(3).Enabled = coptType(1).Value
    ctlb������.Buttons(5).Enabled = coptType(1).Value
    ctlb������.Buttons(6).Enabled = coptType(0).Value Xor coptType(1).Value
    '�޸��ˣ������ 2012-12-3 ��
    '˵�����ж�������в���ʾɾ����ť
    'bug�ţ�0000079
    If coptType(1) = True Then
        ctlb������.Buttons(6).Enabled = False
     End If
    '�޸��ˣ������ 2012-12-3  ��
    ctlb������.Buttons(7).Enabled = coptType(0).Value Xor coptType(1).Value
    
    'cmnuItemRegister(5).Enabled = coptType(0).Value
    'cmnuItemRegister(6).Enabled = coptType(0).Value
    '�޸��ˣ����� 2012.12.05
    'bug�ţ�0000073
    '˵���������½��ۺʹ�����ѡ��ʱ���ܼ��߸�����Ϣ¼�밴ť�����ã�    ����
    If coptType(3).Value = True Or coptType(4).Value = True Then
        ctlb������.Buttons(3).Enabled = False
    Else
        ctlb������.Buttons(3).Enabled = True
    End If
    '2012.12.05     ����
    '2012-04-14 �ڵ�� ��
    '�˵���ɾ������ӡ���������(1)��������(2)�����������
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 �ڵ�� ��
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstmanage", "coptType_Click", Err.Number, Err.Description, False
End Sub





Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblninuse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblninuse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    '�޸ģ�2002-7-1�������ȡ�����۵Ĳ�������Ϊ������ѡ��
    With lcol��������ť
        .Add "��ѯ(&Q)108"
        .Add "|"
        .Add "�ܼ�����Ϣ¼��(&R)102"
        '.Add "����Ǽ�(&R)103"
        .Add "|"
        .Add "�޸�"
        .Add "ɾ��"
        .Add "|"
        .Add "����(&O)113"
        '.Add "|"
        '.Add "��ӡ(&P)107"
        '2012-02-13 �ڵ�� ��
        '�޸����ݣ���ӡ���λ���롱��ť����ݼ�Alt+I��������ͼƬ���110
        '          ��ӡ���ӡ���롱��ť����ݼ�Alt+P��������ͼƬ���107
        '           ��ӡ���ӡ���롱��ť����ݼ�Alt+Q��������ͼƬ���107
        '.Add "��λ����(&I)110"
        '.Add "��ӡ������(&P)107"
        '.Add "���´�����(&Q)107"
        '2012-02-13 �ڵ�� ��
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
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_�ܼ��߸�����Ϣ¼���_ְҵ��ʷ�Ǽ�") = False Then
        ctlb������.Buttons(2).Visible = False
        ctlb������.Buttons(3).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_�ܼ��߸�����Ϣ¼���_�޸�") = False Then
        ctlb������.Buttons(4).Visible = False
        ctlb������.Buttons(5).Visible = False
    End If
'    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_�ܼ��߸�����Ϣ¼���_ɾ��") = False Then
        ctlb������.Buttons(6).Visible = False
        ctlb������.Buttons(7).Visible = False
'    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_�ܼ��߸�����Ϣ¼����_����") = False Then
        ctlb������.Buttons(8).Visible = False
        ctlb������.Buttons(9).Visible = False
    End If
    '2012-06-15 �ڵ�� ��
    '��ӡ����ȡ��
'''    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_�ܼ��߸�����Ϣ¼���_��ӡ") = False Then
'''        ctlb������.Buttons(10).Visible = False
'''        ctlb������.Buttons(11).Visible = False
'''    End If
    '2012-06-15 �ڵ�� ��
    Set lobjTmp = Nothing
    '2012-05-23 ������

    'ȱʡ��ʾ���һ�ܵ������Ա��
    '2012-06-15 �ڵ�� ��
    'Ĭ�ϲ�ѯ����ʱ����������Ա
    mstr��ʼ���� = Format(DateAdd("d", -14, Date), "yyyy-mm-dd")
    '֮ǰ��ʾ��ʱ��Ϊ�̶����޸�Ϊȱʡ����  2015-6-19
    'mstr��ʼ���� = Format(CDate("2012-06-01"), "yyyy-mm-dd")
    '2012-06-15 �ڵ�� ��mstr��ֹ���� = Format(Date, "yyyy-mm-dd")
    mstr�������� = ""
    mstr��λ���� = ""
    mstr���� = ""
    mstr��쵥�� = ""
    mstr�Թܱ�� = ""
    
    sub��ѯ����ʾ
    
    ctlb������.Buttons(4).Enabled = coptType(1).Value
    'cmnuItemRegister(3).Enabled = coptType(1).Value
    ctlb������.Buttons(5).Enabled = coptType(1).Value
    ctlb������.Buttons(6).Enabled = coptType(0).Value
    ctlb������.Buttons(7).Enabled = coptType(0).Value
    'cmnuItemRegister(5).Enabled = coptType(0).Value
    'cmnuItemRegister(6).Enabled = coptType(0).Value

    '2012-04-14 �ڵ�� ��
    '�˵���ɾ������ӡ���������(1)��������(2)�����������
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 �ڵ�� ��

    '2012-02-23 �ڵ�� ��
    cgrdMain.HighLight = flexHighlightWithFocus
    cgrdMain.SelectionMode = flexSelectionListBox
    '2012-02-23 �ڵ�� ��
    ���ʼǺ� = 0

    Exit Sub
errHandler:
   sfsub������ "ְҵ��ʷ¼��", "frmcareerhstmanage", "form_load", Err.Number, Err.Description, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblninuse = False
    
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

Public Sub sub��ѯ����ʾ()
    On Error GoTo errHandler
    '2012-06-15 �ڵ�� ��
    '�������״̬
    
    'Set mobjQueryResult = pobjҵ�����.funcְҵ��ʷ��������ѯ(mstr��ʼ����, mstr��ֹ����, mstr��������, mstr��λ����, mstr����, mstr��쵥��, mstr�Թܱ��, mstrϵͳ���)
    Set mobjQueryResult = pobjҵ�����.func����������ѯ(mstr��ʼ����, mstr��ֹ����, mstr��������, mstr��λ����, mstr����, mstr��쵥��, mstr�Թܱ��, mstrϵͳ���, mstr���֤��)
    '2012-06-15 �ڵ�� ��
    
    sub��ʾ��ѯ���

    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstmage", "sub��ѯ����ʾ", Err.Number, Err.Description, True
End Sub

Private Sub sub��ʾ��ѯ���()
    On Error GoTo errHandler
    Dim lstrWhere As String
    Dim lstrsql As String
    '2012-06-15 �ڵ�� ��
    '�������״̬
'''    If coptType(0).Value Then
'''        mobjQueryResult.Filter = "���״̬='����ʷ¼��'"
'''    ElseIf coptType(1).Value Then
'''        mobjQueryResult.Filter = "���״̬='�����'"
'''        'mobjQueryResult.Filter = "���״̬='���½���' and ����������<>'' and ����ϵͳ���=''"
'''    Else
'''        mobjQueryResult.Filter = "���״̬='���½���'"
'''        'mobjQueryResult.Filter = "(���״̬='���½���' and  ����������='') or (���״̬='���½���' and ����ϵͳ���<>'')"
'''    End If
    If coptType(0).Value Then
        mobjQueryResult.Filter = "���״̬='δ¼���ܼ��߸�����Ϣ'"
    ElseIf coptType(1).Value Then
'        mobjQueryResult.Filter = "���״̬='�����'"
'˵���� �����Ҳ���Բ�¼δ¼���ܼ������Ϣ
'�޸���: ����� 2012 - 12 - 12 �� 2015-6-26  by lanchao ȥ��'δ¼��ȥ���ܼ��߸�����Ϣ'
        'mobjQueryResult.Filter = "���״̬='�����'or ���״̬='δ¼��ȥ���ܼ��߸�����Ϣ'"
        mobjQueryResult.Filter = "���״̬='�����'"
  '�޸���: ����� 2012 - 12 - 12 ��
     
    ElseIf coptType(2).Value Then
        mobjQueryResult.Filter = "���״̬='δ�½���'"
    ElseIf coptType(3).Value Then
'        mobjQueryResult.Filter = "���״̬='���½���'"    '�����½��۸�Ϊ������  2015-11-27 by Ĳ��
        mobjQueryResult.Filter = "���״̬='������'"
    ElseIf coptType(4).Value Then
        'mobjQueryResult.Filter = "���״̬='������'"
    '����״̬
    '�޸��ߣ������ 2012-12-7 ��
       '���˳���������е���
       'Bug�ţ�0000072
    
        Dim lobjRec As Object
        
        lstrsql = "select a.ϵͳ���,����,�Ա�,����,��λ����,�Թܱ��,������ as ������, " _
                    & "convert(varchar(10),�������,120) as �������, " _
                    & "������,isnull(����������,'') as ����������, " _
                    & "isnull(����ϵͳ���,'') as ����ϵͳ���, " _
                    & "���״̬= '������' from ְҵ�����_��������Ϣ�� a inner join  ְҵ�����_�����Ա������Ϣ�� b on a.ϵͳ���=b.ϵͳ��� and ����״̬ = '0'"
        lstrWhere = ""
        If mstr��ʼ���� <> "" Then
            lstrWhere = " and �������>='" & mstr��ʼ���� & "'"
        End If
        If mstr��ֹ���� <> "" Then
            lstrWhere = " and �������<='" & mstr��ֹ���� & "'"
        End If
        If mstr�������� <> "" Then
            lstrWhere = " and a.������='" & mstr�������� & "'"
        End If
        If mstr��λ���� <> "" Then
            lstrWhere = " and ��λ���� like '%" & mstr��λ���� & "%'" & ""
        End If
        If mstr���� <> "" Then
            lstrWhere = " and ���� like '%" & mstr���� & "%'" & ""
        End If
        If mstrϵͳ��� <> "" Then
            lstrWhere = " and a.ϵͳ���='" & mstrϵͳ��� & "'"
        End If
        If lstrWhere <> "" Then
            lstrWhere = " where" & Right(lstrWhere, Len(lstrWhere) - 4)
        End If
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrsql)
        If Not lobjRec.EOF Then
            With cgrdMain
                Set .DataSource = lobjRec
                clblInfo = .Rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .Cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrdMain.Rows > 1 Then
                Dim i As Long
                Set mcolIndex = New Collection
                For i = 0 To cgrdMain.Cols - 1
                    mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
                Next
                cgrdMain.ColHidden(mcolIndex("�Թܱ��")) = True
                cgrdMain.ColHidden(mcolIndex("������")) = True
             End If
            
            Exit Sub
        Else
            cgrdMain.Rows = 1
            Exit Sub
        End If
         '�޸��ߣ������ 2012-12-7 ��
    End If
    '2012-06-15 �ڵ�� ��

    With cgrdMain
        Set .DataSource = mobjQueryResult
        clblInfo = .Rows - 1
        
        .Col = 0
        .Sort = flexSortGenericDescending
        
        '2012-06-15 �ڵ�� ��
        'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
        .AutoSize 0, .Cols - 1, 0, 0
        .ExplorerBar = flexExSort
        .DataMode = flexDMFree
        '2012-05-15 �ڵ�� ��
    End With
         If cgrdMain.Rows > 1 Then
'            Dim i As Long
            Set mcolIndex = New Collection
            For i = 0 To cgrdMain.Cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            cgrdMain.ColHidden(mcolIndex("�Թܱ��")) = True
            cgrdMain.ColHidden(mcolIndex("������")) = True
        End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstmage", "sub��ʾ��ѯ���", Err.Number, Err.Description, True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    ctlb������.Width = Me.ScaleWidth - ctlb������.Left * 2
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Select Case Operate
    '2012-06-15 �ڵ�� ��
    '����ǰѡ��ĳ�������Ա���򽫸������Ա������Ϣ���ݵ�¼�������
    Case "�ܼ�����Ϣ¼��"
        cmnuItemRegister_Click 5
    '2012-06-15 �ڵ�� ��
    Case "��ѯ"
        cmnuItemView_Click 1
        
    Case "ְҵ��ʷ�Ǽ�"
        cmnuItemRegister_Click 1
        Cancel = True
        
    Case "����Ǽ�"
        cmnuItemRegister_Click 3
    
    Case "�޸�"
        Cancel = True
        cmnuItemRegister_Click 5
    
    Case "ɾ��"
        Cancel = True
        cmnuItemRegister_Click 6
    '�޸��ˣ������ 2012-12-3  ��
    '˵�����˳�ʱ���������ѯ
    'bug �� 0000078
     Case "�˳�"
         Cancel = True
         cmnuItemRegister_Click 7
     ' �����  2012-12-3   ��
    Case "����"
        Dim lstrFile As String
        ccmdFile.Filter = "Excel�ļ� (*.xls)|*.xls|�ı��ļ� (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            '2012-04-14 �ڵ�� ��
            '��Ϊ��0�У�Ϊϵͳ��š��������б���ʱΪstring
            cgrdMain.ColDataType(0) = flexDTString
            cgrdMain.SaveGrid lstrFile, flexFileExcel, True   '����excelϵͳ���Ϊ����
            'cgrdMain.SaveGrid lstrFile, flexFileTabText, True
            '2012-04-14 �ڵ���
        End If
        
    '2012-02-13 �ڵ�� ��
    '�޸����ݣ�������ӡ���봰��
    'Case "��ӡ������"
       ' FrmPrintBarCode.Show
    'Case "���´�����"
       ' FrmPrintBarCodeAgain.Show
    '2012-02-13 �ڵ�� ��
    '2012-02-16 ��
    'Case "��λ����"
        'frmImportExcel.Show
    '2012-02-16 �ڵ�� ��
    Case "��ӡ"
        Cancel = True
        If Not cgrdMain.Row > 0 Then Exit Sub
        frmCareerHstRegt.subPrint cgrdMain.TextMatrix(cgrdMain.Row, 0)
    End Select
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstmage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub
