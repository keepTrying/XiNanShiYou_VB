VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportManage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�������"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10725
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton coptType 
      Caption         =   "������������Ա"
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "8023�����Ѵ�ӡ"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "8023����δ��ӡ"
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox cchkAll 
      Caption         =   "ȫѡ "
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.OptionButton coptType 
      Caption         =   "δ��ӡ"
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton coptType 
      Caption         =   "�Ѵ�ӡ"
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
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
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   10095
      _cx             =   17806
      _cy             =   10610
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
      SelectionMode   =   1
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ����ʾ��һ���µ����ݣ�ִ���κβ������빴ѡ�б�Ķ�Ӧ�С�"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   5220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ܼ�¼����"
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   300
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   330
   End
End
Attribute VB_Name = "frmReportManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'���ƣ�ְҵ��������(�������)
'������
'���ܣ�ְҵ��������(�������)�����ӡδ��ӡ���漴��ѯԤ��
'���ߣ������
'ʱ�䣺203.01
'***************************************

Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1


Private mstr��ʼ���� As String
Private mstr��ֹ���� As String
Private mstr�������� As String
Private mstr������ As String
Private mstrϵͳ��� As String
Private mstr������ As String
Private mstr���� As String
Private mstr��λ���� As String

Private mstrState As String '�����ӡ״̬
'��ѯ���
Private mobjQueryResult As Object

Private mcolIndex As New Collection

'���ܣ����ص�ǰ�����Ƿ��Ѿ����ر�־������ϵͳƽ̨��Ҫ��ġ�
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property


'Private Sub cgrdMain_Click()
'     With cgrdMain
'        If .Row < 1 Or .Row > .rows - 1 Then Exit Sub
'        mstrϵͳ��� = .TextMatrix(.Row, mcolIndex("ϵͳ���"))
'        mstr�������� = .TextMatrix(.Row, mcolIndex("��������"))
'        mstr������ = .TextMatrix(.Row, mcolIndex("������"))
'        mstr������ = .TextMatrix(.Row, mcolIndex("������"))
'        mstrState = .TextMatrix(.Row, mcolIndex("����״̬"))
'        mstr������ = .TextMatrix(.Row, mcolIndex("������"))
'    End With
'End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '��ѯ
        mstrQuery = 1
        With frmQuery
            '��ʾ�ɵĲ�ѯ������
            .pstr��ʼ���� = mstr��ʼ����
            .pstr��ֹ���� = mstr��ֹ����
            .pstr�������� = mstr��������
            .pstr���� = mstr����
            .pstr��λ���� = mstr��λ����
            .pstrϵͳ��� = mstrϵͳ���
            '��ȡ�µĲ�ѯ������
            .Show 1, Me
            If .pblnOk Then
                mstr��ʼ���� = .pstr��ʼ����
                mstr��ֹ���� = .pstr��ֹ����
                mstr�������� = .pstr��������
                mstrϵͳ��� = .pstrϵͳ���
                mstr��λ���� = .pstr��λ����
                mstr���� = .pstr����
               
                '���²�ѯ��
                sub��ѯ����ʾ
            End If
        End With
    
    Case 2 'ˢ��
        sub��ѯ����ʾ
    Case 4
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub������ "�������", "frmReportmanage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub

Private Sub cchkAll_Click()
    Dim i As Integer
    If cchkAll.Value = 1 Then
        For i = 1 To cgrdMain.rows - 1
           cgrdMain.Cell(flexcpChecked, i, 0) = flexChecked
        Next i
    Else
        For i = 1 To cgrdMain.rows - 1
           cgrdMain.Cell(flexcpChecked, i, 0) = flexUnchecked
        Next i
    End If
End Sub



'Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Col <> 0 Then
'        Cancel = True
'    Else
''        cgrdMain.CellChecked = flexChecked
'    End If
'End Sub


Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    subClear
    sub��ѯ����ʾ
    Exit Sub
errHandler:
    sfsub������ "�������", "frmReportmanage", "coptType_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    '��ʾ���ȡ�
    frmProcess.proPercent.Max = 4
    frmProcess.Label1.Caption = "���ڳ�ʼ�����棬��ȴ�..."
    frmProcess.proPercent.Value = 1
    frmProcess.Show
    DoEvents
    Me.Enabled = False
    MousePointer = 11
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    With lcol��������ť
        .Add "��ѯ(&Q)108"
        .Add "|"
        .Add "��ӡ(&P)107"
        .Add "|"
        .Add "Ԥ��(&V)102"
        .Add "|"
        .Add "����(&O)110"
        .Add "|"
        .Add "��λ��챨��(&M)107"
        .Add "|"
        .Add "�˳�"
    End With
    frmProcess.proPercent.Value = 2
    DoEvents
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""

'            With cgrdMain
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��������"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
'            .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
'        End With
    'Ĭ��Ϊ�г���һ�¼�¼
    mstr��ֹ���� = Format(Now(), "yyyy-mm-dd")
    mstr��ʼ���� = Format(DateAdd("m", -1, Now()), "yyyy-mm-dd")
    frmProcess.proPercent.Value = 3
    DoEvents
    sub��ѯ����ʾ
    frmProcess.proPercent.Value = 4
    Unload frmProcess
    
'    Exit Sub
errHandler:
    Me.Enabled = True
    MousePointer = 0
    If Err.Number <> 0 Then
        sfsub������ "�������", "frmReportManage", "form_load", Err.Number, Err.Description, False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

Public Sub sub��ѯ����ʾ()
    Dim lobjRec As Object
    Dim strSQL As String
    On Error GoTo errHandler
    If mstr��ֹ���� <> "" And Len(mstr��ֹ����) < 13 Then
        mstr��ֹ���� = mstr��ֹ���� + " 23:59:59"
    End If
    strSQL = "exec ְҵ�����_��ѯ��챨����Ϣ '" & mstr��ʼ���� & "','" & mstr��ֹ���� & "','" & mstr�������� & "','" & mstrϵͳ��� & "','" & mstr��λ���� & "','" & mstr���� & "' "
    dasubSetQueryTimeout 6000
    Set mobjQueryResult = dafuncGetData(strSQL)
    
'    '����8023�жϴ�ӡ�򲻴�ӡ���˴洢����  2016-1-8 by Ĳ��
'     Set lobjRec = dafuncGetData("exec ְҵ�����_��ѯ��챨����Ϣ8023���� '" & mstr��ʼ���� & "','" & mstr��ֹ���� & "','" & mstr�������� & "','" & mstrϵͳ��� & "','" & mstr��λ���� & "','" & mstr���� & "'")
    
    
    If coptType(0).Value Then
       mobjQueryResult.Filter = "����״̬='δ��ӡ'"
    Else
       mobjQueryResult.Filter = "����״̬='�Ѵ�ӡ'"
    End If
'    '����8023���Ӻ�������Ա��Ϣ���ж� 2016-1-8 by Ĳ�� ��
'    If coptType(0).Value Then
'       mobjQueryResult.Filter = "����״̬='δ��ӡ'"
'    ElseIf coptType(1).Value Then
'       mobjQueryResult.Filter = "����״̬='�Ѵ�ӡ'"
''    Set lobjRec = dafuncGetData("exec ְҵ�����_��ѯ��챨����Ϣ8023���� '" & mstr��ʼ���� & "','" & mstr��ֹ���� & "','" & mstr�������� & "','" & mstrϵͳ��� & "','" & mstr��λ���� & "','" & mstr���� & "'")
'    ElseIf coptType(2).Value Then
'     lobjRec.Filter = "����״̬='δ��ӡ'"
'    ElseIf coptType(3).Value Then
'     lobjRec.Filter = "����״̬='�Ѵ�ӡ'"
'     Else
'        Dim lobsy As Object
'        Set lobsy = dafuncGetData("exec ְҵ�����_��ѯ��챨����Ϣ������Ա '" & mstr��ʼ���� & "','" & mstr��ֹ���� & "','" & mstr�������� & "','" & mstrϵͳ��� & "','" & mstr��λ���� & "','" & mstr���� & "'")
''        Set lobjRec = dafuncGetData("exec ְҵ�����_��ѯ��챨����Ϣ8023���� '" & mstr��ʼ���� & "','" & mstr��ֹ���� & "','" & mstr�������� & "','" & mstrϵͳ��� & "','" & mstr��λ���� & "','" & mstr���� & "'")
'        lobsy.Filter = "����״̬='������������Ա'"
'    End If
'     '����8023���Ӻ�������Ա��Ϣ���ж� 2016-1-8 by Ĳ�� ��
    With cgrdMain
        .rows = 1
If coptType(0).Value = True Or coptType(1).Value = True Then
        
        If Not (mobjQueryResult.EOF Or mobjQueryResult.BOF) Then
            Set .DataSource = mobjQueryResult
            
'            .Sort = flexSortGenericDescending
            'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
            .AutoSize 0, .cols - 1, 0, 0
'            .ExplorerBar = flexExSort
'            .DataMode = flexDMFree
            Dim i As Long
            Set mcolIndex = New Collection
            For i = 0 To .cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            .ColHidden(mcolIndex("������")) = True
            For i = 1 To .rows - 1
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            Next i
'            .Editable = True
        End If
        clblInfo.Caption = .rows - 1
        
        
'        '8023��������Ա��Ϣ��ѯ��ʾ����  2016-1-8 by Ĳ�� ��
'ElseIf coptType(2).Value = True Or coptType(3).Value = True Then
'        If Not (lobjRec.EOF Or lobjRec.BOF) Then
'            Set .DataSource = lobjRec
'
''            .Sort = flexSortGenericDescending
'            'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
'            .AutoSize 0, .cols - 1, 0, 0
''            .ExplorerBar = flexExSort
''            .DataMode = flexDMFree
''            Dim i As Long
'            Set mcolIndex = New Collection
'            For i = 0 To .cols - 1
'                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'            Next
'            .ColHidden(mcolIndex("������")) = True
'            For i = 1 To .rows - 1
'                .Cell(flexcpChecked, i, 0) = flexUnchecked
'            Next i
''            .Editable = True
'        End If
'        clblInfo.Caption = .rows - 1
'Else
'        If Not (lobsy.EOF Or lobsy.BOF) Then
'            Set .DataSource = lobsy
'
''            .Sort = flexSortGenericDescending
'            'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
'            .AutoSize 0, .cols - 1, 0, 0
''            .ExplorerBar = flexExSort
''            .DataMode = flexDMFree
''            Dim i As Long
'            Set mcolIndex = New Collection
'            For i = 0 To .cols - 1
'                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'            Next
'            .ColHidden(mcolIndex("������")) = True
'            For i = 1 To .rows - 1
'                .Cell(flexcpChecked, i, 0) = flexUnchecked
'            Next i
''            .Editable = True
'        End If
'        clblInfo.Caption = .rows - 1
'    '8023��������Ա��Ϣ��ѯ��ʾ����  2016-1-8 by Ĳ�� ��
    
End If
    End With
    Exit Sub
errHandler:
    sfsub������ " �������", "frmReportmanage", "sub��ѯ����ʾ", Err.Number, Err.Description, True
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 80
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 80
    ctlb������.Width = Me.ScaleWidth
    
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Select Case Operate

    Case "��ѯ"
        cmnuItemView_Click 1

    Case "��ӡ"
        subPrint False  '�Ƿ�Ԥ��
        Cancel = True
        
    Case "Ԥ��"
     
        subPrint True  '�Ƿ�Ԥ��
        Cancel = True
'         Dim msg As String
'         Cancel = True
'         '         mstr������� = "�������" '����������Ϣʱ�ж����ĸ����巢��������
'             If mstr������ = "" Then
'                    '�ж�word��Ϣ�Ƿ񱣴���
'                    If coptType(0).Value = True Then
'                        '����word��Ϣ������
'                            frmProcess.proPercent.Max = 4
'                            frmProcess.Label1.Caption = "���ڼ��أ���ȴ�..."
'                            frmProcess.proPercent.Value = 0
'                            frmProcess.Show 0, Me
'                            DoEvents
'                             sub�༭word�ĵ� Me, mstrϵͳ���, mstr��������, False
'                    Else
'                        msg = MsgBox("����Ϣû�б���word,��Ҫ������Ϣ�������Ƿ������Ϣ��", vbYesNo + vbDefaultButton1 + vbQuestion, "ϵͳ��ʾ��")
'                            If msg = vbYes Then
'                            '����word��Ϣ������
'                            frmProcess.proPercent.Max = 4
'                            frmProcess.Label1.Caption = "���ڼ��أ���ȴ�..."
'                            frmProcess.proPercent.Value = 0
'                            frmProcess.Show 0, Me
'                            DoEvents
'
'                            sub�༭word�ĵ� Me, mstrϵͳ���, mstr��������, False
'
'                            Else
'                                 Exit Sub
'                            End If
'                    End If
'
'                     Unload frmProcess
'                      '�������״̬��word�Ĵ������ã���Ҫ�����ݿ⹦�����֮���ⲽ������ƣ�
'                    If pstrFilename = "" Then Exit Sub
'                    If coptType(0).Value = True Then pobjҵ�����.funcд�뵥�˵�ǰ���״̬ mstrϵͳ���, 7   '"�ѷ�����"
'
'             Else
'                '��ȡ�����word
'               sub��ȡword�ĵ� Me, mstrϵͳ���, mstr������, False
'            End If
'    Case "�˻�"
'        If cgrdMain.SelectedRows = 0 Or cgrdMain.Row > cgrdMain.rows - 1 Then
'            MsgBox "��ѡ�����ݣ�"
'        Else
'            If MsgBox("ȷ��Ҫ�˻��ܼ�ƣ�", vbYesNo, "ϵͳ��ʾ") = vbYes Then
'               dafuncGetData "update ְҵ�����_��������Ϣ�� set ���״̬='5' where ϵͳ���='" & Trim(cgrdMain.TextMatrix(cgrdMain.Row, 0)) & "'"
'            cgrdMain.RowHidden(cgrdMain.Row) = True
'        End If
'        End If
    Case "����"
        If cgrdMain.rows <= 1 Then
            MsgBox "û�е����ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
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
    Case "�˳�"
         Cancel = True
         subClear
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub������ "�������", "frmReportManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub
Private Sub subClear()
        mstr�������� = ""
        mstrϵͳ��� = ""
        mstr���� = ""
        mstr��λ���� = ""
End Sub
'��WORD����ã����ڱ���WORD�����ݿ�
Public Sub subSave(ByVal paraFile As String, ByVal paraNo As Integer, ByVal para����� As String)
    subSaveDoc paraFile, paraNo, para�����
End Sub

'��ӡ����
Private Sub subPrint(ByVal paraԤ�� As Boolean)


    Dim i As Integer
    Dim lobj���� As Object
    Dim lcolSysNo As Collection
    On Error GoTo errHandler
  
    Set lobj���� = CreateObject("ְҵ������.cls����")
'    sum = 0
    With cgrdMain
        For i = 1 To .rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                Set lcolSysNo = New Collection
                lcolSysNo.Add .TextMatrix(i, 0)

                 
                 lobj����.Sub��ӡ���� "ְҵ�������_" & .TextMatrix(i, mcolIndex("�������")), lcolSysNo, paraԤ��

                If paraԤ�� = False Then
                    dafuncGetData "update ְҵ�����_��������Ϣ�� set ���״̬='7' where ϵͳ���='" & Trim(.TextMatrix(i, 0)) & "'"
                    .RowHidden(i) = True
                End If
            End If
        Next i
        If lcolSysNo.Count < 1 And .rows > 1 Then
            MsgBox "�빴ѡҪ��ӡ��Ԥ��������", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
    End With
errHandler:
   
End Sub


