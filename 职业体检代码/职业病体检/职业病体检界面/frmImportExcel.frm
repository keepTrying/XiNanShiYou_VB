VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportExcel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   7230
   ClientLeft      =   570
   ClientTop       =   900
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Left            =   10080
      Top             =   2040
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdDetails 
      Height          =   4935
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "˫����Ԫ����޸����ݣ��Զ����浽EXCEL��"
      Top             =   2160
      Width           =   10455
      _cx             =   2088781833
      _cy             =   2088772097
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
   Begin VB.CommandButton ccmd��λ��λ 
      Caption         =   "��λ��λ"
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox ctxt��λ���� 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox ccmbTemplate 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox ccmb�������� 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox ccmb��������� 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "�� ��"
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton ccmdImport 
      Caption         =   "ȷ�ϵ���"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   480
      Width           =   900
   End
   Begin MSComDlg.CommonDialog ccdg 
      Left            =   10200
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox ctxtDataPath 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.CommandButton ccmdSelect 
      Caption         =   "ѡ���ļ�"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ע���ļ���ʽ��ȷ�󷽿ɵ��룻ÿ��ֻ�ܵ����ļ�һ�Σ�������������д���ļ��ظ����룬�赽�����������ɾ����"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   9540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��λ���ƣ�"
      Height          =   180
      Left            =   7080
      TabIndex        =   13
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "��ɫ��Ϊ��¼��"
      Height          =   180
      Left            =   7080
      TabIndex        =   11
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   4080
      TabIndex        =   10
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�����Ա���ͣ�"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "�����ļ���Ϣ�� (����˫���޸ģ����޸ĺ���Զ�������ԭ����Excel�ļ�����)"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   6375
   End
End
Attribute VB_Name = "frmImportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'2012-02-29 �ڵ�� �����������
'1���ļ����룬ֻ�������Ա������Ϣ����������Ϣ�������������Ŀ���ĵ��롣
'2�����Ĳ������Ƶ��ǣ��Ϸ��Լ�顢������Ϣ���桢��λ��Ϣ����ʱֻ�����˵�λ���ƣ���
'3�������ǵ�����Ϣ���棬ȱ�������Ŀ��ʼ�����Թܱ�ŵ����ݣ��ȴ�����ģ��ľ����趨���ò��ֵĲ��������
'4����λ��Ϣ�������������֮��Ĳ�������λ�����������������趨��
'������Ե��������ݣ�Ҳ�����ֶ����롣ֻ�ǣ�������ó���λ����֮�����Ϣ����ò�Ҫ���ֶ����롣
'5�������������������ͣ���ȫCopy���Ǽǲ��ֵ��趨���ǲ��ָ��ĵĻ�������Ҳ��Ҫ�޸ġ�
Private mobj���, mobj����ģ�� As Object
Private lobj������� As Object
Private lobj�������� As Object
Private indextmp As Integer         '��ǡ�������ݺ��롱����һ�У��ڱ���excel�ļ�ʱ��ǿ�����ø��еĸ�ʽΪ�ַ���
Private lbol�Ѿ����� As Boolean     '����ļ��Ƿ��Ѿ������һ�Ρ�����Ϊ�����޷����������ظ�����ĸ��µ����⣬�����ļ���ʽ�Ϸ���ֻ�ܵ���һ�Ρ������򷵻ع�������޸ġ���
Private mcol�����Ŀ As New Collection  '��ѡ��������Ŀ
Private mcol�շ���Ŀ As New Collection  '��ѡ����շ���Ŀ

'ѡ������ģ�������б�
Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    MousePointer = 11
    subChangeTemplate       'ѡ������
    MousePointer = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmImportExcel", "ccmbTemplate_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


'ȫ��Ϊ�������𡱣��ֵ����ݣ��ϸ�ǰ���ڸ��ڼ䡢���ʱ��Ӧ�����
Private Sub Ccmb��������_Click()
    Dim lobj����ģ�弯 As Object
    Dim lcolInfo As Collection
    Dim lcol������ As Collection
    Dim i As Integer
    On Error GoTo errHandler
    
    '�����еķǸ�������ģ����뵽���������б���С��ټ�������������
    ccmbTemplate.Clear
    Set lobj����ģ�弯 = CreateObject("ְҵ������.ClsMedicalExamTemplateSet")
    
    
    '---------------
    '���� 2012-04-01 �޸ľ�Ϊ�����ľ䣨ע�����䣬������䣩
    
    'lobj����ģ�弯.�������� = 3
    'lobj����ģ�弯.������� = ccmb��������.ItemData(ccmb��������.ListIndex)
    lobj����ģ�弯.�������� = Trim(ccmb���������.Text)
    lobj����ģ�弯.������� = Trim(Ccmb��������.Text)
    
    
    
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    Set lcol������ = lobj����ģ�弯.������Ԫ�ؼ�
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
        ccmbTemplate.ItemData(ccmbTemplate.NewIndex) = lcol������(i)
    Next
    ccmbTemplate.Text = ccmbTemplate.List(0)
    
    Set lobj����ģ�弯 = Nothing
    Set lcolInfo = Nothing
    Call ccmbTemplate_Click
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmImportExcel", "ccmb��������_click", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'ȫ��Ϊ�������Ա���͡����ֵ����ݣ�ְҵ���������乤��
Private Sub ccmb���������_Click()
    On Error GoTo errHandler
    Set lobj������� = CreateObject("ְҵ������.clsmedicalexam")
    lobj�������.������� = ccmb���������.ItemData(ccmb���������.ListIndex)
    'Call Ccmb��������_Click
    'sub������������List
    ccmbTemplate.Text = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmImportExcel", "ccmb���������_Click", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

Private Sub ccmdExit_Click()
    Unload Me
    Set mobj��� = Nothing
    Set mobj����ģ�� = Nothing
    Set lobj������� = Nothing
    Set lobj�������� = Nothing
    Set frmImportExcel = Nothing
    sub���¹�������ѯ
End Sub

Private Sub ccmdImport_Click()
    Dim lcolTmp As Collection
    Dim lobjTmp As Object
    Dim i, j As Integer
    Dim lintRows As Integer
    On Error GoTo errHandler
    
    lintRows = cgrdDetails.rows - 2
    If lintRows < 1 Then
        MsgBox "û�пɵ������Ϣ����鿴EXCEL��", vbExclamation, "ϵͳ��ʾ"
        Exit Sub
    End If
    If Trim(ccmb���������.Text) = "" Then
        MsgBox "�����Ա���Ͳ���Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
        Exit Sub
    End If
    If Trim(Ccmb��������.Text) = "" Then
        MsgBox "��������Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
        Exit Sub
    End If
    If Trim(ccmbTemplate.Text) = "" Then
        MsgBox "������Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
        Exit Sub
    End If
    If Trim(ctxt��λ����.Text) = "" Then
        MsgBox "��λ���Ʋ���Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
        Exit Sub
    End If
    Me.Enabled = False
    MousePointer = 11
    '��ʾ���ȡ�
    frmProcess.proPercent.Max = lintRows
    frmProcess.Label1.Caption = "���ڵ��룬��ȴ�..."
    frmProcess.proPercent.Value = 0
    frmProcess.Show 0, Me
    DoEvents
    
    Set lobjTmp = CreateObject("ְҵ������.clsmedicalexam")
    For i = 2 To cgrdDetails.rows - 1
        Set lcolTmp = New Collection
        For j = 0 To cgrdDetails.cols - 1
            If cgrdDetails.TextMatrix(0, j) = "" Then GoTo NEXT_j
            If cgrdDetails.TextMatrix(0, j) = "����" And cgrdDetails.TextMatrix(i, j) = "" Then
                GoTo NEXT_i
            End If
            lcolTmp.Add cgrdDetails.TextMatrix(i, j), cgrdDetails.TextMatrix(0, j)
NEXT_j: Next j
        lcolTmp.Add lobjTmp.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1), "ϵͳ���"
        sub�����ļ����� lcolTmp
        frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        cgrdDetails.RowHidden(i) = True
NEXT_i: Next i                   '��ģ����c���������continue��䣬�����д����goto
    Unload frmProcess
    
    
'    Exit Sub
errHandler:
    ccmdImport.Enabled = False
    ccmdSelect.Enabled = False
    Me.Enabled = True
    MousePointer = 0
    If Err.Number <> 0 Then
        MsgBox ("��������ʧ��!"), vbExclamation, "ϵͳ��ʾ"
    Else
        MsgBox ("����ɹ���"), vbInformation, "ϵͳ��ʾ"
        lbol�Ѿ����� = True
        frmRegisterManage.sub��ѯ����ʾ
        Unload Me
        
    End If
End Sub

Private Sub ccmdSelect_Click()
    ccdg.Filter = "All Files (*.*)|*.*|Excel file" & _
            "(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
    ccdg.FileName = ""
    ccdg.ShowOpen
    sub��ʾ������Ϣ
End Sub

Sub sub��ʾ������Ϣ()
    Dim i As Integer
    Dim lstrTmp As String
    
    On Error GoTo errHandler
    
    ctxtDataPath.Text = ccdg.FileName
    With cgrdDetails
        .LoadGrid ccdg.FileName, flexFileExcel, 0
        .FormatString = cgrdDetails.TextMatrix(1, 0)
        For i = 1 To .cols - 1
'            .ColHidden(1) = True    '��������֤�� 2015-11-16 by Ĳ��
            .FormatString = .FormatString & "|" & .TextMatrix(1, i)
        Next i
'        .ColHidden(1) = True    '��������֤�� 2015-11-16 by Ĳ��
        .RowHidden(1) = True
        .AutoSize 0, .cols - 1, 0, 0
    End With

    '������Ϣ��ʽ���жϣ��ò�ͬ��ɫ������޸ĺ�ɸ�Ϊ������ɫ�����ϸ����ܵ��롣
    If lobl�Ѿ����� = False Then ccmdImport.Enabled = sub������Ϣ�Ϸ��Լ��
    
    Exit Sub
errHandler:
    MsgBox ("��ʾ������Ϣ����")
End Sub

Private Sub ccmd��λ��λ_Click()
    On Error GoTo errHandler

'    Dim lobjRec As Object                       '��λ��λ���صĽ����¼��
'    Set lobjRec = pobjҵ�����.func��λ��λ     '������λ��λ���档
'
'    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�(��ʱֻ��ʾ����λ���ơ�)
'    If Not lobjRec Is Nothing Then
'        If lobjRec.RecordCount > 0 Then
'            ctxt��λ����.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
'        End If
'    End If
    
    FrmCompany.Show
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmImportExcel", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub

Private Sub cgrdDetails_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '����ǰ�޸ı��浽excel�У�����ʾ��Ȼ����кϷ��Լ��
    cgrdDetails.ColDataType(indextmp) = flexDTString
    cgrdDetails.SaveGrid ccdg.FileName, flexFileExcel, 0
    cgrdDetails.Editable = flexEDNone
    ccmdImport.Enabled = sub������Ϣ�Ϸ��Լ��
End Sub

Private Sub cgrdDetails_DblClick()
    If lbol�Ѿ����� = True Then Exit Sub        '�Ѿ������ļ��󣬲�����Ԫ��༭
    cgrdDetails.Editable = flexEDKbdMouse
    If cgrdDetails.MouseCol < 0 Or cgrdDetails.MouseRow < 0 Then
        Exit Sub
    ElseIf cgrdDetails.MouseCol >= 0 And cgrdDetails.MouseCol < cgrdDetails.cols Then
        sub¼���޸����� cgrdDetails.MouseRow, cgrdDetails.MouseCol
    End If
End Sub

Sub sub¼���޸�����(ByVal paraRow As Integer, ByVal paraCol As Integer)
    cgrdDetails.Select paraRow, paraCol
    cgrdDetails.EditCell
End Sub

Private Sub Form_Load()
        
    ccmdImport.Enabled = False
    lobl�Ѿ����� = False
    
    Set mobj��� = CreateObject("ְҵ������.clsMedicalExam")
    Set mobj����ģ�� = CreateObject("ְҵ������.clsMedicalExamTemplate")

    sub������������List
    sub�����������List
End Sub

Private Sub subChangeTemplate()
    On Error GoTo errHandler
    
    If mobj���.����.������ <> ccmbTemplate.Text Then
        mobj���.����.������ = ccmbTemplate.Text
        '��������ģ���ȡ���������п��õ���ĸ��
        mobj����ģ��.������ = ccmbTemplate.Text
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmImportExcel", "subChangeTemplate", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'ȫ��Ϊ�������Ա���͡����ֵ����ݣ�ְҵ���������乤��
'����������ͼ�����Ͽ���
Sub sub������������List()
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    Set lobjRec = pobjDict.FetchEx("���������ֵ�")
    ccmb���������.Clear
    'ccmb���������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb���������.AddItem lobjRec("����")
        ccmb���������.ItemData(ccmb���������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb���������.ListIndex = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmImportExcel", "sub������������List", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'ȫ��Ϊ�������𡱣��ֵ����ݣ��ϸ�ǰ���ڸ��ڼ䡢���ʱ��Ӧ�����
'���������������Ͽ���
Sub sub�����������List()
    Dim lobjRec As Object
    On Error GoTo errHandler

    Set lobjRec = pobjDict.FetchEx("��������ֵ�")
    Ccmb��������.Clear
    'Ccmb��������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        Ccmb��������.AddItem lobjRec("����")
        Ccmb��������.ItemData(Ccmb��������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    Ccmb��������.ListIndex = 0
    Ccmb��������.Visible = True
    Call Ccmb��������_Click
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmImportExcel", "sub�����������List", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'�Ϸ��Լ�����ݣ�1����Ҫ��Ϣ�Ƿ�Ϊ��(����Ϊ�գ���ȫ���Ը������Ա��Ϣ)��
'2�������������͸�ʽ�ͷ�Χ�Ƿ���ȷ
'3�������ĳ�б��ǳ��ɫ�������ĳ��Ԫ����������塢��ɫ��
Function sub������Ϣ�Ϸ��Լ��() As Boolean
    Dim i, j As Integer
    Dim changeRowColor As Boolean
    
    On Error GoTo errHandler
    
    sub������Ϣ�Ϸ��Լ�� = True
    For i = 2 To cgrdDetails.rows - 1
        changeRowColor = False
        For j = 0 To cgrdDetails.cols - 1
            '����һ��û�����ݣ���Ϊ���в����롣Ҳ����ʾ��
            If cgrdDetails.TextMatrix(0, j) = "����" And cgrdDetails.TextMatrix(i, j) = "" Then
                '2013-02-26 ������
                '������Ϊɾ���У���Ϊ���غ���ѭ������ʱ��Ч
'                cgrdDetails.RemoveItem (i)
                cgrdDetails.RowHidden(i) = True
                '2013-02-26 ������
                Exit For
            End If
            
            '�жϡ����䡱�����ݺϷ���
            If cgrdDetails.TextMatrix(0, j) = "����" Then
                If IsNumeric(cgrdDetails.TextMatrix(i, j)) = True Then
                    If CLng(cgrdDetails.TextMatrix(i, j)) > 130 Then
                        sub���Ϸ���Ԫ���ǲ�Ⱦɫ changeRowColor, i, j
                    End If
                Else
                    sub���Ϸ���Ԫ���ǲ�Ⱦɫ changeRowColor, i, j
                End If
            End If
            
            '�жϡ����䡱�����ݺϷ���
'            If cgrdDetails.TextMatrix(0, j) = "����" Then
'                'If IsNumeric(cgrdDetails.TextMatrix(i, j)) = False Or CLng(cgrdDetails.TextMatrix(i, j)) > 70 Then sub���Ϸ���Ԫ���ǲ�Ⱦɫ changeRowColor, i, j
'            End If
            
            '�жϡ��������ڡ������ݺϷ���
            If cgrdDetails.TextMatrix(0, j) = "��������" Then
                If IsDate(cgrdDetails.TextMatrix(i, j)) = False Then sub���Ϸ���Ԫ���ǲ�Ⱦɫ changeRowColor, i, j
            End If
            
            '�жϡ���񡱵����ݺϷ��ԣ�ֻ��д���ѻ顱����δ�顱�������족��
'            If cgrdDetails.TextMatrix(0, j) = "���" Then
'                If Not (cgrdDetails.TextMatrix(i, j) = "�ѻ�" Or cgrdDetails.TextMatrix(i, j) = "δ��" Or cgrdDetails.TextMatrix(i, j) = "����") Then sub���Ϸ���Ԫ���ǲ�Ⱦɫ changeRowColor, i, j
'            End If
            
'            '�жϡ����֤�š������ݺϷ���
'            If cgrdDetails.TextMatrix(0, j) = "������ݺ���" Then
'                indextmp = j                '���¡�������ݺ��롱����һ�У�����ʱ�õ�
'
'                '�������ݱ�ǲ�ͬ��ɫ������ɫ�仯
'                Dim lstrSex As String
'                Dim lstrBirth As String
'                sub���ݹ�����ݺ����ȡ���պ��Ա� cgrdDetails.TextMatrix(i, j), lstrBirth, lstrSex
'                If IsDate(lstrBirth) = False Then sub���Ϸ���Ԫ���ǲ�Ⱦɫ changeRowColor, i, j
'            End If

            '�жϡ����֤�š������ݺϷ���
            '��������֤���������֤������֤��   2015-11-16 by Ĳ��
            If cgrdDetails.TextMatrix(0, j) = "������ݺ���" Then
                indextmp = j                '���¡�������ݺ��롱����һ�У�����ʱ�õ�

                '�������ݱ�ǲ�ͬ��ɫ������ɫ�仯
                Dim lstrSex As String
                Dim lstrBirth As String
                sub���ݹ�����ݺ����ȡ���պ��Ա� cgrdDetails.TextMatrix(i, j), lstrBirth, lstrSex
                If cgrdDetails.TextMatrix(i, j) = "" And cgrdDetails.TextMatrix(i, j + 1) <> "" Then
                cgrdDetails.TextMatrix(i, j) = cgrdDetails.TextMatrix(i, j + 1)
                cgrdDetails.TextMatrix(i, j + 1) = ""
'                lstrBirth1 = cgrdDetails.TextMatrix(i, j + 5)
'                dafuncGetData ("insert into ְҵ�����_�����Ա������Ϣ��(��������) value('" & cgrdDetails.TextMatrix(i, j + 5) & "')")
                Else
                If IsDate(lstrBirth) = False Then sub���Ϸ���Ԫ���ǲ�Ⱦɫ changeRowColor, i, j
                End If
            End If

        Next j
        
        If changeRowColor = True Then
            For j = 0 To cgrdDetails.cols - 1
                cgrdDetails.Cell(flexcpBackColor, i, j) = &HC0FFFF      '����ɫǳ��
            Next j
            sub������Ϣ�Ϸ��Լ�� = False
        Else
            'ȡ����Ԫ��Ⱦɫ
            For j = 0 To cgrdDetails.cols - 1
                cgrdDetails.Cell(flexcpForeColor, i, j) = vbBlack       '������ɫ��ɫ
                cgrdDetails.Cell(flexcpFontBold, i, j) = 0              '����һ���ϸ
                cgrdDetails.Cell(flexcpBackColor, i, j) = 0             '��ɫ
            Next j
        End If
    Next i
    
    For i = cgrdDetails.rows - 1 To 2 Step -1
        For j = 0 To cgrdDetails.cols - 1
            '����һ��û�����ݣ���Ϊ���в����롣Ҳ����ʾ��
            If cgrdDetails.TextMatrix(0, j) = "����" Then
                 If cgrdDetails.TextMatrix(i, j) = "" Then
                    '2013-02-26 ������
                    '������Ϊɾ���У���Ϊ���غ���ѭ������ʱ��Ч
                    cgrdDetails.RemoveItem (i)
    '                cgrdDetails.RowHidden(i) = True
                    '2013-02-26 ������
                End If
                Exit For
            End If
        Next j
    Next i
    
    Exit Function
errHandler:
    sub������Ϣ�Ϸ��Լ�� = False
    MsgBox ("���ݸ�ʽ��������ϸ����˫���޸ġ�")
End Function

Sub sub���Ϸ���Ԫ���ǲ�Ⱦɫ(changeRowColor As Boolean, ByVal index_X As Integer, ByVal index_Y As Integer)
    '��������ȾΪ��ɫ������Ӵ֣��������ɫ�仯
    changeRowColor = True
    cgrdDetails.Cell(flexcpForeColor, index_X, index_Y) = vbBlue        '������ɫ
    cgrdDetails.Cell(flexcpFontBold, index_X, index_Y) = 1              '�������
End Sub


Sub sub�����ļ�����(ByVal paraCol As Collection)
        On Error GoTo errHandler

        Dim lobjRec As Object
        Dim lstrError As String
        
        Set mcol�շ���Ŀ = New Collection
        Set mcol�����Ŀ = New Collection
        
        mobj���.ϵͳ��� = paraCol("ϵͳ���")
        Set lobj������ = CreateObject("ְҵ������.clsmedicalexamsheet")
        lobj������.������ = ccmbTemplate.Text
        With mobj���
            If .����.������ <> ccmbTemplate.Text Then
                .����.������ = paraCol("��������") + "-" + paraCol("�������") + "-" + paraCol("Σ������")
            End If
''            If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
''                If .����.�Թܱ����ĸ <> FrmRegister.clblLetter.Caption Then
''                    .����.�Թܱ����ĸ = FrmRegister.clblLetter.Caption
''                End If
''            Else
''                .����.�Թܱ����ĸ = FrmRegister.clblLetter.Caption
''                .�Թܱ�� = FrmRegister.ctxtTubeNo.Text
''            End If
            On Error Resume Next
            .�����Ա.ϵͳ��� = paraCol("ϵͳ���")
            .�����Ա.���� = paraCol("����")
            .�����Ա.�Ա� = paraCol("�Ա�")
            .�����Ա.��λ���� = ctxt��λ����
'            If ctxt��λ����.Text = "" Then .�����Ա.��λ���� = paraCol("��λ����")
            .�����Ա.Σ������ = paraCol("Σ������")
            .�����Ա.����Դ = paraCol("����Դ")
            .�����Ա.ְҵ���� = paraCol("������")
            .�����Ա.�ֹ��� = paraCol("�ֹ���")
            .�����Ա.ְ���ְ�� = paraCol("ְ���ְ��")
            .�����Ա.ְҵΣ������ = paraCol("ְҵΣ������")
            .�����Ա.������� = paraCol("�������")
            .�����Ա.���� = paraCol("����")
            .�����Ա.�ʱ� = paraCol("�ʱ�")
            .�����Ա.סַ = paraCol("סַ")
            .�����Ա.��� = paraCol("���")
            .�����Ա.�绰���� = paraCol("�绰����")
            .�����Ա.���� = paraCol("����")
            .�����Ա.������ = paraCol("������")
            .�����Ա.��ϵ�绰 = paraCol("��ϵ�绰") '������ġ��绰���롱�ظ����������Ǹ����˵���ϵ�绰
            .�����Ա.�������� = paraCol("��������")
            .�����Ա.��ҵ��� = paraCol("��ҵ���")
            .�����Ա.��λ��ַ = paraCol("��λ��ַ")
            .�����Ա.������ = paraCol("������")
            .�����Ա.���� = paraCol("����")
             If paraCol("��������") <> "" Then .�����Ա.�������� = Format(paraCol("��������"), "yyyy/mm/dd")
'            .�����Ա.�������� = DateAdd("yy-mm-dd", Val(paraCol("��������")), Date)
'            If paraCol("��������") = "" Then .�����Ա.�������� = DateAdd("yyyy", -Val(paraCol("����")), Date)
            .�����Ա.������ݺ��� = paraCol("������ݺ���")
            .�����Ա.�������� = paraCol("��������")
            .�����Ա.��ҵ��� = paraCol("��ҵ���")
            .�����Ա.Ƭ�� = paraCol("Ƭ��")
            .�����Ա.�Ļ��̶� = paraCol("�Ļ��̶�")
            .�����Ա.���� = paraCol("����")
            .�����Ա.ְҵ���� = paraCol("ְҵ����")
            .�����Ա.�������� = paraCol("��������")
            .�����Ա.������� = paraCol("�������")
'           If paraCol("��λ������") = "" Then
'                .�����Ա.��λ������ = ""
'            Else
                dasubSetQueryTimeout 600
                Set lobjRec = dafuncGetData("select * from ��λ����_��λ������Ϣ�� where ��λ����='" & .�����Ա.��λ���� & "'")
                If lobjRec.RecordCount > 0 Then mstr��λ������ = lobjRec("������")
                If .�����Ա.��λ������ <> mstr��λ������ Then
                    '����λ������¸�ֵ���������»�ȡ���������ࡢ��ҵ���Ƭ����
                    .�����Ա.��λ������ = mstr��λ������
                End If
'            End If
            
            '���渽����Ϣ
            'For i = 1 To ciptBase.ItemCount
                'If ciptBase.Box1(i - 1).TrueText <> ciptBase.Box1(i - 1).Text And ciptBase.Box1(i - 1).Text <> "" Then
             '   If ciptBase.InfoCollection(i).�ֵ����� <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
             '       .����.Sub�����Ϣֵ ciptBase.InfoCollection(i).����, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
             '   Else
             '       .����.Sub�����Ϣֵ ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
             '   End If
           ' Next i
            
            '����Ϊ��������
            'If ccmb���������.Text = "����" Then
            '    .�������� = P_EXAM_FIRST
            'Else
            '    .�������� = P_EXAM_ANNUAL
            'End If
            .������� = Now 'Format(cdtpDate.Value, "yyyy-mm-dd hh:mm:ss")
            
            '�޸ģ�2004-1-9��������쵥�ţ�
            '.��쵥�� = ctxt��쵥��.Text
            'ֱ��ȡ�����ֵ 2015-7-2 by lanchao
            '.��������� = ccmb���������.Text
            '.�������� = Ccmb��������.Text
            .��������� = paraCol("��������")
            .�������� = paraCol("�������")
            Dim tjbh As String
            tjbh = paraCol("��������") + "-" + paraCol("�������") + "-" + paraCol("Σ������")
            On Error GoTo errHandler
            Set lobj������ = CreateObject("ְҵ������.clsmedicalexamsheet")
            lobj������.������ = Trim(tjbh)
            If mcol�����Ŀ.Count = 0 Then
                Set mcol�����Ŀ = lobj������.������Ŀ��("")
            End If
            Set .col�����Ŀ = mcol�����Ŀ
            Set lobj������ = Nothing
        End With
        
        '���ܣ����������Ŀ
        'ʱ�䣺2012-06-04
        '���ߣ�����
        save�Ż��������Ŀ mcol�����Ŀ, paraCol("ϵͳ���")
        'ʱ�䣺2012-06-04
        
        If mcol�շ���Ŀ.Count > 0 Then
            pobjҵ�����.Sub���Ǽ� mobj���, , , mcol�շ���Ŀ, Val(1)
        Else
            pobjҵ�����.Sub���Ǽ� mobj���, , , , Val(1)
        End If
        
        Set lobjRec = CreateObject("ְҵ������.clsMoney")
        lobjRec.mstrϵͳ��� = paraCol("ϵͳ���")
        lobjRec.mstr�����Ա���� = paraCol("����")
        Set lobjRec.col�����Ŀ = mcol�����Ŀ
        Dim lstr�շ����� As String
    '    lstrError = lobjRec.func�շ�(lstr�շ�����)
        mobj���.�շ����� = lstr�շ�����
        Set lobjRec = Nothing
       ' If lstrError <> "" And lstrError <> "Cancel" Then
      '      MsgBox lstrError, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
      '  End If

        
        'mobj���.Sub�������Ǽ���Ϣ
        
        '2012-06-25 �ڵ�� ��
        '��ʼ����������Ϣ���С��������״̬���ֶ�
        subInit�������״̬ mcol�����Ŀ, Trim(paraCol("ϵͳ���"))
        '2012-06-25 �ڵ�� ��
        MousePointer = 0
    Exit Sub
errHandler:
   MousePointer = 0
   sfsub������ "ְҵ��ʷ¼��", "frmImportExcel", "sub�����ļ�����", Err.Number, Err.Description, False
End Sub

'���ܣ�����ְҵ���Ǽ���ѡ��������Ŀ
'���ߣ�����
'ʱ�䣺2012-06-04
'˵��������Ҫ�鿴���ݿ������Ƿ�����ͬ�������Ŀ��Ȼ���ٽ������ӻ����޸�

Public Sub save�Ż��������Ŀ(ByRef para�����Ŀ As Collection, ByVal paraϵͳ��� As String)
    Dim lstrSql As String
    Dim MedicProjt As String
    Dim rs As Object
    Dim i As Integer
    Dim col�����Ŀ As Collection
    On Error GoTo errHandler
    
    Set rs = dafuncGetData("select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ���������ֵ�') and ���� like '%��'")
    
    For i = 1 To rs.RecordCount
        
        lstrSql = "delete ְҵ�����_�����Ϣ_" & rs("����") & " where ϵͳ���='" & paraϵͳ��� & "'"
        dafuncGetData lstrSql
        rs.MoveNext
    Next i
    
    Set col�����Ŀ = para�����Ŀ
    
    For i = 1 To col�����Ŀ.Count
        MedicProjt = Left(Trim(col�����Ŀ(i)("����")), 2)
        
        lstrSql = "select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ���������ֵ�') and ���= '" & MedicProjt & "'"
        Set rs = dafuncGetData(lstrSql)
        
        lstrSql = "insert into ְҵ�����_�����Ϣ_" & rs("����") & "(ϵͳ���,�����Ŀ) values(" _
            & "'" & paraϵͳ��� & "','" & col�����Ŀ(i)("����") & "')"
        dafuncGetData lstrSql
    Next i
    
    
    Exit Sub
errHandler:
   sfsub������ "ְҵ������", "frmImportExcel", "save�Ż��������Ŀ", Err.Number, Err.Description, False
End Sub

'2012-07-09 �ڵ��
'��ӹ�������ѯ����
Sub sub���¹�������ѯ()
    frmRegisterManage.mstr��ֹ���� = Now
    frmRegisterManage.sub��ѯ����ʾ
End Sub

'2012-06-25 �ڵ��
'��ӳ�ʼ���������״̬������
'�����ж�ÿ�������Ա���������(�����)���ҵ����״̬��
'0������Ҫ����Ŀ��ң�1������Ҫ����Ŀ��ң�2����ÿ����Ѿ������ꣻ
'3����ÿ������������۲��������޸ġ�(���У�2��3״̬�����������ս���)
'״̬��һ������Ϊ13���ַ���(6-25ʱ��13����д����Ŀ��ң��ַ�������Ϊ18)
Sub subInit�������״̬(paraCol As Collection, paraSysNo As String)
    Dim i As Integer
    Dim paraDeptNo As Integer
    Dim paraState, strSQL As String
    
    
    For i = 1 To 19: paraState = paraState & "0": Next
    paraState = paraState & "1"
    
    For i = 1 To paraCol.Count
        paraDeptNo = CInt(Left(paraCol.Item(i).Item(1), 2))
        paraState = Left(paraState, paraDeptNo - 1) & "1" & Right(paraState, Len(paraState) - (paraDeptNo))
    Next
    
    strSQL = "update ְҵ�����_��������Ϣ�� set �������״̬='" & paraState & "' where ϵͳ���='" & paraSysNo & "'"
    dafuncGetData strSQL
End Sub

