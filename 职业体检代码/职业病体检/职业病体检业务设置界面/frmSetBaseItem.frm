VERSION 5.00
Object = "{DA03AE6C-F6DB-11D1-9E6E-0040053F8E31}#3.3#0"; "DYCOMMINPUT.OCX"
Begin VB.Form frmSetBaseItem 
   Caption         =   "������츽�ӻ�����Ŀ"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12300
   Icon            =   "frmSetBaseItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12300
   Begin ��Դͨ��¼��ؼ�.DyInputGrid cgrdInput 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   11456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ErrorColor      =   12648447
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "      (2) ѡ�������е�ĳ�У�����Del��������ɾ��һ��"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   4590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "˵����(1) ��ν������Ŀ���ǳ����������Ա����䡢��λ���ơ�������������Ա���е�           ��������������ԡ�"
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   7260
   End
End
Attribute VB_Name = "frmSetBaseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��

'������:��Х��
'german
'���ܣ�����ɾ������
Private Sub delete_any_Click()
    Dim loop_a As Integer
    Dim lcolInfo As New Collection
    Dim flag As Boolean
    
    flag = False
    
    For loop_a = 0 To cgrdInput.Rows Step 1
            If (cgrdInput.Value(loop_a, "�Ƿ�ɾ��") = "1") Then
            'string_debug = cgrdInput.Value(loop_a, "ɾ����־") & " " & string_debug
                'MsgBox "okay", , "okay"
                'lobjItem.���� = cgrdInput.Value(loop_a, "����")
                'lobjItem.subDelete_any
                pobjҵ�����.Sub������츽����Ŀ 3, lcolInfo, cgrdInput.Value(loop_a, "������Ŀ")
                'MsgBox CStr(loop_a), , "[DEBUG��Ϣ]"
                '����1 �������� ����2 ��Ŀ����
                flag = True
            ElseIf (cgrdInput.Value(loop_a, "�Ƿ�ɾ��") <> "") Then
                MsgBox "�Ƿ�ɾ�� �ֶ���Ҫô���� ҪôΪ1��������д��ȷ�������Ա����ʶ��", 16, "��Ϣ"
                Exit Sub
            End If
        Next
        
    If (flag = False) Then
        MsgBox "��û��Ҫ����ɾ�������ݣ���ѡ��", 16, "��Ϣ"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub
Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lcolInfo As Collection
    
    On Error GoTo errHandler
    '��ʼ��¼�����񣺸�����Ŀ,¼�����,��������,���ݳ���,ö��ֵ
    cgrdInput.InputTemplate.subRemoveAllItem
    With cgrdInput.InputTemplate
        .subAddItem "������Ŀ", 0, True, True, 30
        .subAddItem "¼�����", 0, False, True, 30
        .subAddItem "��������", 0, True, True, 20, 0, "1 ������,2 ������,3 �ı���" ' dyInputSingleselecttext
        .subAddItem "���ݳ���", 3, True, True, 4, 0, , 300, 1
        .subAddItem "ö��ֵ", 0, False, True, 50
    End With
    cgrdInput.subDraw
    
    '���á�pobjҵ�����.������츽����Ŀ����
    '��Ϣ�����ݿ������Ѿ��޸�
    Set lobjRec = pobjҵ�����.������츽����Ŀ  '������ҵ�����clsManageMedicalExam
    
    '�ѻ�ȡ����Ŀ����¼�������С�
    gfsubLoadDyGridFromRec cgrdInput, lobjRec
    
    cgrdInput.subExpand '�Զ���������Ԫ�������ݵĴ�С������Ӧ�������
    lobjRec.Close
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetBaseItem", "Form_Load", 6666, lstrError, False
End Sub

'�޸���: ��Х��
'german
Private Sub cgrdInput_AddNew(paraValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lcolInfo As Collection
    Dim llng�������� As Long
    Dim loop_a As Integer
    
    On Error GoTo errHandler
    '��ȡ��������Ŀ��Ϣ��
    Set lcolInfo = New Collection
    
    '///////////////////////////////////////////////////////////
    '�ж��û�����������Ƿ��Ǻ���
    'german
    For loop_a = 1 To Len(paraValue("������Ŀ")("ֵ")) Step 1
        If (Asc(Mid(paraValue("������Ŀ")("ֵ"), loop_a, 1)) > 0) Then
            MsgBox "[������Ŀ]:����������ݱ����Ǻ��֣����ݼ�¼�����Ѿ�������", 16, "��Ϣ"
            Exit Sub
        End If
    Next
    
    For loop_a = 1 To Len(paraValue("ö��ֵ")("ֵ")) Step 1
        If (Asc(Mid(paraValue("ö��ֵ")("ֵ"), loop_a, 1)) >= 48 And Asc(Mid(paraValue("ö��ֵ")("ֵ"), loop_a, 1)) <= 48) Then
            MsgBox "[������Ŀ]:����������ݲ��������֣����ݼ�¼�����Ѿ�������", 16, "��Ϣ"
            Exit Sub
        End If
    Next
    
    '///////////////////////////////////////////////////////////
    
    lcolInfo.Add paraValue("������Ŀ")("ֵ"), "��Ŀ����"
    lcolInfo.Add IIf(IsNull(paraValue("¼�����")("ֵ")), "", paraValue("¼�����")("ֵ")), "¼�����"
    
    llng�������� = Left(paraValue("��������")("ֵ"), 1)
    
    lcolInfo.Add llng��������, "��������"
    lcolInfo.Add paraValue("���ݳ���")("ֵ"), "���ݳ���"
    lcolInfo.Add IIf(IsNull(paraValue("ö��ֵ")("ֵ")), "", paraValue("ö��ֵ")("ֵ")), "ö��ֵ"
    
    '���������ݿ��С�
    pobjҵ�����.Sub������츽����Ŀ 1, lcolInfo  '������ҵ�����clsManageMedicalExam
    
    Exit Sub
    
errHandler:
    ErrorInfo = func������(Err.Number, Err.Description)
    Cancel = True
End Sub


Private Sub cgrdInput_RowChange(ByVal paraRow As Long, paraNewValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lcolInfo As Collection
    Dim llng�������� As Long
    
    On Error GoTo errHandler
    '��ȡ��������Ŀ��Ϣ��
    Set lcolInfo = New Collection
    lcolInfo.Add paraNewValue("������Ŀ")("ֵ"), "��Ŀ����"
    lcolInfo.Add IIf(IsNull(paraNewValue("¼�����")("ֵ")), "", paraNewValue("¼�����")("ֵ")), "¼�����"
    llng�������� = Left(paraNewValue("��������")("ֵ"), 1)
    lcolInfo.Add llng��������, "��������"
    lcolInfo.Add paraNewValue("���ݳ���")("ֵ"), "���ݳ���"
    lcolInfo.Add IIf(IsNull(paraNewValue("ö��ֵ")("ֵ")), "", paraNewValue("ö��ֵ")("ֵ")), "ö��ֵ"
    
    '���������ݿ��С�
    pobjҵ�����.Sub������츽����Ŀ 2, lcolInfo, cgrdInput.Value(paraRow, "������Ŀ")

    Exit Sub
errHandler:
    ErrorInfo = func������(Err.Number, Err.Description)
    Cancel = True
    cgrdInput.ItemSetfocus "������Ŀ"
    Exit Sub
    Resume
End Sub

Private Sub cgrdInput_Delete(ByVal paraRow As Long, Cancel As Boolean, ErrorInfo As String)
    Dim lcolInfo As New Collection
    
    On Error GoTo errHandler
    
    '�����ݿ���ɾ��������Ŀ��
    'MsgBox CStr(paraRow), , "[DEBUG��Ϣ]"
    pobjҵ�����.Sub������츽����Ŀ 3, lcolInfo, cgrdInput.Value(paraRow, "������Ŀ")

    Exit Sub
errHandler:
    ErrorInfo = func������(Err.Number, Err.Description)
    Cancel = True
End Sub

Private Sub ccmdExit_Click()
    Unload Me
End Sub
