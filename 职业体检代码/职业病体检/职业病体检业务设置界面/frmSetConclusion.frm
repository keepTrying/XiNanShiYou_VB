VERSION 5.00
Object = "{DA03AE6C-F6DB-11D1-9E6E-0040053F8E31}#3.3#0"; "DYCOMMINPUT.OCX"
Begin VB.Form frmSetConclusion 
   Caption         =   "���ս���ģ������"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10830
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ccmd���� 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ǰ��Ϣ"
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.Label LblDate 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label LblNo 
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label LblName 
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "��ǰʱ�䣺"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "ҽʦ��ţ�"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "ҽʦ������"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����б�"
      Height          =   5895
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin ��Դͨ��¼��ؼ�.DyInputGrid cgrdInput 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9128
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵����ѡ���ұ������е�ĳ�У�����Del��������ɾ��һ��,ֱ���޸ĵ�Ԫ���еĽ����ܹ�ֱ���޸Ľ���"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   5520
         Width           =   8100
      End
   End
End
Attribute VB_Name = "frmSetConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ccmd����_Click()
    Unload Me
End Sub

Private Sub cgrdInput_BeforeAddNew(paraCancel As Boolean)
    On Error GoTo errHandler
    '���� 2012.12.28  ����
    '˵������addnew֮���ѯһ��������䵽���ʹ��������Ӧ���ݡ�
    Set lobj����ģ�� = CreateObject("ְҵ������.clsConclusionSet")
    Set lobjRec = lobj����ģ��.func��ȡ���ս���ģ��
    
    gfsubLoadDyGridFromRec cgrdInput, lobjRec '���������б��Լ����ݿ����
    cgrdInput.subExpand
    lobjRec.Close
    '���� 2012.12.28  ����
    Exit Sub
errHandler:
    paraCancel = True
End Sub

'���ߣ�����
'���ܣ�������Ŀ��ʱ�򣬼�������У�飬�����������޹ص���������
'ʱ�䣺2012-05-29
Private Sub cgrdInput_AddNew(paraValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobj����ģ�� As Object
    Dim lobjRec As Object
    Dim flag As Boolean
    Dim flag1 As Boolean
    
    On Error GoTo errHandler

    'У�����ģ��������Ƿ�Ϊ��
    If (IsNull(paraValue("����ģ��")("ֵ"))) Or (paraValue("����ģ��")("ֵ") = "") Then
        MsgBox "����д�Ľ���ģ�岻��Ϊ��", 16, "��Ϣ"
        cgrdInput.subRefresh
        Exit Sub
    End If
    
    Set lobj����ģ�� = CreateObject("ְҵ������.clsConclusionSet")
    
    flag = lobj����ģ��.func�Ƿ����(paraValue("����ģ��")("ֵ"), paraValue("���۱�׼")("ֵ"))
    If flag = False Then
        flag1 = lobj����ģ��.func������ս���ģ��("16", "���ս���¼��", paraValue("����ģ��")("ֵ"), um�û���, paraValue("���۱�׼")("ֵ"))
    
        If flag1 = False Then
            MsgBox "��д��������������Ϣ�������������ַ���", 16, "��Ϣ"
            cgrdInput.subRefresh
            Exit Sub
        End If
    End If
    Set lobj����ģ�� = Nothing
    Exit Sub
errHandler:
    Cancel = True
    ErrorInfo = func������(Err.Number, Err.Description)
End Sub

Private Sub cgrdInput_BeforeRowChange(ByVal paraRow As Long, paraCancel As Boolean)
    On Error GoTo errHandler
    
    cgrdInput.ItemEnabled("��������") = False
    cgrdInput.ItemEnabled("����ҽʦ") = False
    
    Exit Sub
errHandler:
    paraCancel = True
    
End Sub

Private Sub cgrdInput_ItemChange(ByVal paraItem As String)
    
    On Error GoTo errHandler
'    If cgrdInput.ItemValue(paraItem) <> "" Then
'        Select Case paraItem
'        Case "����ģ��"
''            '���������Ŀ����
''            Set lobjItem = CreateObject("ְҵ������.clsTestItem")
''
''            '��¼�����"����"���������Ƿ�Ψһ�����Ѵ��ڣ���ʾ����
''            lobjItem.���� = cgrdInput.ItemValue(paraItem)
''            If lobjItem.�Ƿ���� Then
''                Err.Raise 6666, , "���롰" & cgrdInput.ItemValue(paraItem) & "���Ѵ��ڡ����벻�����ظ���������������롣", sf����
''            End If
'        End Select
'    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    If Err.Number <> 60844 Then
        lstrError = func������(Err.Number, Err.Description)
        sffuncMsg lstrError, sf����
        'cgrdInput.ItemValue(paraItem) = ""
        cgrdInput.ItemSetfocus paraItem
    End If
    Exit Sub
    Resume
End Sub

'���ߣ����� 2012.11.30
'˵������ascii����Ʋ��������������   ����
'bug�ţ�0000044
Private Sub cgrdInput_ItemKeyPress(ByVal paraItem As String, KeyAscii As Integer)
    Dim a As Integer
    Dim b As Integer
    a = 1
    b = 1
    If KeyAscii >= 48 And KeyAscii <= 57 Or ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123)) Then
    Else
        a = 0
    End If
        
    If KeyAscii >= -20319 And KeyAscii <= -3652 Or KeyAscii = 8 Then
    Else
        b = 0
    End If
    
    If a = 0 And b = 0 Then
    KeyAscii = 0
    End If
End Sub
'                            ����

Private Sub cgrdInput_RowChange(ByVal paraRow As Long, paraNewValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim lstrEnum As String
    
    On Error GoTo errHandler
    '���������Ŀ����
'    Set lobjItem = CreateObject("ְҵ������.clsTestItem")
'    'MsgBox "german", , "german"
'    '���ݵ�ǰ¼������Ϣ����lobjItem�����ԡ�
'    With lobjItem
'        .���� = cgrdInput.Value(paraRow, "����")
'        .���� = paraNewValue("����")("ֵ")
'
'        lstrEnum = IIf(IsNull(paraNewValue("ö����Դ")("ֵ")), "", paraNewValue("ö����Դ")("ֵ"))
'
'        '�����Ķ��Ż���Ӣ�Ķ��š�
'        If lstrEnum <> "" Then
'            lstrEnum = gffuncStrReplace(lstrEnum, "��", ",")
'        End If
'
'        .ö����Դ = lstrEnum
'        .ȱʡֵ = IIf(IsNull(paraNewValue("ȱʡֵ")("ֵ")), "", paraNewValue("ȱʡֵ")("ֵ"))
'        .���� = paraNewValue("����")("ֵ")
'        .������ = Right(ctrwType.SelectedItem.Key, Len(ctrwType.SelectedItem.Key) - 1)
'
'        '���ȱʡֵ�Ƿ���ö����Դ�С�
'        If .ȱʡֵ <> "" And lstrEnum <> "" Then
'            If Right(lstrEnum, 1) <> "," Then lstrEnum = lstrEnum & ","
'            If InStr(1, lstrEnum, .ȱʡֵ & ",") = 0 Then
'                Err.Raise 6666, , "ȱʡֵ������ö����Դ�С�"
'            End If
'        End If
'        .�ȽϷ�ʽ = IIf(IsNull(paraNewValue("�ȽϷ�ʽ")("ֵ")), "", paraNewValue("�ȽϷ�ʽ")("ֵ"))
'        .��׼ֵ = IIf(IsNull(paraNewValue("��׼ֵ")("ֵ")), "", paraNewValue("��׼ֵ")("ֵ"))
'        .��λ = IIf(IsNull(paraNewValue("��λ")("ֵ")), "", paraNewValue("��λ")("ֵ"))
'
'    End With
'
'    '���浱ǰ����Ϣ�����ݿ⡣
'    lobjItem.subSave
'
'    Set lobjItem = Nothing
    Exit Sub
errHandler:
    ErrorInfo = func������(Err.Number, Err.Description)
    Cancel = True
End Sub

'���ߣ�����
'���ܣ���ӳ���ɾ������
'ʱ�䣺2012-05-29
Private Sub cgrdInput_Delete(ByVal paraRow As Long, Cancel As Boolean, ErrorInfo As String)
    Dim lobj����ģ�� As Object
    Dim flag As Boolean
    
    On Error GoTo errHandler
    
    '���������Ŀ����
    Set lobj����ģ�� = CreateObject("ְҵ������.clsConclusionSet")
    flag = lobj����ģ��.funcɾ�����ս���ģ��(cgrdInput.Value(paraRow, "����ģ��"), cgrdInput.Value(paraRow, "���۱�׼"))
    
    If flag = False Then
         MsgBox "ɾ��ʧ�ܣ����˳��������½��룡", 16, "��Ϣ"
            cgrdInput.subRefresh
            Exit Sub
    End If
    
    'cgrdInput.Value(paraRow, "����")
    Exit Sub
errHandler:
    ErrorInfo = func������(Err.Number, Err.Description)
    Cancel = True
End Sub

Private Sub Form_Load()
    Dim lobj����ģ�� As Object
    Dim lobjRec As Object
    
    On Error GoTo errHandler

    LblName.Caption = um�û���
    LblNo.Caption = um�û����
    LblDate.Caption = Date
    
    
    '��ʼ��¼������:���룬���ƣ�ȱʡֵ��ö����Դ(��ѡֵ)�����ԡ�
    'subAddItem,��3������true������ɫΪǳ��
    cgrdInput.Enabled = True
    With cgrdInput.InputTemplate
        .subAddItem "����ģ��", 0, True, True, 1200
        .subAddItem "��������", 0, False, False, 100, , , , , LblDate.Caption
        .subAddItem "���۱�׼", 4, False, True, 50, , "�ϸ�,���ϸ�", , , "�ϸ�"
    End With
    cgrdInput.subDraw
    
    Set lobj����ģ�� = CreateObject("ְҵ������.clsConclusionSet")
    Set lobjRec = lobj����ģ��.func��ȡ���ս���ģ��
    
    gfsubLoadDyGridFromRec cgrdInput, lobjRec '���������б��Լ����ݿ����
    cgrdInput.subExpand
    lobjRec.Close
Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetTestItem", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub
