VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{DA03AE6C-F6DB-11D1-9E6E-0040053F8E31}#3.3#0"; "DYCOMMINPUT.OCX"
Begin VB.Form frmSetTestItem 
   Caption         =   "�����Ŀ����"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13200
   ClipControls    =   0   'False
   Icon            =   "frmSetTestItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   13200
   Begin ��Դͨ��¼��ؼ�.DyInputGrid cgrdInput 
      Height          =   6975
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   12303
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   2355
   End
   Begin MSComctlLib.TreeView ctrwType 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   12303
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˵����ѡ���ұ������е�ĳ�У�����Del��������ɾ��һ��,ֱ���޸ĵ�Ԫ���еĵ����ܹ�ֱ���޸ĵ���"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   8100
   End
End
Attribute VB_Name = "frmSetTestItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ܣ����������Ŀ  ����
'ʱ�䣺2012-04
'���ߣ���Х��

Option Explicit
'�޸��ˣ�����2012.11.29
'bug�ţ�0000036
'˵������ӱ���   ����
Public mstrNode As String
'�޸�             ����

Private mobj�����Ŀ�� As Object  'ClsTestItemSet,�����ȡָ��������������Ŀ��

'german
'���ܣ��л��Ƿ�����ɾ����־
'���ߣ���Х��
Private Sub any_delete_Click()
    If (flag_delete_any = False) Then
        flag_delete_any = True
    Else
        flag_delete_any = False
    End If
End Sub

'german
'���ܣ�����
'���ߣ���Х��
Private Sub cgrdInput_Click()
'    Dim loop_a As Integer
'    Dim string_debug As String
'    string_debug = ""
'    'MsgBox cgrdInput.Rows, , "��Ϣ"
'    'MsgBox cgrdInput.Cols, , "��Ϣ"
'    For loop_a = 0 To cgrdInput.Rows Step 1
'        If (cgrdInput.Value(loop_a, "ɾ����־") = "1") Then
'        'string_debug = cgrdInput.Value(loop_a, "ɾ����־") & " " & string_debug
'            'MsgBox "okay", , "okay"
'        End If
'    Next
End Sub

'���ߣ����� 2012.11.30
'˵������ascii����Ʋ��������������   ����
'bug�ţ�0000044
Private Sub cgrdInput_ItemKeyPress(ByVal paraItem As String, KeyAscii As Integer)
    Dim a As Integer
    Dim b As Integer
    a = 1
    b = 1
    If KeyAscii >= 48 And KeyAscii <= 57 Or ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123)) Or KeyAscii = 44 Then
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

Private Sub Form_Load()
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    
    'ͨ���ֵ�����ȡ"������"������id����ź����ƣ���ʾ��ctrvType�У�key=id����
    Set lobjRec = pobjDict.Fetch("ְҵ���������ֵ�")
    
    ctrwType.Nodes.Add , , "R", "������"
    Do While Not lobjRec.EOF
        Dim str As String
        str = lobjRec!����
        If Right(str, 1) = "��" Then
            ctrwType.Nodes.Add "R", tvwChild, "I" & lobjRec!InnerID, lobjRec!��� & " " & lobjRec!����
        End If
        lobjRec.movenext
    Loop
    
    '��ʼ��¼������:���룬���ƣ�ȱʡֵ��ö����Դ(��ѡֵ)�����ԡ�
    'subAddItem,��3������true������ɫΪǳ��
    With cgrdInput.InputTemplate
        .subAddItem "����", 0, True, True, 5
        .subAddItem "����", 0, True, True, 30
        .subAddItem "ȱʡֵ", 0, False, True, 300
        .subAddItem "ö����Դ", 0, False, True, 500
        .subAddItem "����", 4, True, True, 10, , "����,����" 'dyInputSingleselecttext
        .subAddItem "�ȽϷ�ʽ", 4, False, True, 20, , "=,��,��,��,����,������,��Χ"
        .subAddItem "��׼ֵ", 0, False, True, 300
        .subAddItem "��λ", 0, False, True, 50
        .subAddItem "����", 0, True, True, 10 'german
    End With
    cgrdInput.subDraw
    cgrdInput.Enabled = False
    
    '��������mobj�����Ŀ����
    Set mobj�����Ŀ�� = CreateObject("ְҵ������.clsTestItemSet")
    
    On Error Resume Next
    If ctrwType.Nodes.Count > 0 Then
        ctrwType.Nodes(1).Expanded = True
    End If
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetTestItem", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    insert_danjia '�޸ĵ��� 'german
    Set mobj�����Ŀ�� = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub
Private Sub cgrdInput_BeforeAddNew(paraCancel As Boolean)
    On Error GoTo errHandler
    cgrdInput.ItemEnabled("����") = True
    sub�������
    Exit Sub
errHandler:
    paraCancel = True
End Sub

'�޸��ߣ���Х��
'�������ӣ�������Ŀ��ʱ�򣬼�������У�飬�����������޹ص���������
'3.1
'german
Private Sub cgrdInput_AddNew(paraValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim lstrEnum As String
    Dim tmp_string_1 As String 'german
    Dim loop_a As Integer 'german
    
    On Error GoTo errHandler
    '���������Ŀ����
    Set lobjItem = CreateObject("ְҵ������.clsTestItem")
    
    tmp_string_1 = paraValue("����")("ֵ") 'german
    'MsgBox tmp_string_1, , "��Ϣ"
    '����У��
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'german
'    For loop_a = 1 To Len(tmp_string_1) Step 1
'        If (Asc(Mid(tmp_string_1, loop_a, 1)) < 48 Or Asc(Mid(tmp_string_1, loop_a, 1)) > 57) Then
'            MsgBox "��ı���ֵ����Ϊ���֣����ݴ��������Ѿ����ܾ�", 16, "��Ϣ"
'            Exit Sub
'        End If
'    Next
    'ȡ��forѭ����ֱ���ж�
    '�޸ģ�����
    'bug�ţ�0000044
    '2012.11.29    ����
    If IsNumeric(tmp_string_1) = False Then
        MsgBox "��ı���ֵ����Ϊ���֣����ݴ��������Ѿ����ܾ�", 16, "��Ϣ"
        Exit Sub
    End If
    '2012.11.29    ����
    If (paraValue("����")("ֵ") = "") Then
        MsgBox "������Ʋ���Ϊ�գ����ݴ��������Ѿ����ܾ�", 16, "��Ϣ"
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'german
        'У�鵥�۵������Ƿ�ȫ��Ϊ���֣�����С��0
        If (IsNull(paraValue("����")("ֵ"))) Then
            MsgBox "����д�ĵ��۲���Ϊ��", 16, "��Ϣ"
            cgrdInput.subRefresh
            Exit Sub
        End If
        
        tmp_string_1 = paraValue("����")("ֵ") 'german
        
        For loop_a = 1 To Len(tmp_string_1) Step 1
            If (Asc(Mid(tmp_string_1, loop_a, 1)) < 48 Or Asc(Mid(tmp_string_1, loop_a, 1)) > 57) Then
                MsgBox "����д�ĵ��۱���Ϊ���֣����ݴ��������Ѿ����ܾ�", 16, "��Ϣ"
                Exit Sub
            End If
        Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '���ݵ�ǰ¼������Ϣ����lobjItem�����ԡ�
    With lobjItem
        .���� = paraValue("����")("ֵ")
        If .�Ƿ���� Then
            Err.Raise 6666, , "�����õ������Ŀ�Ѵ��ڣ������Ŀ���벻�����ظ���"
        End If
        
        .���� = paraValue("����")("ֵ")
        
        .ȱʡֵ = IIf(IsNull(paraValue("ȱʡֵ")("ֵ")), "", paraValue("ȱʡֵ")("ֵ"))
        lstrEnum = IIf(IsNull(paraValue("ö����Դ")("ֵ")), "", paraValue("ö����Դ")("ֵ"))
        
        '�����Ķ��Ż���Ӣ�Ķ��š�
        lstrEnum = gffuncStrReplace(lstrEnum, "��", ",")
        
        .ö����Դ = lstrEnum
        .���� = paraValue("����")("ֵ")
        .������ = Right(ctrwType.SelectedItem.Key, Len(ctrwType.SelectedItem.Key) - 1)
        
        '���ȱʡֵ�Ƿ���ö����Դ�С�
        If .ȱʡֵ <> "" And lstrEnum <> "" Then
            If Right(lstrEnum, 1) <> "," Then lstrEnum = lstrEnum & ","
            If InStr(1, lstrEnum, .ȱʡֵ & ",") = 0 Then
                Err.Raise 6666, , "ȱʡֵ������ö����Դ�С�"
            End If
        End If
        
        .�ȽϷ�ʽ = IIf(IsNull(paraValue("�ȽϷ�ʽ")("ֵ")), "", paraValue("�ȽϷ�ʽ")("ֵ"))
        .��׼ֵ = IIf(IsNull(paraValue("��׼ֵ")("ֵ")), "", paraValue("��׼ֵ")("ֵ"))
        .��λ = IIf(IsNull(paraValue("��λ")("ֵ")), "", paraValue("��λ")("ֵ"))
        
        .���� = paraValue("����")("ֵ")
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End With
    
    '���浱ǰ����Ϣ�����ݿ⡣
    lobjItem.subSave
    
    Exit Sub
errHandler:
    Cancel = True
    ErrorInfo = func������(Err.Number, Err.Description)
End Sub

Private Sub cgrdInput_BeforeRowChange(ByVal paraRow As Long, paraCancel As Boolean)
    On Error GoTo errHandler
    cgrdInput.ItemEnabled("����") = False
    
    Exit Sub
errHandler:
    paraCancel = True
    
End Sub

Private Sub cgrdInput_ItemChange(ByVal paraItem As String)
    Dim lobjItem As Object
    
    On Error GoTo errHandler
    If cgrdInput.ItemValue(paraItem) <> "" Then
        Select Case paraItem
        Case "����"
            '���������Ŀ����
            Set lobjItem = CreateObject("ְҵ������.clsTestItem")
            
            '��¼�����"����"���������Ƿ�Ψһ�����Ѵ��ڣ���ʾ����
            lobjItem.���� = cgrdInput.ItemValue(paraItem)
            If lobjItem.�Ƿ���� Then
                Err.Raise 6666, , "���롰" & cgrdInput.ItemValue(paraItem) & "���Ѵ��ڡ����벻�����ظ���������������롣", sf����
            End If
        End Select
    End If
    
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

Private Sub cgrdInput_RowChange(ByVal paraRow As Long, paraNewValue As Collection, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim lstrEnum As String
    
    On Error GoTo errHandler
    '���������Ŀ����
    Set lobjItem = CreateObject("ְҵ������.clsTestItem")
    'MsgBox "german", , "german"
    '���ݵ�ǰ¼������Ϣ����lobjItem�����ԡ�
    With lobjItem
        .���� = cgrdInput.Value(paraRow, "����")
        .���� = paraNewValue("����")("ֵ")
        
        lstrEnum = IIf(IsNull(paraNewValue("ö����Դ")("ֵ")), "", paraNewValue("ö����Դ")("ֵ"))
        
        '�����Ķ��Ż���Ӣ�Ķ��š�
        If lstrEnum <> "" Then
            lstrEnum = gffuncStrReplace(lstrEnum, "��", ",")
        End If
        
        .ö����Դ = lstrEnum
        .ȱʡֵ = IIf(IsNull(paraNewValue("ȱʡֵ")("ֵ")), "", paraNewValue("ȱʡֵ")("ֵ"))
        .���� = paraNewValue("����")("ֵ")
        .������ = Right(ctrwType.SelectedItem.Key, Len(ctrwType.SelectedItem.Key) - 1)
        
        '���ȱʡֵ�Ƿ���ö����Դ�С�
        If .ȱʡֵ <> "" And lstrEnum <> "" Then
            If Right(lstrEnum, 1) <> "," Then lstrEnum = lstrEnum & ","
            If InStr(1, lstrEnum, .ȱʡֵ & ",") = 0 Then
                Err.Raise 6666, , "ȱʡֵ������ö����Դ�С�"
            End If
        End If
        .�ȽϷ�ʽ = IIf(IsNull(paraNewValue("�ȽϷ�ʽ")("ֵ")), "", paraNewValue("�ȽϷ�ʽ")("ֵ"))
        .��׼ֵ = IIf(IsNull(paraNewValue("��׼ֵ")("ֵ")), "", paraNewValue("��׼ֵ")("ֵ"))
        .��λ = IIf(IsNull(paraNewValue("��λ")("ֵ")), "", paraNewValue("��λ")("ֵ"))
        
    End With
    
    '���浱ǰ����Ϣ�����ݿ⡣
    lobjItem.subSave

    Set lobjItem = Nothing
    Exit Sub
errHandler:
    ErrorInfo = func������(Err.Number, Err.Description)
    Cancel = True
End Sub

'�޸��ߣ���Х��
'german
'�޸����� ��ӳ���ɾ������
Private Sub cgrdInput_Delete(ByVal paraRow As Long, Cancel As Boolean, ErrorInfo As String)
    Dim lobjItem As Object
    Dim loop_a As Integer
    
    On Error GoTo errHandler
    '���������Ŀ����
    Set lobjItem = CreateObject("ְҵ������.clsTestItem")
    
    '���ݵ�ǰ¼������Ϣ����lobjItem�����ԡ�
    lobjItem.���� = cgrdInput.Value(paraRow, "����") '�����û�ѡ���еı���ֵ����ɾ���������ݵ�ʱ�����ڶ�λ
    
    'MsgBox CStr(paraRow), , "german" 'german
    
    lobjItem.subDelete '�����ݿ��н�������¼ɾ��
    'ɾ�����еĸ���Ŀ��
'    If (flag_delete_any = False) Then 'german �Ƿ�����ɾ�����ݣ����������ô�Ͱ��յ�������ɾ��������
'        lobjItem.subDelete (flag_database_delete) 'german ����1��Ϊ�����ڰ���DEL����ͬʱ�������ݿ��е���������ɾ�� ����֮
'    Else '����ɾ������,������ѭ��ɨ���û�ָ��Ҫɾ�������ݣ�Ȼ������ɾ��
'        For loop_a = 0 To cgrdInput.Rows Step 1
'            If (cgrdInput.Value(loop_a, "ɾ����־") = "1") Then
'            'string_debug = cgrdInput.Value(loop_a, "ɾ����־") & " " & string_debug
'                'MsgBox "okay", , "okay"
'                lobjItem.���� = cgrdInput.Value(loop_a, "����")
'                lobjItem.subDelete_any
'            End If
'        Next
        
    'End If
        
    Set lobjItem = Nothing
    Exit Sub
errHandler:
    ErrorInfo = func������(Err.Number, Err.Description)
    Cancel = True
End Sub

Private Sub ctrwType_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lobjRec As Object          '��ǰѡ�д�������������Ŀ��
    Dim lcolInfo As Collection     '���뵽¼��������һ�е����ݡ�
    
    On Error GoTo errHandler
    insert_danjia '�޸ĵ��� 'german
    cgrdInput.subClear
    
    If Node.Parent Is Nothing Then
        'ѡ�д����ͽڵ㣬������¼�롣
        cgrdInput.Enabled = False
    Else
        cgrdInput.Enabled = True
        
        '�޸��ˣ�����2012.11.29
        'bug�ţ�0000036
        '�޸�˵����mstrNode��ȡ��������ݣ�ͨ����sub������ݡ���ȡ����   ����
'        '����mobj�����Ŀ��.������=��ǰ�ڵ��key��
'        mobj�����Ŀ��.������ = Right(Node.Key, Len(Node.Key) - 1) 'ClsTestItemSet,�����ȡָ��������������Ŀ��
'
'        'see
'        '��ȡָ��������������Ŀ�����룬���ƣ�ȱʡֵ��ö����Դ�����ԣ������ࡣ
'        Set lobjRec = mobj�����Ŀ��.�����Ŀ '����һ�����ݿ����
'
'        '��lobjRec���������м�¼��ʾ��cgrdInput�С�
'        gfsubLoadDyGridFromRec cgrdInput, lobjRec '���������б��Լ����ݿ����
'        cgrdInput.subExpand
'        lobjRec.Close
        mstrNode = Node.Key
        sub�������
        '2012.11.29     ����
    End If
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmSetTestItem", "ctrwType_NodeClick", Err.Number, Err.Description, False
End Sub

'�޸��ˣ�����2012.11.29
'bug�ţ�0000036
'�޸�˵����ˢ�º���
Private Sub sub�������()
    Dim lobjRec As Object
    '����mobj�����Ŀ��.������=��ǰ�ڵ��key��
    mobj�����Ŀ��.������ = Right(mstrNode, Len(mstrNode) - 1) 'ClsTestItemSet,�����ȡָ��������������Ŀ��
    
    'see
    '��ȡָ��������������Ŀ�����룬���ƣ�ȱʡֵ��ö����Դ�����ԣ������ࡣ
    Set lobjRec = mobj�����Ŀ��.�����Ŀ '����һ�����ݿ����
    
    '��lobjRec���������м�¼��ʾ��cgrdInput�С�
    gfsubLoadDyGridFromRec cgrdInput, lobjRec '���������б��Լ����ݿ����
    cgrdInput.subExpand
    lobjRec.Close
End Sub

Private Sub cgrdInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
    
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmSetTestItem", "cgrdInput_KeyDown", Err.Number, Err.Description, False
End Sub
Private Sub ccmdExit_Click()
    ccmdExit.Caption = "���ڱ�������...."
    insert_danjia 'ִ�е����޸Ĺ���
    ccmdExit.Caption = "����(&X)"
    Unload Me
End Sub

'���ߣ���Х��
'���ܣ��л��Ƿ�ǿ��ɾ�����ݹ��ܱ�־
'german
'Private Sub is_delete_Click()
'    If flag_database_delete = False Then
'        flag_database_delete = True
'        'MsgBox "�任Ϊ��", , "��Ϣ"
'    Else
'        flag_database_delete = False
'        'MsgBox "�任Ϊ��", , "��Ϣ"
'    End If
'End Sub

'���ߣ���Х��
'���ܣ����������޸Ĺ���
'����:3.13
'german
Private Sub insert_danjia()
    Dim loop_a As Integer
    Dim lobjItem
    
    On Error Resume Next
    Set lobjItem = CreateObject("ְҵ������.clsTestItem")
    For loop_a = 1 To cgrdInput.Rows - 1 Step 1
        lobjItem.���� = CStr(cgrdInput.Value(loop_a, "����"))
        lobjItem.���� = cgrdInput.Value(loop_a, "����")
        lobjItem.SubSaveUnitprice
    Next
End Sub

