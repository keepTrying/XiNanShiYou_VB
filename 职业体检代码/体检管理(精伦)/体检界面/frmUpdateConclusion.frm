VERSION 5.00
Begin VB.Form frmUpdateConclusion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����������"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "�޸���ϴ������"
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   8055
      Begin VB.CommandButton ccmdClear 
         Caption         =   "���(&R)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1200
         Width           =   1020
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox cchkTemplate 
         Caption         =   "���鸴��"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox ctxtDiagnosis 
         Height          =   1935
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "<<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   1020
      End
      Begin VB.ListBox clstAllDiagnosis 
         Height          =   1140
         ItemData        =   "frmUpdateConclusion.frx":0000
         Left            =   4440
         List            =   "frmUpdateConclusion.frx":0007
         TabIndex        =   8
         Top             =   480
         Width           =   3345
      End
      Begin VB.Label clblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƚ��롰�������á�������ѡ������ġ��Ƿ񸴲��������ԣ�"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2160
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   5760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϴ��������"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���п�ѡ�����"
         Height          =   180
         Index           =   1
         Left            =   4665
         TabIndex        =   10
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "�޸�������"
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   8055
      Begin VB.ListBox clstSelectedConclusion 
         Height          =   1320
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   3015
      End
      Begin VB.ListBox clstAllConclusion 
         Height          =   1320
         Left            =   4440
         TabIndex        =   13
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "<<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   765
         Width           =   1020
      End
      Begin VB.CommandButton ccmdDel 
         Caption         =   ">>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ۣ�"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���п�ѡ�����ۣ�"
         Height          =   180
         Index           =   0
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "frmUpdateConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��
'����޸ģ��

Private mstrϵͳ��� As String
Private mstr������ As String
Private mstr������ As String
Private mstr���������� As String

Private Sub clstAllConclusion_DblClick()
    On Error Resume Next
    If clstAllConclusion.ListIndex >= 0 Then
        ccmdAdd_Click 0
    End If
    
End Sub

Private Sub clstAllDiagnosis_Click()
    On Error Resume Next
    If clstAllDiagnosis.ListIndex >= 0 Then
        ccmdAdd(1).Enabled = True
    End If
End Sub

Private Sub clstAllDiagnosis_DblClick()
    On Error Resume Next
    If clstAllDiagnosis.ListIndex >= 0 Then
        ccmdAdd_Click 1
    End If
    
End Sub

Private Sub clstSelectedConclusion_DblClick()
    On Error Resume Next
    If clstSelectedConclusion.ListIndex >= 0 Then
        ccmdDel_Click
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    Dim lobj����ģ�弯 As Object 'clsMedicalExamTemplateSet��
    Dim lcolInfo As Collection     '��ȡ�ĸ����������ơ�
    Dim i As Long
    On Error GoTo errHandler
    '��������ģ�弯����
    Set lobj����ģ�弯 = CreateObject("������.clsMedicalExamTemplateSet")
    
    '��ȡ���и�����������
    lobj����ģ�弯.�������� = 2
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    
    '��ʾ�ڸ������������б���С�
    ccmbTemplate.Clear
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
    Next
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmUpdateConclusion", "clstSelectedConclusion_Click", 6666, lstrError, False
End Sub

Public Property Get ϵͳ���() As String
    ϵͳ��� = mstrϵͳ���
End Property

Public Property Let ϵͳ���(ByVal vNewValue As String)
    Dim lobj��� As Object         'clsMedicalExam��
    Dim lobj����ģ�� As Object   'clsMedicalExamTemplate��
    Dim lcolInfo As Collection     '����ģ��������ۼ����ԡ�
    Dim lstrAllDiagnosis As String '����ģ���������ԡ���ϴ����������
    Dim lstrItem As String
    Dim i As Long
    
    On Error GoTo errHandler
    If mstrϵͳ��� = vNewValue Then Exit Property
    
    '����������
    Set lobj��� = CreateObject("������.clsMedicalExam")
    lobj���.ϵͳ��� = vNewValue
    
    '��������ģ�����
    Set lobj����ģ�� = CreateObject("������.clsMedicalExamTemplate")
    lobj����ģ��.������ = lobj���.����.������
    
    '������ģ��������ۼ�������Ԫ����ʾ��clstAllConclusion�С�
    Set lcolInfo = lobj����ģ��.�����ۼ�
    clstAllConclusion.Clear
    For i = 1 To lcolInfo.Count
        clstAllConclusion.AddItem lcolInfo(i)("����")
    Next
    
    '������ģ���������ԡ���ϴ���������в�ֺ���ʾ��clstAllDiagnosis�С�
    lstrAllDiagnosis = lobj����ģ��.��ϴ������
    clstAllDiagnosis.Clear
    i = 1
    lstrItem = gffuncGetItemFromList(lstrAllDiagnosis, i, ",")
    Do While lstrItem <> ""
        clstAllDiagnosis.AddItem lstrItem
        i = i + 1
        lstrItem = gffuncGetItemFromList(lstrAllDiagnosis, i, ",")
    Loop
    
    mstrϵͳ��� = vNewValue

    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmUpdateConclusion", "Property Let ϵͳ���", 6666, lstrError, True
End Property
Public Property Get ��ϴ������() As String
    ��ϴ������ = mstr������
End Property

Public Property Let ��ϴ������(ByVal vNewValue As String)
    On Error Resume Next
    mstr������ = vNewValue
    
    '�Ѵ�������"��ϴ������"��ʾ�ڡ���ϴ��������¼����С�
    ctxtDiagnosis.Text = mstr������
    
End Property

Public Property Get ������() As String
    On Error Resume Next
    ������ = mstr������
End Property
Public Property Let ������(ByVal vNewValue As String)
    Dim lstrItem As String
    Dim i As Long
    
    On Error GoTo errHandler
    mstr������ = vNewValue
    '�Ѵ�������"������"��ֺ���ʾ�ڡ�ѡ�������ۡ��б��С�
    clstSelectedConclusion.Clear
    i = 1
    lstrItem = gffuncGetItemFromList(mstr������, i, ",")
    Do While lstrItem <> ""
        clstSelectedConclusion.AddItem lstrItem
        i = i + 1
        lstrItem = gffuncGetItemFromList(mstr������, i, ",")
    Loop
    
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmUpdateConclusion", "Property Let ������", 6666, lstrError, True
End Property

Public Property Let ����������(ByVal vNewValue As String)
    On Error GoTo errHandler

    mstr���������� = vNewValue
    
    '��"mstr����������"Ϊ�գ���ѡ��cchkTemplate������ccmbTemplate���ɼ�������ѡ��cchkTemplate������ccmbTemplate�ɼ���
    If mstr���������� = "" Then
        cchkTemplate.Value = 0
        ccmbTemplate.Visible = False
    Else
        cchkTemplate.Value = 1
        ccmbTemplate.Visible = True
        
        '�ø��������б�ѡ�е�ǰ���ԡ���������������
        ccmbTemplate.ListIndex = gffuncItemIsInComboBox(ccmbTemplate, mstr����������)
        
    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmUpdateConclusion", "Property Let ����������", 6666, lstrError, True
End Property

Public Property Get ����������() As String
    ���������� = mstr����������
End Property

Private Sub cchkTemplate_Click()
    On Error Resume Next
    If cchkTemplate.Value = 1 Then
        ccmbTemplate.Visible = True
        If ccmbTemplate.ListCount = 0 Then
            clblInfo.Visible = True
        End If
        ccmbTemplate.SetFocus
    Else
        ccmbTemplate.Visible = False
    End If
End Sub

Private Sub ccmdAdd_Click(Index As Integer)
    On Error GoTo errHandler
    If Index = 0 Then
        '��������ۡ�
        clstSelectedConclusion.AddItem clstAllConclusion.Text
        clstAllConclusion.RemoveItem clstAllConclusion.ListIndex
        ccmdAdd(0).Enabled = False
    Else
        If ctxtDiagnosis.Text <> "" Then
            ctxtDiagnosis.Text = ctxtDiagnosis.Text & IIf(Right(Trim(ctxtDiagnosis.Text), 1) = ",", "", ",") & clstAllDiagnosis.Text
        Else
            ctxtDiagnosis.Text = clstAllDiagnosis.Text
        End If
        ccmdClear.Enabled = True
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmUpdateConclusion", "ccmdAdd_Click", 6666, lstrError, False
End Sub

Private Sub ccmdCancel_Click()
    '���ش��塣
    Me.Hide
End Sub

Private Sub ccmdClear_Click()
    ctxtDiagnosis.Text = ""
    
End Sub

Private Sub ccmdDel_Click()
    On Error GoTo errHandler
    'ɾ�������ۡ�
    clstAllConclusion.AddItem clstSelectedConclusion.Text
    clstSelectedConclusion.RemoveItem clstSelectedConclusion.ListIndex
    ccmdDel.Enabled = False
   
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmUpdateConclusion", "ccmdDel_Click", 6666, lstrError, False
End Sub

'���ܣ�ȷ��������¼�������ô������ԣ��������ء�
Private Sub ccmdOk_Click()
    Dim i As Long
    
    On Error GoTo errHandler
    '��ѡ�е��������б���л�ȡ�����ۡ�
    mstr������ = ""
    For i = 0 To clstSelectedConclusion.ListCount - 1
        mstr������ = mstr������ & clstSelectedConclusion.List(i) & ","
    Next
    If mstr������ <> "" Then mstr������ = Left(mstr������, Len(mstr������) - 1)
    
    mstr������ = Trim(ctxtDiagnosis.Text)
    
    If cchkTemplate.Value = 1 Then
        mstr���������� = ccmbTemplate.Text
    Else
        mstr���������� = ""
    End If
    
    '���ش��塣
    Me.Hide
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmUpdateConclusion", "ccmdOK_Click", 6666, lstrError, False
End Sub

Private Sub clstAllConclusion_Click()
    On Error Resume Next
    If clstAllConclusion.ListIndex >= 0 Then
        ccmdAdd(0).Enabled = True
    Else
        ccmdAdd(0).Enabled = False
    End If
End Sub

Private Sub clstSelectedConclusion_Click()
    On Error Resume Next
    ccmdDel.Enabled = True
End Sub

Private Sub ccmbTemplate_GotFocus()
    On Error Resume Next
    If ccmbTemplate.Text = "" And ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.ListIndex = 0
    End If
End Sub


